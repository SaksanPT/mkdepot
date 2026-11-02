import { Fragment, useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from '../lib/supabase'
import { Fold } from './ui'

/* ────────────────────────────────────────────────────────────────────
   Shipto — สั่งน้ำมันเครื่องตรงจากผู้ผลิตส่งถึงสถานี
   คำนวณ → ติ๊กเลือกสถานี → ยืนยัน (รอจัดส่ง) → ใส่เลข PO แล้วปิดใบ
   ใบที่ยังไม่ถึงถูกนับเป็นของกำลังมาในการคำนวณเติมปกติแล้ว ไม่สั่งซ้ำ
                                                                      */

interface Vendor {
  id: number; vendor_name: string; short_name: string
  min_litre: number | null; min_cases: number | null; full_case: boolean
  lead_min: number; lead_max: number; doh_max: number
  cycle_days: number; order_dows: number[]; active: boolean; sort: number
}

interface PLine {
  plant_code: string; branch_name: string; mat_code: string; item_name: string
  pack_size: number; litre_per_piece: number
  stock_pcs: number; incoming_pcs: number; spd: number; cover_days: number | null
  triggered: boolean; order_pcs: number; order_cases: number; order_litre: number
  topup_cases: number; line_doh: number | null
  station_litre: number; station_cases: number; station_doh: number | null
  first_out_days: number | null; over_doh: boolean; has_pending: boolean
  snapshot_date: string
}

interface Station {
  plant_code: string; branch_name: string; lines: PLine[]
  litre: number; cases: number; doh: number | null; firstOut: number | null
  over: boolean; pending: boolean
}

interface OLine {
  mat_code: string; qty_pcs: number; cases: number; litre: number
  doh_after: number | null; topup_cases: number
  items: { template_descr: string | null; desc_th: string | null } | null
}

interface Order {
  id: number; vendor_id: number; plant_code: string; status: string
  total_litre: number; total_cases: number; doh_after: number | null; over_doh: boolean
  confirmed_at: string; po_no: string | null; closed_at: string | null
  cancelled_at: string | null; snapshot_date: string | null
  shipto_lines: OLine[]
}

const DOW = ['', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']
const isoDow = (d: Date) => (d.getDay() === 0 ? 7 : d.getDay())
const n0 = (v: number | null | undefined) => (v ?? 0).toLocaleString('th-TH', { maximumFractionDigits: 0 })
const n1 = (v: number | null | undefined) =>
  v === null || v === undefined ? '—' : v.toLocaleString('th-TH', { maximumFractionDigits: 1 })
const daysAgo = (iso: string) => Math.floor((Date.now() - new Date(iso).getTime()) / 86400000)
const lineName = (l: OLine) => l.items?.template_descr ?? l.items?.desc_th ?? l.mat_code

type View = 'all' | 'ready' | 'over'

export default function Shipto() {
  const [vendors, setVendors] = useState<Vendor[]>([])
  const [vid, setVid] = useState<number | null>(null)
  const [names, setNames] = useState<Map<string, string>>(new Map())

  const [rows, setRows] = useState<PLine[] | null>(null)
  const [picked, setPicked] = useState<Set<string>>(new Set())
  const [open, setOpen] = useState<Set<string>>(new Set())
  const [view, setView] = useState<View>('all')
  const [q, setQ] = useState('')

  const [orders, setOrders] = useState<Order[]>([])
  const [history, setHistory] = useState<Order[]>([])
  const [oPicked, setOPicked] = useState<Set<number>>(new Set())
  const [po, setPo] = useState<Record<number, string>>({})
  const [askCancel, setAskCancel] = useState(false)

  const [busy, setBusy] = useState(false)
  const [err, setErr] = useState('')
  const [msg, setMsg] = useState('')

  const vendor = vendors.find((v) => v.id === vid) ?? null
  const today = isoDow(new Date())

  useEffect(() => {
    void (async () => {
      const [v, s] = await Promise.all([
        supabase.from('shipto_vendors').select('*').eq('active', true).order('sort'),
        supabase.from('stations').select('plant_code, branch_name'),
      ])
      if (v.error) { setErr(v.error.message); return }
      const list = (v.data ?? []) as unknown as Vendor[]
      setVendors(list)
      setVid((cur) => cur ?? list[0]?.id ?? null)
      setNames(new Map(((s.data ?? []) as { plant_code: string; branch_name: string }[])
        .map((r) => [r.plant_code, r.branch_name])))
    })()
  }, [])

  useEffect(() => {
    if (vid === null) return
    setRows(null); setPicked(new Set()); setOpen(new Set()); setOPicked(new Set())
    setErr(''); setMsg(''); setAskCancel(false)
    void loadOrders(vid)
  }, [vid])

  async function loadOrders(v: number) {
    const sel = 'id, vendor_id, plant_code, status, total_litre, total_cases, doh_after, over_doh, ' +
      'confirmed_at, po_no, closed_at, cancelled_at, snapshot_date, ' +
      'shipto_lines(mat_code, qty_pcs, cases, litre, doh_after, topup_cases, items(template_descr, desc_th))'
    const since = new Date(Date.now() - 30 * 86400000).toISOString()
    const [p, h] = await Promise.all([
      supabase.from('shipto_orders').select(sel).eq('vendor_id', v).eq('status', 'รอจัดส่ง')
        .order('confirmed_at'),
      supabase.from('shipto_orders').select(sel).eq('vendor_id', v).neq('status', 'รอจัดส่ง')
        .gte('confirmed_at', since).order('confirmed_at', { ascending: false }).limit(300),
    ])
    if (p.error) { setErr(p.error.message); return }
    setOrders((p.data ?? []) as unknown as Order[])
    setHistory((h.data ?? []) as unknown as Order[])
  }

  async function calc() {
    if (vid === null) return
    setBusy(true); setErr(''); setMsg('')
    const { data, error } = await supabase.rpc('shipto_preview', { p_vendor_id: vid })
    setBusy(false)
    if (error) { setErr(error.message); return }
    const list = (data ?? []) as unknown as PLine[]
    setRows(list)
    // ติ๊กให้เฉพาะสถานีที่ DOH ไม่เกิน และยังไม่มีใบรอจัดส่ง — ที่เกินให้ผู้ใช้ตัดสินใจเอง
    setPicked(new Set(list.filter((l) => !l.over_doh && !l.has_pending).map((l) => l.plant_code)))
    setOpen(new Set(list.filter((l) => l.over_doh).map((l) => l.plant_code)))
  }

  const stations = useMemo<Station[]>(() => {
    if (!rows) return []
    const m = new Map<string, Station>()
    rows.forEach((l) => {
      let s = m.get(l.plant_code)
      if (!s) {
        s = {
          plant_code: l.plant_code, branch_name: l.branch_name, lines: [],
          litre: l.station_litre, cases: l.station_cases, doh: l.station_doh,
          firstOut: l.first_out_days, over: l.over_doh, pending: l.has_pending,
        }
        m.set(l.plant_code, s)
      }
      s.lines.push(l)
    })
    return [...m.values()]
  }, [rows])

  const shown = useMemo(() => {
    const t = q.trim().toLowerCase()
    return stations.filter((s) => {
      if (view === 'ready' && (s.over || s.pending)) return false
      if (view === 'over' && !s.over) return false
      if (t && !`${s.branch_name} ${s.plant_code} ${s.lines.map((l) => l.item_name).join(' ')}`
        .toLowerCase().includes(t)) return false
      return true
    })
  }, [stations, view, q])

  const sum = useMemo(() => {
    const p = stations.filter((s) => picked.has(s.plant_code))
    return {
      total: stations.length,
      over: stations.filter((s) => s.over).length,
      picked: p.length,
      litre: p.reduce((a, s) => a + s.litre, 0),
      cases: p.reduce((a, s) => a + s.cases, 0),
    }
  }, [stations, picked])

  function toggle(set: Set<string>, k: string, fn: (s: Set<string>) => void) {
    const n = new Set(set); if (n.has(k)) n.delete(k); else n.add(k); fn(n)
  }

  async function confirm() {
    if (vid === null || picked.size === 0) return
    setBusy(true); setErr('')
    const { data, error } = await supabase.rpc('shipto_confirm', {
      p_vendor_id: vid, p_plants: [...picked],
    })
    setBusy(false)
    if (error) { setErr(error.message); return }
    const n = (data as number | null) ?? 0
    setMsg(`ยืนยันแล้ว ${n} สถานี — ขึ้นสถานะรอจัดส่ง และถูกนับเป็นของกำลังมาในการคำนวณเติมปกติแล้ว` +
      (n < picked.size ? ` (ข้าม ${picked.size - n} สถานีที่มีใบรอจัดส่งอยู่แล้ว)` : ''))
    setRows(null); setPicked(new Set())
    await loadOrders(vid)
  }

  async function closeOrders() {
    if (vid === null || oPicked.size === 0) return
    setBusy(true); setErr('')
    const items = [...oPicked].map((id) => ({ id, po_no: (po[id] ?? '').trim() }))
    const { data, error } = await supabase.rpc('shipto_close', { p_items: items })
    setBusy(false)
    if (error) { setErr(error.message); return }
    const noPo = items.filter((i) => !i.po_no).length
    setMsg(`ปิดใบแล้ว ${(data as number | null) ?? 0} ใบ` +
      (noPo ? ` · ${noPo} ใบไม่ได้ใส่เลข PO ระบบจะนับเป็นของกำลังมาต่ออีก ${vendor ? vendor.lead_max + 3 : 10} วันแล้วเลิกนับเอง` : ''))
    setOPicked(new Set())
    await loadOrders(vid)
  }

  async function cancelOrders() {
    if (vid === null || oPicked.size === 0) return
    setBusy(true); setErr(''); setAskCancel(false)
    const { data, error } = await supabase.rpc('shipto_cancel', { p_ids: [...oPicked] })
    setBusy(false)
    if (error) { setErr(error.message); return }
    setMsg(`ยกเลิกแล้ว ${(data as number | null) ?? 0} ใบ — เลิกนับเป็นของกำลังมา`)
    setOPicked(new Set())
    await loadOrders(vid)
  }

  async function reopen(id: number) {
    if (vid === null) return
    const { data, error } = await supabase.rpc('shipto_reopen', { p_id: id })
    if (error) { setErr(error.message); return }
    if (!data) setErr('เปิดใบนี้กลับไม่ได้ — สถานีนี้มีใบรอจัดส่งใบใหม่อยู่แล้ว')
    await loadOrders(vid)
  }

  /* ── ดาวน์โหลด ─────────────────────────────────────────────────── */
  function exportOrders() {
    const list = oPicked.size ? orders.filter((o) => oPicked.has(o.id)) : orders
    if (!list.length || !vendor) return
    const ws = XLSX.utils.json_to_sheet(list.flatMap((o) => o.shipto_lines.map((l) => ({
      'ผู้ผลิต': vendor.vendor_name,
      'รหัสสถานี': o.plant_code,
      'สถานี': names.get(o.plant_code) ?? o.plant_code,
      'รหัสสินค้า': l.mat_code,
      'สินค้า': lineName(l),
      'ลัง': l.cases,
      'ชิ้น': l.qty_pcs,
      'ลิตร': l.litre,
      'DOH หลังของถึง': l.doh_after,
      'วันที่ยืนยัน': o.confirmed_at.slice(0, 10),
    }))))
    ws['!cols'] = [{ wch: 30 }, { wch: 10 }, { wch: 30 }, { wch: 12 }, { wch: 40 },
      { wch: 6 }, { wch: 6 }, { wch: 7 }, { wch: 13 }, { wch: 12 }]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'รอจัดส่ง')
    XLSX.writeFile(wb, `Shipto_${vendor.short_name}_รอจัดส่ง_${new Date().toISOString().slice(0, 10).replace(/-/g, '')}.xlsx`)
  }

  function exportPreview() {
    if (!stations.length || !vendor) return
    const ws = XLSX.utils.json_to_sheet(stations.flatMap((s) => s.lines.map((l) => ({
      'เลือก': picked.has(s.plant_code) ? '✓' : '',
      'รหัสสถานี': s.plant_code,
      'สถานี': s.branch_name,
      'รหัสสินค้า': l.mat_code,
      'สินค้า': l.item_name,
      'คงเหลือ': l.stock_pcs,
      'กำลังมา': l.incoming_pcs,
      'ขาย/วัน': l.spd,
      'พอขาย(วัน)': l.cover_days,
      'สั่ง(ลัง)': l.order_cases,
      'สั่ง(ชิ้น)': l.order_pcs,
      'ลิตร': l.order_litre,
      'เติมให้ถึงขั้นต่ำ(ลัง)': l.topup_cases || '',
      'DOH หลังของถึง': l.line_doh,
      'DOH สถานี': s.doh,
      'เกิน DOH': s.over ? 'เกิน' : '',
    }))))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'ผลคำนวณ')
    XLSX.writeFile(wb, `Shipto_${vendor.short_name}_คำนวณ_${new Date().toISOString().slice(0, 10).replace(/-/g, '')}.xlsx`)
  }

  /* ── หน้าจอ ────────────────────────────────────────────────────── */
  const isOrderDay = vendor ? vendor.order_dows.includes(today) : false
  const needUntil = vendor ? vendor.cycle_days + vendor.lead_max : 0
  const minText = vendor
    ? [vendor.min_litre ? `${n0(vendor.min_litre)} ลิตร` : '', vendor.min_cases ? `${n0(vendor.min_cases)} ลัง` : '']
      .filter(Boolean).join(' และ ') || 'ไม่มี'
    : ''

  return (
    <>
      <h2>Shipto — สั่งตรงจากผู้ผลิต</h2>
      <p className="lede">
        น้ำมันเครื่องที่ผู้ผลิตส่งตรงถึงสถานี แยกตาม VENDOR NAME ใน Master Item ·
        ระบบเลือกสถานีที่ของจะหมดก่อนของรอบหน้ามาถึง ปัดเต็มลัง แล้วเติมให้ถึงขั้นต่ำ ·
        ยืนยันแล้วขึ้น <strong>รอจัดส่ง</strong> และถูกนับเป็นของกำลังมาในการคำนวณเติมปกติทันที
      </p>

      <div className="row" style={{ gap: 8, marginBottom: 14, flexWrap: 'wrap' }}>
        {vendors.map((v) => (
          <button key={v.id} className={`btn ${v.id === vid ? '' : 'ghost'}`} onClick={() => setVid(v.id)}>
            {v.short_name}
          </button>
        ))}
        {vendors.length === 0 && !err && <span className="hint">กำลังโหลด…</span>}
      </div>

      {vendor && (
        <div className="card">
          <div className="spread" style={{ alignItems: 'flex-start', gap: 12, flexWrap: 'wrap' }}>
            <div>
              <h3 style={{ marginBottom: 6 }}>{vendor.vendor_name}</h3>
              <div className="row" style={{ gap: 6, flexWrap: 'wrap' }}>
                <span className="tag">ขั้นต่ำต่อสถานี {minText}</span>
                {vendor.full_case && <span className="tag">เต็มลัง ไม่แตกลัง</span>}
                <span className="tag">ส่งของ {vendor.lead_min}–{vendor.lead_max} วัน</span>
                <span className="tag">DOH ไม่เกิน {n0(vendor.doh_max)} วัน</span>
                <span className="tag oil">ต้องพอขาย {needUntil} วัน</span>
              </div>
              <p className="hint" style={{ margin: '8px 0 0' }}>
                {needUntil} วัน = รอบสั่งถัดไปอีก {vendor.cycle_days} วัน + ของรอบนั้นส่งช้าสุด {vendor.lead_max} วัน
              </p>
            </div>
            <div style={{ textAlign: 'right' }}>
              <span className={`tag ${isOrderDay ? 'ok' : 'alarm'}`}>
                วันนี้วัน{DOW[today]}{isOrderDay ? ' · วันสั่ง' : ''}
              </span>
              <div className="hint" style={{ marginTop: 6 }}>
                วันสั่ง: {vendor.order_dows.map((d) => DOW[d]).join(' · ')}
              </div>
            </div>
          </div>
        </div>
      )}

      {err && <div className="note bad" style={{ marginBottom: 12 }}>{err}</div>}
      {msg && <div className="note" style={{ marginBottom: 12 }}>{msg}</div>}

      {/* ── ใบรอจัดส่ง ── */}
      {vendor && (
        <div className="card">
          <div className="spread" style={{ flexWrap: 'wrap', gap: 10 }}>
            <div>
              <h3>รอจัดส่ง <small style={{ color: 'var(--ink-3)' }}>{orders.length} ใบ</small></h3>
              <p className="hint" style={{ marginBottom: 0 }}>
                ออก PO แล้วให้ใส่เลข PO ติ๊ก แล้วกดปิดใบ · ใบที่ยังไม่ปิดเกิน 21 วันจะเลิกนับเป็นของกำลังมา
              </p>
            </div>
            {orders.length > 0 && (
              <div className="row" style={{ gap: 8 }}>
                <button className="btn ghost" onClick={exportOrders}>
                  ดาวน์โหลด{oPicked.size ? ` ${oPicked.size} ใบ` : 'ทั้งหมด'}
                </button>
                {oPicked.size > 0 && (askCancel ? (
                  <>
                    <button className="btn ghost" onClick={() => setAskCancel(false)}>ไม่ยกเลิก</button>
                    <button className="btn" style={{ background: 'var(--alarm)' }} onClick={cancelOrders} disabled={busy}>
                      ยืนยันยกเลิก {oPicked.size} ใบ
                    </button>
                  </>
                ) : (
                  <>
                    <button className="btn ghost" onClick={() => setAskCancel(true)}>ยกเลิก {oPicked.size} ใบ</button>
                    <button className="btn" onClick={closeOrders} disabled={busy}>
                      ออก PO แล้ว ปิด {oPicked.size} ใบ
                    </button>
                  </>
                ))}
              </div>
            )}
          </div>

          {orders.length === 0 ? (
            <div className="note" style={{ marginTop: 12 }}>ไม่มีใบรอจัดส่ง</div>
          ) : (
            <div className="tw" style={{ maxHeight: '40vh', marginTop: 14 }}>
              <table>
                <thead>
                  <tr>
                    <th style={{ width: 34 }}>
                      <input type="checkbox"
                        checked={oPicked.size === orders.length}
                        onChange={() => setOPicked(oPicked.size === orders.length
                          ? new Set() : new Set(orders.map((o) => o.id)))} />
                    </th>
                    <th>สถานี</th><th>สินค้า</th>
                    <th className="num">ลัง</th><th className="num">ลิตร</th>
                    <th className="num">DOH</th><th className="num">ยืนยันเมื่อ</th>
                    <th style={{ width: 150 }}>เลข PO</th>
                  </tr>
                </thead>
                <tbody>
                  {orders.map((o) => {
                    const age = daysAgo(o.confirmed_at)
                    return (
                      <tr key={o.id}>
                        <td>
                          <input type="checkbox" checked={oPicked.has(o.id)}
                            onChange={() => {
                              const n = new Set(oPicked)
                              if (n.has(o.id)) n.delete(o.id); else n.add(o.id)
                              setOPicked(n)
                            }} />
                        </td>
                        <td>
                          {names.get(o.plant_code) ?? o.plant_code}
                          <div className="hint" style={{ margin: 0 }}>{o.plant_code}</div>
                        </td>
                        <td style={{ fontSize: 12.5, color: 'var(--ink-2)' }}>
                          {o.shipto_lines.map((l) => `${lineName(l)} ×${n0(l.cases)}`).join(' · ')}
                        </td>
                        <td className="num"><strong>{n0(o.total_cases)}</strong></td>
                        <td className="num">{n0(o.total_litre)}</td>
                        <td className="num" style={{ color: o.over_doh ? 'var(--alarm)' : undefined }}>
                          {n1(o.doh_after)}
                        </td>
                        <td className="num" style={{ color: age > 10 ? 'var(--alarm)' : 'var(--ink-3)' }}>
                          {age === 0 ? 'วันนี้' : `${age} วัน`}
                        </td>
                        <td>
                          <input value={po[o.id] ?? ''} placeholder="4500…"
                            style={{ width: '100%' }}
                            onChange={(e) => {
                              setPo({ ...po, [o.id]: e.target.value })
                              if (e.target.value && !oPicked.has(o.id)) setOPicked(new Set(oPicked).add(o.id))
                            }} />
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* ── คำนวณ ── */}
      {vendor && (
        <div className="card">
          <div className="row" style={{ flexWrap: 'wrap', gap: 10 }}>
            <button className="btn" onClick={calc} disabled={busy}>
              {busy && !rows ? 'กำลังคำนวณ…' : `คำนวณ ${vendor.short_name}`}
            </button>
            {rows && rows[0] && (
              <span style={{ fontSize: 12.5, color: 'var(--ink-3)' }}>ใช้สต็อก ณ {rows[0].snapshot_date}</span>
            )}
            {stations.length > 0 && (
              <>
                <span style={{ flex: 1 }} />
                <button className="btn ghost" onClick={exportPreview}>ดาวน์โหลดผลคำนวณ</button>
                <button className="btn" onClick={confirm} disabled={busy || picked.size === 0}>
                  ยืนยัน {picked.size} สถานี → รอจัดส่ง
                </button>
              </>
            )}
          </div>
          {!isOrderDay && rows === null && (
            <p className="hint" style={{ margin: '10px 0 0' }}>
              วันนี้ไม่ใช่วันสั่ง — คำนวณดูได้ แต่ตัวเลข "ต้องพอขาย {needUntil} วัน" ตั้งไว้สำหรับการสั่งวันจันทร์–อังคาร
            </p>
          )}
          {rows && stations.length === 0 && (
            <div className="note" style={{ marginTop: 12 }}>
              ไม่มีสถานีที่ของจะหมดก่อนของรอบหน้ามาถึง — ยังไม่ต้องสั่ง {vendor.short_name} รอบนี้
            </div>
          )}
        </div>
      )}

      {stations.length > 0 && vendor && (
        <>
          <dl className="stats">
            <div className="stat"><dt>สถานีที่ต้องสั่ง</dt><dd>{sum.total}</dd></div>
            <div className="stat"><dt>เลือกไว้</dt><dd>{sum.picked}</dd></div>
            <div className="stat"><dt>รวม</dt><dd>{n0(sum.cases)} <small>ลัง</small></dd></div>
            <div className="stat"><dt>รวม</dt><dd>{n0(sum.litre)} <small>ลิตร</small></dd></div>
          </dl>

          {sum.over > 0 && (
            <div className="note bad" style={{ marginBottom: 12 }}>
              {sum.over} สถานีสั่งถึงขั้นต่ำแล้ว DOH หลังของถึงเกิน {n0(vendor.doh_max)} วัน —
              แสดงไว้แต่ไม่ได้ติ๊กให้ ดูรายละเอียดใต้ชื่อสถานี แล้วติ๊กเองถ้าจะสั่ง
            </div>
          )}

          <div className="row" style={{ gap: 8, marginBottom: 10, flexWrap: 'wrap' }}>
            {([['all', `ทั้งหมด ${sum.total}`], ['ready', 'พร้อมสั่ง'], ['over', `DOH เกิน ${sum.over}`]] as [View, string][])
              .map(([k, label]) => (
                <button key={k} className={`btn ${view === k ? '' : 'ghost'}`} onClick={() => setView(k)}>{label}</button>
              ))}
            <input placeholder="ค้นหาสถานี / สินค้า" value={q} onChange={(e) => setQ(e.target.value)}
              style={{ flex: 1, minWidth: 180 }} />
          </div>

          <div className="tw" style={{ marginBottom: 20 }}>
            <table>
              <thead>
                <tr>
                  <th style={{ width: 34 }}>
                    <input type="checkbox"
                      checked={shown.length > 0 && shown.filter((s) => !s.pending).every((s) => picked.has(s.plant_code))}
                      onChange={(e) => {
                        const n = new Set(picked)
                        shown.filter((s) => !s.pending).forEach((s) => {
                          if (e.target.checked) n.add(s.plant_code); else n.delete(s.plant_code)
                        })
                        setPicked(n)
                      }} />
                  </th>
                  <th>สถานี</th>
                  <th className="num">หมดใน</th>
                  <th className="num">ลัง</th>
                  <th className="num">ลิตร</th>
                  <th className="num">DOH หลังของถึง</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {shown.map((s) => {
                  const isOpen = open.has(s.plant_code)
                  return (
                    <Fragment key={s.plant_code}>
                      <tr style={{ opacity: s.pending ? 0.5 : 1 }}>
                        <td>
                          <input type="checkbox" disabled={s.pending} checked={picked.has(s.plant_code)}
                            onChange={() => toggle(picked, s.plant_code, setPicked)} />
                        </td>
                        <td>
                          <span style={{ cursor: 'pointer' }} onClick={() => toggle(open, s.plant_code, setOpen)}>
                            {isOpen ? '▾' : '▸'} {s.branch_name}
                          </span>
                          {s.over && (
                            <div style={{ fontSize: 12, color: 'var(--alarm)', marginTop: 2 }}>
                              สั่ง {s.lines.map((l) => `${l.item_name} ${n0(l.order_cases)} ลัง`).join(', ')}
                              {' '}→ DOH {n1(s.doh)} วัน (เกิน {n0(vendor.doh_max)})
                            </div>
                          )}
                        </td>
                        <td className="num" style={{ color: (s.firstOut ?? 99) <= vendor.lead_max ? 'var(--alarm)' : undefined }}>
                          {s.firstOut === null ? '—' : s.firstOut <= 0 ? 'หมดแล้ว' : `${n1(s.firstOut)} วัน`}
                        </td>
                        <td className="num"><strong>{n0(s.cases)}</strong></td>
                        <td className="num">{n0(s.litre)}</td>
                        <td className="num" style={{ color: s.over ? 'var(--alarm)' : undefined }}>
                          <strong>{n1(s.doh)}</strong>
                        </td>
                        <td>
                          {s.pending ? <span className="tag">มีใบรอจัดส่งแล้ว</span>
                            : s.over ? <span className="tag alarm">DOH เกิน</span>
                            : <span className="tag ok">พร้อมสั่ง</span>}
                        </td>
                      </tr>
                      {isOpen && (
                        <tr>
                          <td></td>
                          <td colSpan={6} style={{ background: 'var(--paper-2, transparent)', paddingTop: 0 }}>
                            <table style={{ fontSize: 12.5 }}>
                              <thead>
                                <tr>
                                  <th>สินค้า</th>
                                  <th className="num">คงเหลือ</th><th className="num">กำลังมา</th>
                                  <th className="num">ขาย/วัน</th><th className="num">พอขาย</th>
                                  <th className="num">สั่ง (ลัง)</th><th className="num">ชิ้น</th>
                                  <th className="num">ลิตร</th><th className="num">DOH หลังของถึง</th><th></th>
                                </tr>
                              </thead>
                              <tbody>
                                {s.lines.map((l) => (
                                  <tr key={l.mat_code}>
                                    <td>{l.item_name}</td>
                                    <td className="num">{n0(l.stock_pcs)}</td>
                                    <td className="num">{l.incoming_pcs ? n0(l.incoming_pcs) : '—'}</td>
                                    <td className="num">{n1(l.spd)}</td>
                                    <td className="num" style={{ color: l.triggered ? 'var(--alarm)' : undefined }}>
                                      {l.cover_days === null ? '—' : `${n1(l.cover_days)} วัน`}
                                    </td>
                                    <td className="num"><strong>{n0(l.order_cases)}</strong></td>
                                    <td className="num">{n0(l.order_pcs)}</td>
                                    <td className="num">{n1(l.order_litre)}</td>
                                    <td className="num" style={{ color: (l.line_doh ?? 0) > vendor.doh_max ? 'var(--alarm)' : undefined }}>
                                      {n1(l.line_doh)}
                                    </td>
                                    <td>
                                      {l.topup_cases > 0 && (
                                        <span className="tag" title="เพิ่มเพื่อให้ถึงขั้นต่ำของผู้ผลิต">
                                          {l.triggered ? `+${l.topup_cases} ลังให้ถึงขั้นต่ำ` : 'เติมให้ถึงขั้นต่ำ'}
                                        </span>
                                      )}
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </td>
                        </tr>
                      )}
                    </Fragment>
                  )
                })}
              </tbody>
            </table>
          </div>
        </>
      )}

      {vendor && (
        <Fold title="ใบที่ปิด / ยกเลิกแล้ว" note={`30 วันล่าสุด ${history.length} ใบ`}
          hint="ใบที่ปิดแล้วยังนับเป็นของกำลังมาจนกว่าเลข PO จะเข้า ME2N หรือพ้นวันส่งช้าสุด + 3 วัน">
          {history.length === 0 ? (
            <div className="note">ยังไม่มี</div>
          ) : (
            <div className="tw" style={{ maxHeight: '40vh' }}>
              <table>
                <thead>
                  <tr>
                    <th>สถานี</th><th>สถานะ</th><th>เลข PO</th>
                    <th className="num">ลัง</th><th className="num">ลิตร</th>
                    <th className="num">ยืนยัน</th><th className="num">ปิด/ยกเลิก</th><th></th>
                  </tr>
                </thead>
                <tbody>
                  {history.map((o) => (
                    <tr key={o.id} style={{ opacity: o.status === 'ยกเลิก' ? 0.5 : 1 }}>
                      <td>{names.get(o.plant_code) ?? o.plant_code}</td>
                      <td><span className={`tag ${o.status === 'ปิดแล้ว' ? 'ok' : ''}`}>{o.status}</span></td>
                      <td>{o.po_no ?? '—'}</td>
                      <td className="num">{n0(o.total_cases)}</td>
                      <td className="num">{n0(o.total_litre)}</td>
                      <td className="num">{o.confirmed_at.slice(0, 10)}</td>
                      <td className="num">{(o.closed_at ?? o.cancelled_at ?? '').slice(0, 10)}</td>
                      <td>
                        {o.status === 'ปิดแล้ว' && (
                          <button className="btn ghost" style={{ padding: '2px 10px' }} onClick={() => reopen(o.id)}>
                            เปิดกลับ
                          </button>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </Fold>
      )}

      {vendor && <RulesEditor vendor={vendor} onSaved={(v) => setVendors(vendors.map((x) => (x.id === v.id ? v : x)))} />}
    </>
  )
}

/** แก้เงื่อนไขผู้ผลิต — บันทึกได้เฉพาะ admin (RLS กันไว้ฝั่งฐานข้อมูล) */
function RulesEditor({ vendor, onSaved }: { vendor: Vendor; onSaved: (v: Vendor) => void }) {
  const [f, setF] = useState(vendor)
  const [note, setNote] = useState('')
  useEffect(() => { setF(vendor); setNote('') }, [vendor])

  const numOrNull = (s: string) => (s.trim() === '' ? null : Number(s))
  const field = (label: string, key: keyof Vendor, nullable = false) => (
    <label style={{ display: 'grid', gap: 4, fontSize: 12.5 }}>
      <span style={{ color: 'var(--ink-3)' }}>{label}</span>
      <input type="number" min={0} value={(f[key] as number | null) ?? ''}
        onChange={(e) => setF({ ...f, [key]: nullable ? numOrNull(e.target.value) : Number(e.target.value) })} />
    </label>
  )

  async function save() {
    setNote('')
    if (f.lead_max < f.lead_min) { setNote('ส่งช้าสุดต้องไม่น้อยกว่าส่งเร็วสุด'); return }
    const { data, error } = await supabase.from('shipto_vendors').update({
      min_litre: f.min_litre, min_cases: f.min_cases, full_case: f.full_case,
      lead_min: f.lead_min, lead_max: f.lead_max, doh_max: f.doh_max,
      cycle_days: f.cycle_days, order_dows: f.order_dows,
    }).eq('id', f.id).select()
    if (error) { setNote(error.message); return }
    if (!data || data.length === 0) { setNote('บันทึกไม่ได้ — ต้องเป็นผู้ดูแลระบบ (admin)'); return }
    onSaved(data[0] as unknown as Vendor)
    setNote('บันทึกแล้ว — กดคำนวณใหม่เพื่อใช้เงื่อนไขนี้')
  }

  return (
    <Fold title="เงื่อนไขผู้ผลิต" note={vendor.short_name} hint="แก้ได้เฉพาะผู้ดูแลระบบ">
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(150px, 1fr))', gap: 12 }}>
        {field('ขั้นต่ำ (ลิตร) เว้นว่าง = ไม่มี', 'min_litre', true)}
        {field('ขั้นต่ำ (ลัง) เว้นว่าง = ไม่มี', 'min_cases', true)}
        {field('ส่งเร็วสุด (วัน)', 'lead_min')}
        {field('ส่งช้าสุด (วัน)', 'lead_max')}
        {field('DOH สูงสุด (วัน)', 'doh_max')}
        {field('ระยะห่างรอบสั่ง (วัน)', 'cycle_days')}
      </div>
      <div className="row" style={{ gap: 14, marginTop: 12, flexWrap: 'wrap' }}>
        <label className="row" style={{ gap: 6, fontSize: 13 }}>
          <input type="checkbox" checked={f.full_case} onChange={(e) => setF({ ...f, full_case: e.target.checked })} />
          สั่งเต็มลัง ไม่แตกลัง
        </label>
        <span style={{ fontSize: 13, color: 'var(--ink-3)' }}>วันสั่ง:</span>
        {[1, 2, 3, 4, 5, 6, 7].map((d) => (
          <label key={d} className="row" style={{ gap: 4, fontSize: 13 }}>
            <input type="checkbox" checked={f.order_dows.includes(d)}
              onChange={(e) => setF({
                ...f,
                order_dows: e.target.checked
                  ? [...f.order_dows, d].sort((a, b) => a - b)
                  : f.order_dows.filter((x) => x !== d),
              })} />
            {DOW[d]}
          </label>
        ))}
      </div>
      <div className="row" style={{ gap: 10, marginTop: 12 }}>
        <button className="btn" onClick={save}>บันทึกเงื่อนไข</button>
        {note && <span style={{ fontSize: 12.5, color: 'var(--ink-3)' }}>{note}</span>}
      </div>
    </Fold>
  )
}
