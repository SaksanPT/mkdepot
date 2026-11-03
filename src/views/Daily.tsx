import { useEffect, useState, type ReactNode } from 'react'
import { supabase } from '../lib/supabase'

interface Ready { power_bi?: boolean; trips?: boolean; me2n?: boolean; wms?: boolean }
interface Sum {
  kpi?: { short?: number; avail?: number }
  alerts?: { transfer_pending?: number; depot_urgent?: number }
  last_run?: { lines?: number; exported?: boolean } | null
}

/* ไอคอนเส้นบาง ใช้ชุดเดียวกันทั้งหน้า */
const P = {
  upload: 'M12 15V4m0 0L8.5 7.5M12 4l3.5 3.5M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3',
  file: 'M6 3h7l5 5v13a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1zM13 3v5h5',
  calc: 'M8 8h8M8 12h3M8 16h3M15 12v4',
  doc: 'M5 3h14v18H5zM9 7h6M9 11h6M9 15h3',
  swap: 'M7 7h11l-3-3m3 13H7l3 3',
  pin: 'M12 21s7-6.2 7-11a7 7 0 1 0-14 0c0 4.8 7 11 7 11z',
  alert: 'M12 8v5m0 3.5h.01M10.3 3.9 2.6 17.1A1.6 1.6 0 0 0 4 19.5h16a1.6 1.6 0 0 0 1.4-2.4L13.7 3.9a1.6 1.6 0 0 0-2.8 0Z',
  search: 'M11 19a8 8 0 1 0 0-16 8 8 0 0 0 0 16zM21 21l-4.3-4.3',
  truck: 'M3 7h10v9H3zM13 10h4l3 3v3h-7z',
  clock: 'M12 7.5V12l3 2',
  box: 'M4 8.5 12 4l8 4.5v7L12 20l-8-4.5z M4 8.5 12 13l8-4.5M12 13v7',
  chart: 'M4 19V5m0 14h16M8 16v-5m4 5V7m4 9v-3',
  plus: 'M12 5v14M5 12h14',
  layers: 'M12 3 3 8l9 5 9-5zM3 14l9 5 9-5',
  gear: 'M12 2.5v2M12 19.5v2M4.2 4.2l1.4 1.4M18.4 18.4l1.4 1.4M2.5 12h2M19.5 12h2M4.2 19.8l1.4-1.4M18.4 5.6l1.4-1.4',
}

const Ico = ({ d, circle }: { d: string; circle?: boolean }) => (
  <svg viewBox="0 0 24 24" width="19" height="19" fill="none"
    stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round">
    {circle && <circle cx="12" cy="12" r="8.5" />}
    <path d={d} />
  </svg>
)

interface CardDef {
  tab: string; tone: string
  a: string; b: string          // ไอคอนคู่
  title: string; desc: string
  pill?: ReactNode; pillTone?: string
}

export default function Daily({ snapshotDate, setSnapshotDate, go }: {
  snapshotDate: string
  setSnapshotDate: (d: string) => void
  go: (tab: string) => void
}) {
  const [ready, setReady] = useState<Ready>({})
  const [sum, setSum] = useState<Sum | null>(null)
  const [ship, setShip] = useState<{ pending: number; orderDay: boolean }>({ pending: 0, orderDay: false })

  useEffect(() => { void check() }, [snapshotDate])
  useEffect(() => {
    supabase.rpc('dashboard_summary').then(({ data }) => setSum((data as Sum) ?? null))
    // Shipto: ใบที่รอออก PO และวันนี้เป็นวันสั่งของผู้ผลิตรายไหนหรือไม่
    void (async () => {
      const d = new Date().getDay()
      const dow = d === 0 ? 7 : d
      const [o, v] = await Promise.all([
        supabase.from('shipto_orders').select('id', { count: 'exact', head: true }).eq('status', 'รอจัดส่ง'),
        supabase.from('shipto_vendors').select('order_dows').eq('active', true),
      ])
      setShip({
        pending: o.count ?? 0,
        orderDay: ((v.data ?? []) as { order_dows: number[] }[]).some((r) => r.order_dows?.includes(dow)),
      })
    })()
  }, [])

  async function check() {
    const { data } = await supabase.from('import_batches')
      .select('source').eq('status', 'committed').eq('snapshot_date', snapshotDate)
    const m: Ready = {}
    ;(data ?? []).forEach((r) => { m[r.source as keyof Ready] = true })
    setReady(m)
  }

  const files = [
    { k: 'power_bi' as const, label: 'สต็อกรายสาขา', must: true },
    { k: 'trips' as const, label: 'เที่ยวรถ', must: true },
    { k: 'me2n' as const, label: 'ของระหว่างทาง', must: false },
    { k: 'wms' as const, label: 'สต็อกคลัง', must: false },
  ]
  const done = files.filter((f) => ready[f.k]).length
  const mustDone = files.filter((f) => f.must).every((f) => ready[f.k])

  const short = sum?.kpi?.short ?? 0
  const pending = sum?.alerts?.transfer_pending ?? 0
  const urgent = sum?.alerts?.depot_urgent ?? 0
  const run = sum?.last_run

  const main: CardDef[] = [
    {
      tab: 'import', tone: 'g-violet', a: P.upload, b: P.file,
      title: 'อัปไฟล์ประจำรอบ',
      desc: 'นำเข้าสต็อกรายสาขา เที่ยวรถ ของระหว่างทาง และสต็อกคลัง ระบบตรวจความครบให้อัตโนมัติ',
      pill: `${done}/4 ไฟล์`, pillTone: mustDone ? 'ok' : 'hot',
    },
    {
      tab: 'run', tone: 'g-blue', a: P.calc, b: P.doc,
      title: 'คำนวณยอดเติม',
      desc: 'คำนวณว่าแต่ละสาขาต้องเติมกี่ชิ้น ตรวจเหตุผลรายบรรทัดได้ แล้วออกไฟล์สั่งซื้อ',
      pill: run?.lines ? `${run.lines} บรรทัด` : 'ยังไม่คำนวณ',
      pillTone: run?.exported ? 'ok' : run?.lines ? 'warn' : undefined,
    },
    {
      tab: 'transfer', tone: 'g-green', a: P.swap, b: P.pin,
      title: 'โอนเกลี่ยสินค้า',
      desc: 'จับคู่สาขาที่ของล้นกับสาขาที่ขาดในพื้นที่ใกล้กัน ลดการเบิกใหม่จากคลัง',
      pill: pending > 0 ? `กำลังโอน ${pending}` : 'ไม่มีค้าง',
      pillTone: pending > 0 ? 'warn' : 'ok',
    },
    {
      tab: 'shortage', tone: 'g-rose', a: P.alert, b: P.search,
      title: 'ของขาด',
      desc: 'ดูทุกบรรทัดที่สต็อกเป็นศูนย์ พร้อมสาเหตุว่าติดที่คลัง ที่รถ หรือของกำลังมา',
      pill: short > 0 ? `${short} บรรทัด` : 'ไม่มีของขาด',
      pillTone: short > 0 ? 'hot' : 'ok',
    },
    {
      tab: 'shipto', tone: 'g-amber', a: P.box, b: P.truck,
      title: 'Shipto สั่งตรง',
      desc: 'สั่งน้ำมันเครื่องและหัวเชื้อตรงจากผู้ผลิตส่งถึงสถานี เชฟรอน · มิราเคิล สั่งทุกจันทร์ อังคาร',
      pill: ship.pending > 0 ? `รอออก PO ${ship.pending} ใบ`
        : ship.orderDay ? 'วันนี้วันสั่ง' : 'ไม่มีค้าง',
      pillTone: ship.pending > 0 ? 'warn' : ship.orderDay ? 'hot' : 'ok',
    },
    {
      tab: 'depot', tone: 'g-teal', a: P.box, b: P.chart,
      title: 'สั่งเข้าคลัง',
      desc: 'คำนวณจากยอดที่คลังจ่ายออกจริง บอกจำนวนที่ต้องสั่งเป็นลังเต็ม',
      pill: urgent > 0 ? `ด่วน ${urgent}` : 'ปกติ',
      pillTone: urgent > 0 ? 'hot' : 'ok',
    },
    {
      tab: 'manual', tone: 'g-violet', a: P.plus, b: P.swap,
      title: 'โอนกำหนดเอง',
      desc: 'เลือกสาขาต้นทางปลายทางและจำนวนเอง บันทึกลงระบบเดียวกับการโอนปกติ',
    },
    {
      tab: 'transferB', tone: 'g-green', a: P.layers, b: P.box,
      title: 'โอนข้ามคลาส',
      desc: 'เอาของเกินจากสินค้า Class A ไปเติมสาขาที่สินค้าเป็น Class B และของหมด',
    },
  ]

  const more: CardDef[] = [
    {
      tab: 'home', tone: 'g-blue', a: P.chart, b: P.search,
      title: 'ภาพรวมและกราฟ',
      desc: 'Availability DOH และยอดขายย้อนหลัง พร้อมรายการที่ต้องจัดการวันนี้',
    },
    {
      tab: 'kpi', tone: 'g-teal', a: P.chart, b: P.clock,
      title: 'KPI ย้อนหลัง',
      desc: 'ดูตัวเลขรายวันย้อนหลัง เทียบกับเป้าหมายที่ตั้งไว้',
    },
    {
      tab: 'groups', tone: 'g-rose', a: P.pin, b: P.layers,
      title: 'กลุ่มสถานี',
      desc: 'ตั้งว่าสาขาไหนฝากของไว้ที่ไหนได้ รอบที่รถไม่เข้าจะได้ของผ่านจุดฝาก',
    },
    {
      tab: 'settings', tone: 'g-amber', a: P.gear, b: P.calc,
      title: 'ตั้งค่าการคำนวณ',
      desc: 'ปรับจำนวนวันที่เผื่อ LeadTime Safety stock และเงื่อนไขการโอน',
    },
  ]

  const Card = (c: CardDef) => (
    <button key={c.tab + c.title} className={`gcard ${c.tone}`} onClick={() => go(c.tab)}>
      <span className="icons">
        <i><Ico d={c.a} /></i>
        <i className="alt"><Ico d={c.b} /></i>
      </span>
      <b>{c.title}</b>
      <p>{c.desc}</p>
      {c.pill && <span className={`pill ${c.pillTone ?? ''}`}>{c.pill}</span>}
    </button>
  )

  return (
    <div className="glass-bg">
      <div className="glass-head">
        <div>
          <h2>งานประจำวัน</h2>
          <div className="sub">
            {mustDone
              ? 'ไฟล์ที่จำเป็นครบแล้ว คำนวณได้เลย'
              : 'ยังขาดไฟล์ที่จำเป็น อัปให้ครบก่อนคำนวณ'}
          </div>
          <div className="chips">
            {files.map((f) => (
              <span key={f.k} className={`chip ${ready[f.k] ? 'good' : f.must ? 'bad' : ''}`}>
                {ready[f.k] ? '✓' : f.must ? '!' : '—'} {f.label}
              </span>
            ))}
          </div>
        </div>
        <div style={{ textAlign: 'right' }}>
          <div style={{ fontSize: 12.5, color: 'var(--ink-3)', marginBottom: 6 }}>
            ข้อมูล ณ วันที่
          </div>
          <input type="date" value={snapshotDate}
            onChange={(e) => setSnapshotDate(e.target.value)} />
        </div>
      </div>

      <div className="gcards">{main.map(Card)}</div>

      <div className="gsec">
        <h3>ดูข้อมูลและตั้งค่า</h3>
      </div>
      <div className="gcards">{more.map(Card)}</div>
    </div>
  )
}
