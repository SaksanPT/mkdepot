import * as XLSX from 'xlsx'

/* ────────────────────────────────────────────────────────────────────
   ตัวช่วย
   ไฟล์จริงมีหัวตารางไม่ตรงกันเล็กน้อย เช่น
     "ยอดขาย(ลิตร) เฉลี่ย 7 วัน"  กับ  "ยอดขาย(ลิตร)เฉลี่ย 7 วัน"
   จึงตัดช่องว่างทั้งหมดก่อนเทียบ                                       */

const norm = (s: unknown) => String(s ?? '').replace(/\s+/g, '').toLowerCase()

const num = (v: unknown): number => {
  if (v === null || v === undefined || v === '') return 0
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''))
  return Number.isFinite(n) ? n : 0
}

export type Grid = unknown[][]

export async function readSheet(file: File, sheetName?: string): Promise<{ grid: Grid; sheets: string[] }> {
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type: 'array', cellDates: true })
  const name = sheetName && wb.SheetNames.includes(sheetName) ? sheetName : wb.SheetNames[0]
  const grid = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[name], { header: 1, defval: null, raw: true })
  return { grid, sheets: wb.SheetNames }
}

/** หาแถวหัวตาราง โดยดูว่าแถวไหนมีคำเหล่านี้ครบ (ไฟล์ export มักมีแถวชื่อรายงาน/ตัวกรองนำหน้า) */
function findHeader(grid: Grid, must: string[], limit = 12): number {
  const want = must.map(norm)
  for (let r = 0; r < Math.min(grid.length, limit); r++) {
    const cells = (grid[r] || []).map(norm)
    if (want.every((w) => cells.some((c) => c.includes(w)))) return r
  }
  throw new Error(
    `หาหัวตารางไม่เจอ — ต้องมีคอลัมน์ ${must.join(', ')}\n` +
      `หัวตารางที่เจอในไฟล์: ${(grid[0] || []).filter(Boolean).slice(0, 12).join(' | ')}`
  )
}

function indexer(header: unknown[]) {
  const cols = header.map(norm)
  return (...alts: string[]): number => {
    for (const a of alts) {
      const k = norm(a)
      const i = cols.findIndex((c) => c === k)
      if (i >= 0) return i
    }
    for (const a of alts) {
      const k = norm(a)
      const i = cols.findIndex((c) => c.includes(k))
      if (i >= 0) return i
    }
    return -1
  }
}

const need = (i: number, label: string) => {
  if (i < 0) throw new Error(`ไม่พบคอลัมน์ "${label}" ในไฟล์`)
  return i
}

export interface ParseResult<T> {
  rows: T[]
  skipped: number
  warnings: string[]
}

/* ────────────────────────────────────────────────────────────────────
   1. POWER_BI / STATION DOH BY BRANCH
   ให้ทั้งสต็อกรายสาขา และข้อมูลสาขา (upsert stations ไปด้วยเลย)      */

export interface StockRow {
  plant_code: string; mat_code: string
  /** คลาสของสินค้าตัวนี้ที่สาขานี้ — สาขาเดียวมีได้หลายคลาส */
  class_fix: string | null; class_dyna: string | null; depot_class: string | null
  stock_l: number; stock_pcs: number
  /** ระหว่างทางตามที่ระบบสาขาเห็น ใช้เทียบกับ ME2N หาของที่ยังไม่ได้ทำรับ */
  transit_l: number; transit_pcs: number
  sales_7_l: number; sales_30_l: number; sales_90_l: number
  sales_7_pcs: number; sales_30_pcs: number; sales_90_pcs: number
}
export interface StationRow {
  plant_code: string; branch_name: string; depot: string
  class_fix: string | null; class_dyna: string | null
}

export function parsePowerBI(grid: Grid) {
  const h = findHeader(grid, ['PlantCode', 'รหัส (MatCode)'])
  const c = indexer(grid[h])
  const i = {
    plant: need(c('PlantCode'), 'PlantCode'),
    branch: need(c('สาขา'), 'สาขา'),
    clsFix: c('Class-สาขา(3ด.Fix)'),
    clsDyn: c('Class-สาขา(3ด.Dyna)'),
    mat: need(c('รหัส (MatCode)', 'MatCode'), 'รหัส (MatCode)'),
    depot: c('คลัง'),
    depotCls: c('Class-คลัง(3ด.Fix)'),
    stkL: need(c('คงเหลือ(ลิตร)'), 'คงเหลือ(ลิตร)'),
    stkP: need(c('คงเหลือ(ชิ้น)'), 'คงเหลือ(ชิ้น)'),
    trL: c('ระหว่างทาง(ลิตร)'), trP: c('ระหว่างทาง(ชิ้น)'),
    s7L: c('ยอดขาย(ลิตร)เฉลี่ย 7 วัน'), s30L: c('ยอดขาย(ลิตร)เฉลี่ย 30 วัน'), s90L: c('ยอดขาย(ลิตร)เฉลี่ย 90 วัน'),
    s7P: c('ยอดขาย(ชิ้น)เฉลี่ย 7 วัน'), s30P: c('ยอดขาย(ชิ้น)เฉลี่ย 30 วัน'), s90P: c('ยอดขาย(ชิ้น)เฉลี่ย 90 วัน'),
  }

  const stock = new Map<string, StockRow>()
  const stations = new Map<string, StationRow>()
  let skipped = 0

  for (let r = h + 1; r < grid.length; r++) {
    const row = grid[r]; if (!row) continue
    const plant = String(row[i.plant] ?? '').trim()
    const mat = String(row[i.mat] ?? '').trim().replace(/\.0$/, '')
    if (!plant || !mat) { skipped++; continue }

    /** จัดรูปแบบคลาสให้เหมือนกันทั้งสองช่อง จะได้เทียบกันได้ตรง ๆ */
    const norm = (v: unknown) => {
      const t = String(v ?? '').trim()
      return /^Class [ABC]$/i.test(t) ? 'Class ' + t.slice(-1).toUpperCase() : null
    }
    const clean = norm(row[i.clsFix])
    const cleanDyn = norm(row[i.clsDyn])

    stock.set(`${plant}|${mat}`, {
      plant_code: plant, mat_code: mat,
      class_fix: clean,
      class_dyna: cleanDyn,
      depot_class: i.depotCls >= 0 ? String(row[i.depotCls] ?? '').trim() || null : null,
      stock_l: num(row[i.stkL]), stock_pcs: Math.round(num(row[i.stkP])),
      transit_l: i.trL >= 0 ? num(row[i.trL]) : 0,
      transit_pcs: i.trP >= 0 ? Math.round(num(row[i.trP])) : 0,
      sales_7_l: num(row[i.s7L]), sales_30_l: num(row[i.s30L]), sales_90_l: num(row[i.s90L]),
      sales_7_pcs: num(row[i.s7P]), sales_30_pcs: num(row[i.s30P]), sales_90_pcs: num(row[i.s90P]),
    })

    // คลาสที่เก็บกับสาขาเป็นเพียงค่าอ้างอิง — สูตรใช้คลาสรายบรรทัดจาก stock_snapshots
    if (!stations.has(plant)) {
      stations.set(plant, {
        plant_code: plant,
        branch_name: String(row[i.branch] ?? '').trim() || plant,
        depot: String(row[i.depot] ?? 'แม่กลอง').trim() || 'แม่กลอง',
        class_fix: null,
        class_dyna: null,
      })
    }
  }

  const warnings: string[] = []
  if (i.s30L < 0) warnings.push('ไม่พบคอลัมน์ยอดขายลิตรเฉลี่ย 30 วัน — สูตรคำนวณจะได้ 0')

  const rows = [...stock.values()]
  const byClass = rows.reduce<Record<string, number>>((a, r) => {
    const k = r.class_fix ?? 'ไม่ระบุ'; a[k] = (a[k] ?? 0) + 1; return a
  }, {})
  const noClass = rows.filter((r) => !r.class_fix).length
  if (noClass) warnings.push(`${noClass} แถวไม่มีคลาสสินค้า — จะไม่ถูกนำมาคำนวณ`)

  return { stock: rows, stations: [...stations.values()], skipped, warnings, byClass }
}

/* 2. Master Item — ลิตรต่อชิ้น จำเป็นต่อสูตร ต้องนำเข้าก่อนเสมอ */

export interface ItemRow {
  mat_code: string; desc_en: string | null; desc_th: string | null
  litre_per_piece: number; pack_size: number | null
  /** BT / GAL — จากคอลัมน์ UOM ไม่ใช่ 'หน่วยโอน' ซึ่งเป็นจำนวนต่อลัง */
  uom: string | null
  /** จำนวนชิ้นต่อลัง จากคอลัมน์ 'หน่วยโอน' — >1 คือต้องสั่งยกลัง */
  units_per_case: number
  is_booster: boolean
  template_descr: string | null
  /** VENDOR NAME — ใช้แยกการสั่ง Shipto ตามผู้ผลิต (เชฟรอน / มิราเคิล) */
  vendor_name: string | null
  vendor_code: string | null
  /** Status DC — ACT / DEL / New · DEL ไม่ถูกนำไปสั่ง Shipto */
  status_dc: string | null
  /** OIL TYPE — BENZENE / DIESEL */
  oil_type: string | null
}

export function parseMasterItem(grid: Grid): ParseResult<ItemRow> {
  const h = findHeader(grid, ['MATERIAL CODE', 'LITRE'])
  const c = indexer(grid[h])
  const i = {
    mat: need(c('MATERIAL CODE'), 'MATERIAL CODE'),
    en: c('DESC'), th: c('DESC (TH)'),
    litre: need(c('LITRE'), 'LITRE'),
    pack: c('PACK SIZE'), uom: c('UOM'), perCase: c('หน่วยโอน'),
    vName: c('VENDOR NAME'), vCode: c('VENDOR CODE'),
    statusDc: c('Status DC'), oilType: c('OIL TYPE'),
  }
  const txt = (row: unknown[], k: number) =>
    k >= 0 ? String(row[k] ?? '').trim().replace(/\.0$/, '') || null : null
  const out = new Map<string, ItemRow>()
  const warnings: string[] = []
  let skipped = 0

  for (let r = h + 1; r < grid.length; r++) {
    const row = grid[r]; if (!row) continue
    const mat = String(row[i.mat] ?? '').trim().replace(/\.0$/, '')
    if (!/^\d{6,}$/.test(mat)) { skipped++; continue }
    const litre = num(row[i.litre])
    if (litre <= 0) { warnings.push(`SKU ${mat} ไม่มีค่า LITRE — ข้ามไป เพราะสูตรหารด้วยค่านี้`); skipped++; continue }
    const en = String(row[i.en] ?? '').trim() || null
    const th = String(row[i.th] ?? '').trim() || null
    const perCase = i.perCase >= 0 ? num(row[i.perCase]) : 1

    out.set(mat, {
      mat_code: mat,
      desc_en: en,
      desc_th: th,
      litre_per_piece: litre,
      pack_size: i.pack >= 0 ? num(row[i.pack]) || null : null,
      uom: i.uom >= 0 ? String(row[i.uom] ?? '').trim() || null : null,
      units_per_case: perCase > 0 ? perCase : 1,
      is_booster: /booster|หัวเชื้อ/i.test(`${en ?? ''} ${th ?? ''}`),
      template_descr: th ?? en,
      vendor_name: txt(row, i.vName),
      vendor_code: txt(row, i.vCode),
      status_dc: txt(row, i.statusDc)?.toUpperCase() ?? null,
      oil_type: txt(row, i.oilType),
    })
  }
  return { rows: [...out.values()], skipped, warnings }
}

/* ────────────────────────────────────────────────────────────────────
   3. ME2N — PO ที่เปิดไว้
   เก็บทุกแถวรวมที่ทำรับครบแล้ว (ค้างส่ง = 0) เพราะต้องแยกให้ออกระหว่าง
   "ME2N มีรายการนี้แต่ทำรับครบแล้ว" กับ "ME2N ยังไม่มีรายการนี้เลย"
   สองกรณีนี้ต้องคิดต่างกันตอนกระทบยอดของระหว่างทาง
                                                                      */

export interface TransitRow {
  plant_code: string; mat_code: string; po_no: string
  po_date: string | null
  qty_pcs: number        // ยังค้างส่ง
  ordered_pcs: number    // สั่งไว้ทั้งหมด
}

const toDate = (v: unknown): string | null => {
  if (v instanceof Date) return v.toISOString().slice(0, 10)
  if (typeof v === 'string' && /\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10)
  return null
}

export function parseME2N(grid: Grid): ParseResult<TransitRow> {
  const h = findHeader(grid, ['Plant', 'Material'])
  const c = indexer(grid[h])
  const i = {
    plant: need(c('Plant'), 'Plant'),
    mat: need(c('Material'), 'Material'),
    po: c('Purchasing Document', 'PO'),
    date: c('Document Date'),
    qty: need(c('Still to be delivered (qty)', 'Still to be deliv'), 'Still to be delivered'),
    ord: c('Order Quantity'),
  }
  const agg = new Map<string, TransitRow>()
  let skipped = 0
  let openLines = 0

  for (let r = h + 1; r < grid.length; r++) {
    const row = grid[r]; if (!row) continue
    const plant = String(row[i.plant] ?? '').trim()
    const mat = String(row[i.mat] ?? '').trim().replace(/\.0$/, '')
    if (!plant || !mat) { skipped++; continue }

    const qty = num(row[i.qty])
    const po = i.po >= 0 ? String(row[i.po] ?? '').trim() : ''
    const k = `${plant}|${mat}|${po}`
    const prev = agg.get(k)
    if (prev) {
      prev.qty_pcs += qty
      prev.ordered_pcs += i.ord >= 0 ? num(row[i.ord]) : 0
    } else {
      agg.set(k, {
        plant_code: plant, mat_code: mat, po_no: po,
        po_date: i.date >= 0 ? toDate(row[i.date]) : null,
        qty_pcs: qty,
        ordered_pcs: i.ord >= 0 ? num(row[i.ord]) : qty,
      })
    }
    if (qty > 0) openLines++
  }

  const rows = [...agg.values()]
  const warnings: string[] = []
  const open = rows.filter((r) => r.qty_pcs > 0)
  warnings.push(
    `ยังค้างส่ง ${open.length} รายการ ${open.reduce((s, r) => s + r.qty_pcs, 0).toLocaleString()} ชิ้น` +
    ` · ทำรับครบแล้ว ${rows.length - open.length} รายการ`
  )
  void openLines
  return { rows, skipped, warnings }
}


/* ────────────────────────────────────────────────────────────────────
   4. WMS — สต็อกคลัง
   รูปแบบใหม่ (ส.ค. 69): Item | Description | On Hand
   รองรับหัวตารางแบบเก่าไว้ด้วย เผื่อไฟล์เดิมยังใช้อยู่
                                                                      */

export interface DepotRow { mat_code: string; qty_pcs: number }

export function parseWMS(grid: Grid): ParseResult<DepotRow> {
  let h: number
  try { h = findHeader(grid, ['Item', 'On Hand']) }
  catch {
    try { h = findHeader(grid, ['Material']) }
    catch { h = findHeader(grid, ['รหัสสินค้า']) }
  }
  const c = indexer(grid[h])
  const mi = need(c('Item', 'Material', 'รหัสสินค้า', 'MATERIAL CODE'), 'Item / Material')
  const qi = need(
    c('On Hand', 'OnHand', 'Qty', 'จำนวน', 'Unrestricted', 'คงเหลือ', 'Stock'),
    'On Hand / จำนวน'
  )

  const agg = new Map<string, number>()
  let skipped = 0
  for (let r = h + 1; r < grid.length; r++) {
    const row = grid[r]; if (!row) continue
    const mat = String(row[mi] ?? '').trim().replace(/\.0$/, '')
    if (!/^\d{6,}$/.test(mat)) { skipped++; continue }
    agg.set(mat, (agg.get(mat) ?? 0) + num(row[qi]))
  }

  const rows = [...agg].map(([mat_code, qty_pcs]) => ({ mat_code, qty_pcs }))
  const warnings: string[] = []
  if (!rows.length) warnings.push('อ่านไฟล์ WMS ไม่ได้สักแถว — ตรวจว่าเลือกชีตถูกไหม')
  else {
    const zero = rows.filter((r) => r.qty_pcs <= 0).length
    if (zero) warnings.push(`${zero} SKU ของหมดในคลัง — ระบบจะไม่สั่งตัวเหล่านี้`)
  }
  return { rows, skipped, warnings }
}


/* ────────────────────────────────────────────────────────────────────
   5. Datastation2 — master สถานีทั้งประเทศ
   ให้คู่ Site Code2 (รหัสหน้าชื่อ) → Plant Code (รหัสใน POWER_BI)
                                                                      */

export interface MasterStationRow {
  plant_code: string; site_code_1: string | null; site_code_2: string | null
  kind: string | null; station_type: string | null
  station_name: string | null; province: string | null; area: string | null
  /** ผู้จัดการเขต — ขอบเขตที่ใช้จริงในการโอนเกลี่ย ไม่ใช่คอลัมน์ 'เขตพื้นที่' */
  area_id: string | null; area_manager: string | null
  area_phone: string | null; region_name: string | null
  district: string | null; subdistrict: string | null
  closed_date: string | null
}

export function parseDatastation(grid: Grid): ParseResult<MasterStationRow> {
  const h = findHeader(grid, ['Plant Code', 'Site Code2'])
  const c = indexer(grid[h])
  const i = {
    plant: need(c('Plant Code'), 'Plant Code'),
    s1: c('Site Code1'), s2: need(c('Site Code2'), 'Site Code2'),
    name: c('ชื่อสถานีบริการ'), prov: c('จังหวัด'), area: c('เขตพื้นที่'),
    kind: c('ชนิด'), type: c('Type'), closed: c('ปิดสถานี'),
    aid: c('ID_area'), amgr: c('Name_Area'), asur: c('Surname_Area'),
    aph: c('Phone_Area'), rgn: c('Name_Region'),
    tam: c('ตำบล/แขวง'), amp: c('อำเภอ/เขต'),
    amphoe: c('อำเภอ/เขต'), tambon: c('ตำบล/แขวง'),
  }
  const out = new Map<string, MasterStationRow>()
  let skipped = 0

  for (let r = h + 1; r < grid.length; r++) {
    const row = grid[r]; if (!row) continue
    const plant = String(row[i.plant] ?? '').trim().toUpperCase()
    if (!/^[A-Z0-9]{4}$/.test(plant)) { skipped++; continue }

    const closedRaw = i.closed >= 0 ? row[i.closed] : null
    let closed: string | null = null
    if (closedRaw instanceof Date) closed = closedRaw.toISOString().slice(0, 10)
    else if (typeof closedRaw === 'string' && /\d{4}-\d{2}-\d{2}/.test(closedRaw)) {
      closed = closedRaw.slice(0, 10)
    }

    // สาขาหนึ่งมีได้หลายแถว แยกตามชนิด (OIL/LPG) — คีย์จึงต้องรวม site code
    const s2 = String(row[i.s2] ?? '').trim().toUpperCase() || null
    out.set(`${plant}|${s2 ?? ''}`, {
      plant_code: plant,
      site_code_1: i.s1 >= 0 ? String(row[i.s1] ?? '').trim() || null : null,
      site_code_2: s2,
      kind: i.kind >= 0 ? String(row[i.kind] ?? '').trim() || null : null,
      station_type: i.type >= 0 ? String(row[i.type] ?? '').trim() || null : null,
      station_name: i.name >= 0 ? String(row[i.name] ?? '').trim() || null : null,
      area_id: i.aid >= 0 ? String(row[i.aid] ?? '').trim() || null : null,
      area_manager: i.amgr >= 0
        ? [String(row[i.amgr] ?? '').trim(), i.asur >= 0 ? String(row[i.asur] ?? '').trim() : '']
            .filter(Boolean).join(' ') || null
        : null,
      area_phone: i.aph >= 0 ? String(row[i.aph] ?? '').trim() || null : null,
      region_name: i.rgn >= 0 ? String(row[i.rgn] ?? '').trim() || null : null,
      subdistrict: i.tam >= 0 ? String(row[i.tam] ?? '').trim() || null : null,
      district: i.amp >= 0 ? String(row[i.amp] ?? '').trim() || null : null,
      province: i.prov >= 0 ? String(row[i.prov] ?? '').trim() || null : null,
      area: i.area >= 0 ? String(row[i.area] ?? '').trim() || null : null,
      closed_date: closed,
    })
  }
  return { rows: [...out.values()], skipped, warnings: [] }
}


/* ────────────────────────────────────────────────────────────────────
   6. เที่ยวรถ
   ชื่อลูกค้าเก็บรหัสสาขาไว้หลังขีดสุดท้าย ยาว 4 ตัวเสมอ:
     "076 - ท่ามะกา-S073"                    -> S073
     "40K - ถ.พระราม2(กม.53)-SG72"           -> SG72
     "08F - กระทุ่มแบน5(ถ.เศรษฐกิจ)-SB50"    -> SB50
   ตรวจแล้วกับ PlantCode ทั้ง 151 รหัส และไฟล์เที่ยวรถทั้งหมด — ยาว 4 ตัวทุกรายการ
   แถวที่ไม่มีขีดท้าย เช่น "อิสรินทร์" คือลูกค้าที่ไม่ใช่สถานีในความดูแล
                                                                      */

/** คลังที่รับที่ไม่นับเป็นเที่ยวส่ง — รถโอนเป็นเอาท์ซอส ไม่ได้เข้าคลัง */
export const EXCLUDED_PICKUP = ['รถโอน']

export interface TripRow {
  /** รหัสหน้าชื่อ = Site Code2 — ตัวหลักที่ใช้จับคู่ เชื่อถือได้กว่ารหัสท้าย */
  front_code: string | null
  /** รหัสท้ายชื่อ — บางสาขาตรงกับ PlantCode บางสาขาไม่ตรง */
  tail_code: string | null
  trip_no: number | null
  pickup_point: string | null
  /** ชื่อลูกค้าเต็มตามไฟล์ ใช้แสดงตอนต้องผูกรหัสเอง */
  source_name: string
}

export function parseTrips(grid: Grid): ParseResult<TripRow> & { excluded: number } {
  const h = findHeader(grid, ['ชื่อลูกค้า'])
  const c = indexer(grid[h])
  const ci = need(c('ชื่อลูกค้า'), 'ชื่อลูกค้า')
  const ti = c('เที่ยว', 'Trip No.')
  const pi = c('คลังที่รับ')

  const out = new Map<string, TripRow>()
  const warnings: string[] = []
  let skipped = 0
  let excluded = 0

  for (let r = h + 1; r < grid.length; r++) {
    const row = grid[r]
    if (!row) continue

    const raw = String(row[ci] ?? '').trim()
    if (!raw) { skipped++; continue }

    const pickup = pi >= 0 ? String(row[pi] ?? '').trim() : ''

    // ตัดรถโอนออกก่อน — ของไม่ได้ผ่านคลัง จึงไม่นับเป็นรอบส่งของเรา
    if (EXCLUDED_PICKUP.some((x) => pickup.includes(x))) { excluded++; continue }

    // รหัสหน้า: อยู่ก่อนขีดแรก ยาวไม่คงที่ (076, 11H, Y69)
    const fm = raw.match(/^\s*([A-Za-z0-9]{2,4})\s*-/)
    const front = fm ? fm[1].toUpperCase() : null

    // รหัสท้าย: หลังขีดสุดท้าย ยาว 4 ตัวเสมอ
    const dash = raw.lastIndexOf('-')
    const t = dash >= 0 ? raw.slice(dash + 1).trim().toUpperCase() : ''
    const tail = /^[A-Z0-9]{4}$/.test(t) ? t : null

    if (!front && !tail) {
      if (warnings.length < 5) warnings.push(`ดึงรหัสสาขาจาก "${raw}" ไม่ได้ — ข้ามแถวนี้`)
      skipped++
      continue
    }

    out.set(front ?? tail!, {
      front_code: front,
      tail_code: tail,
      trip_no: ti >= 0 ? num(row[ti]) || null : null,
      pickup_point: pickup || null,
      source_name: raw,
    })
  }

  if (excluded) warnings.push(`ตัดรถโอนออก ${excluded} แถว (ไม่ได้เข้าคลัง จึงไม่นับเป็นรอบส่ง)`)

  return { rows: [...out.values()], skipped, warnings, excluded }
}
