import { useCallback, useEffect, useRef, useState, type ReactNode } from 'react'
import { flushSync } from 'react-dom'
import type { Session } from '@supabase/supabase-js'
import { supabase } from './lib/supabase'
import Login, { Pending } from './views/Login'
import Dashboard from './views/Dashboard'
import ImportPage from './views/Import'
import Daily from './views/Daily'
import Run from './views/Run'
import Transfers from './views/Transfers'
import TransfersB from './views/TransfersB'
import ManualTransfer from './views/ManualTransfer'
import Shipto from './views/Shipto'
import Depot from './views/Depot'
import Receiving from './views/Receiving'
import Shortage from './views/Shortage'
import ClassPrep from './views/ClassPrep'
import StationGroups from './views/StationGroups'
import KpiPage from './views/Kpi'
import Settings from './views/Settings'
import Users from './views/Users'

const today = () => new Date().toISOString().slice(0, 10)

/** จำหน้าที่เปิดค้างไว้ 1 ชั่วโมง รีเฟรชแล้วกลับมาที่เดิม
 *  ใช้ sessionStorage ไม่ได้เพราะปิดแท็บแล้วหาย จึงใช้ localStorage พร้อมเวลาหมดอายุ */
const TAB_KEY = 'mkdepot.tab'
const TAB_TTL = 60 * 60 * 1000

function loadTab(): { tab: string; back: boolean } | null {
  try {
    const raw = localStorage.getItem(TAB_KEY)
    if (!raw) return null
    const v = JSON.parse(raw) as { tab: string; back: boolean; at: number }
    if (Date.now() - v.at > TAB_TTL) { localStorage.removeItem(TAB_KEY); return null }
    return { tab: v.tab, back: v.back }
  } catch { return null }
}

function saveTab(tab: string, back: boolean) {
  try {
    localStorage.setItem(TAB_KEY, JSON.stringify({ tab, back, at: Date.now() }))
  } catch { /* โหมดส่วนตัวอาจเขียนไม่ได้ ไม่ใช่เรื่องคอขาดบาดตาย */ }
}

type VTDoc = Document & {
  startViewTransition?: (cb: () => void) => { finished: Promise<void> }
}
type Tab = 'home' | 'daily' | 'import' | 'run' | 'transfer' | 'transferB' | 'manual' | 'shipto' | 'receiving'
  | 'shortage' | 'classprep' | 'groups' | 'depot' | 'kpi' | 'settings' | 'users'

/** ตัวกรองที่ส่งข้ามหน้าได้ เช่นกดตัวเลขในหน้าภาพรวมแล้วเด้งไปหน้าของขาด */
export interface Preset { kind?: string; name?: string }
interface Me { username: string; role: 'staff' | 'admin'; is_active: boolean }
interface Item { id: Tab; label: string; step?: string }

export default function App() {
  const [session, setSession] = useState<Session | null>(null)
  const [me, setMe] = useState<Me | null>(null)
  const [ready, setReady] = useState(false)

  // อ่านค่าที่จำไว้ครั้งเดียวตอนเปิดหน้า
  // ถ้าจำหน้าหลังบ้านไว้แต่โหมดเป็นโหมดง่าย ให้กลับไปหน้างานประจำวัน
  const [saved] = useState(() => {
    const v = loadTab()
    if (!v) return null
    if (!v.back && v.tab !== 'daily' && v.tab !== 'home') return { tab: 'daily', back: false }
    return v
  })
  const [tab, setTab] = useState<Tab>((saved?.tab as Tab) ?? 'daily')
  const tabRef = useRef<Tab>((saved?.tab as Tab) ?? 'daily')
  const [snapshotDate, setSnapshotDate] = useState(today())
  const [preset, setPreset] = useState<Preset | undefined>()
  /** โหมดง่ายซ่อนเมนูทั้งหมด เหลือเฉพาะงานประจำวัน · โหมดเต็มคือหลังบ้าน */
  const [full, setFull] = useState(saved?.back ?? false)

  useEffect(() => { tabRef.current = tab; saveTab(tab, full) }, [tab, full])

  const enterFull = useCallback(() => {
    setFull(true)
    navigate('home')
  }, [])

  const exitFull = useCallback(() => {
    setFull(false)
    navigate('daily')
  }, [])


  /** เปลี่ยนหน้าแบบมอร์ฟ — เบราว์เซอร์ถ่ายภาพหน้าเดิมแล้วค่อย ๆ กลายเป็นหน้าใหม่
   *  flushSync บังคับให้ DOM อัปเดตเสร็จภายในคอลแบ็ก ไม่งั้นภาพจะไม่ตรง
   *  เบราว์เซอร์ที่ไม่รองรับจะเปลี่ยนหน้าตามปกติ */
  const navigate = useCallback((next: string, p?: Preset) => {
    const go = () => { setPreset(p); setTab(next as Tab) }
    if (next === tabRef.current && !p) return
    const doc = document as VTDoc
    if (!doc.startViewTransition || window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
      go(); return
    }
    doc.startViewTransition(() => { flushSync(go) })
  }, [])

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session))
    const { data: sub } = supabase.auth.onAuthStateChange((_e, s) => {
      setSession(s)
      if (!s) { setMe(null); setReady(true) }
    })
    return () => sub.subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (!session) { setReady(true); return }
    let cancelled = false
    supabase.from('app_users').select('username, role, is_active')
      .eq('user_id', session.user.id).maybeSingle()
      .then(({ data }) => {
        if (cancelled) return
        setMe((data as Me | null) ?? {
          username: session.user.email?.split('@')[0] ?? 'ผู้ใช้',
          role: 'staff', is_active: false,
        })
        setReady(true)
      })
    return () => { cancelled = true }
  }, [session])

  if (!ready) return null
  if (!session) return <Login />
  if (!me?.is_active) return <Pending username={me?.username ?? ''} />

  /** เมนูแบ่งตามลักษณะงาน ไม่ใช่เรียงตัวเลขยาว ๆ */
  const groups: { label: string | null; items: Item[] }[] = [
    {
      label: null,
      items: [
        { id: 'home', label: 'ภาพรวม' },
        { id: 'daily', label: 'งานประจำวัน' },
      ],
    },
    {
      label: 'งานประจำวัน',
      items: [
        { id: 'import', step: '1', label: 'นำเข้าข้อมูล' },
        { id: 'run', step: '2', label: 'คำนวณยอดเติม' },
        { id: 'transfer', step: '3', label: 'โอนเกลี่ยสินค้า' },
        { id: 'transferB', step: '4', label: 'โอนข้ามคลาส' },
        { id: 'manual', label: 'โอนกำหนดเอง' },
        { id: 'shipto', label: 'Shipto สั่งตรง' },
      ],
    },
    {
      label: 'ตรวจสอบ',
      items: [
        { id: 'shortage', label: 'ของขาด' },
        { id: 'receiving', label: 'ยังไม่ได้ทำรับ' },
        { id: 'depot', label: 'สั่งเข้าคลัง' },
        { id: 'classprep', label: 'เตรียมปรับคลาส' },
      ],
    },
    { label: 'ติดตามผล', items: [{ id: 'kpi', label: 'KPI ย้อนหลัง' }] },
    {
      label: 'ระบบ',
      items: [
        { id: 'groups', label: 'กลุ่มสถานี' },
        { id: 'settings', label: 'ตั้งค่าการคำนวณ' },
        ...(me.role === 'admin' ? [{ id: 'users' as Tab, label: 'ผู้ใช้' }] : []),
      ],
    },
  ]

  const initial = (me.username || '?').trim().charAt(0).toUpperCase()

  /** ไอคอนของแต่ละเมนู ใช้เส้นบางให้เข้ากับสไตล์โดยรวม */
  const ICON: Record<string, ReactNode> = {
    home: <path d="M4 10.5 12 4l8 6.5V19a1 1 0 0 1-1 1h-4v-6H9v6H5a1 1 0 0 1-1-1z" />,
    daily: <><rect x="4" y="4" width="7" height="7" rx="2" /><rect x="13" y="4" width="7" height="7" rx="2" /><rect x="4" y="13" width="7" height="7" rx="2" /><rect x="13" y="13" width="7" height="7" rx="2" /></>,
    import: <path d="M12 15V4m0 0L8.5 7.5M12 4l3.5 3.5M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3" />,
    run: <><rect x="4.5" y="3" width="15" height="18" rx="2.5" /><path d="M8.5 8h7M8.5 12h3M8.5 16h3M15 12v4" /></>,
    transfer: <path d="M7 7h11l-3-3m3 13H7l3 3" />,
    transferB: <path d="M4 8.5 12 4l8 4.5v7L12 20l-8-4.5z M4 8.5 12 13l8-4.5M12 13v7" />,
    manual: <path d="M12 5v14M5 12h14" />,
    shipto: <><path d="M3.5 7.5 9 5l5.5 2.5v6L9 16l-5.5-2.5z M3.5 7.5 9 10l5.5-2.5M9 10v6" /><path d="M15 17h6m0 0-2.5-2.5M21 17l-2.5 2.5" /></>,
    shortage: <path d="M12 8v5m0 3.5h.01M10.3 3.9 2.6 17.1A1.6 1.6 0 0 0 4 19.5h16a1.6 1.6 0 0 0 1.4-2.4L13.7 3.9a1.6 1.6 0 0 0-2.8 0Z" />,
    receiving: <><path d="M3 7h10v9H3zM13 10h4l3 3v3h-7z" /><circle cx="7" cy="18" r="1.6" /><circle cx="17" cy="18" r="1.6" /></>,
    depot: <path d="M4 8.5 12 4l8 4.5v7L12 20l-8-4.5z M4 8.5 12 13l8-4.5M12 13v7" />,
    kpi: <path d="M4 19V5m0 14h16M8 16v-5m4 5V7m4 9v-3" />,
    groups: <><circle cx="8" cy="9" r="3" /><circle cx="17" cy="7.5" r="2.2" /><path d="M3 19c0-2.8 2.2-5 5-5s5 2.2 5 5M15 19c0-2 .8-3.6 2-4.4" /></>,
    settings: <><circle cx="12" cy="12" r="3.2" /><path d="M12 2.5v2M12 19.5v2M4.2 4.2l1.4 1.4M18.4 18.4l1.4 1.4M2.5 12h2M19.5 12h2M4.2 19.8l1.4-1.4M18.4 5.6l1.4-1.4" /></>,
    users: <><circle cx="12" cy="8" r="3.4" /><path d="M5 20c0-3.6 3.1-6.5 7-6.5s7 2.9 7 6.5" /></>,
  }

  const Icon = ({ id }: { id: string }) => (
    <svg viewBox="0 0 24 24" width="18" height="18" fill="none"
      stroke="currentColor" strokeWidth="1.7"
      strokeLinecap="round" strokeLinejoin="round">
      {ICON[id] ?? ICON.home}
    </svg>
  )

  const simpleNav: Item[] = [
    { id: 'daily', label: 'งานประจำวัน' },
    { id: 'home', label: 'ภาพรวม' },
  ]

  return (
    <div className="shell">
      <aside className="rail">
        <div className="rail-top">
          <span className="logo" aria-hidden="true">
            <svg viewBox="0 0 24 24" width="18" height="18" fill="none">
              <path d="M4 17V9m0 0 8-5 8 5m-16 0 8 5 8-5m0 0v8m-8 5v-8"
                stroke="currentColor" strokeWidth="2"
                strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </span>
          <span className="brand-txt">
            <b>เติมสินค้า</b>
            <i>คลังแม่กลอง</i>
          </span>
        </div>

        <nav className="nav">
          {full ? groups.map((g, gi) => (
            <div className="nav-group" key={gi}>
              {g.label && <p className="nav-label">{g.label}</p>}
              {g.items.map((t) => (
                <button key={t.id} aria-current={tab === t.id}
                  onClick={() => navigate(t.id)}>
                  <Icon id={t.id} />
                  <span>{t.label}</span>
                  {t.step && <i className="step">{t.step}</i>}
                </button>
              ))}
            </div>
          )) : (
            <div className="nav-group">
              {simpleNav.map((t) => (
                <button key={t.id} aria-current={tab === t.id}
                  onClick={() => navigate(t.id)}>
                  <Icon id={t.id} />
                  <span>{t.label}</span>
                </button>
              ))}
            </div>
          )}
        </nav>

        <div className="rail-foot">
          <button className="mode-btn" onClick={full ? exitFull : enterFull}>
            {full ? 'โหมดง่าย' : 'หลังบ้าน'}
          </button>
          <div className="userchip" title={me.username}>
            <span className="avatar">{initial}</span>
            <span className="uname">
              <b>{me.username}</b>
              {me.role === 'admin' && <i>ผู้ดูแล</i>}
            </span>
            <button className="iconbtn sm" title="ออกจากระบบ"
              onClick={() => supabase.auth.signOut()}>
              <svg viewBox="0 0 24 24" width="15" height="15" fill="none">
                <path d="M15 17l5-5-5-5M20 12H9M12 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h6"
                  stroke="currentColor" strokeWidth="1.9"
                  strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
          </div>
        </div>
      </aside>

      <main className="main">
        {/* key=tab ทำให้ React สร้างใหม่ แอนิเมชันจึงเล่นซ้ำ */}
        <div className="page" key={tab}>
          {tab === 'home' && <Dashboard go={navigate} />}
          {tab === 'daily' && (
            <Daily snapshotDate={snapshotDate} setSnapshotDate={setSnapshotDate} go={navigate} />
          )}
          {tab === 'import' && <ImportPage snapshotDate={snapshotDate} setSnapshotDate={setSnapshotDate} />}
          {tab === 'run' && <Run snapshotDate={snapshotDate} />}
          {tab === 'transfer' && <Transfers snapshotDate={snapshotDate} />}
          {tab === 'transferB' && <TransfersB snapshotDate={snapshotDate} />}
          {tab === 'manual' && <ManualTransfer />}
          {tab === 'shipto' && <Shipto />}
          {tab === 'shortage' && <Shortage preset={preset} />}
          {tab === 'receiving' && <Receiving />}
          {tab === 'depot' && <Depot />}
          {tab === 'classprep' && <ClassPrep />}
          {tab === 'kpi' && <KpiPage />}
          {tab === 'groups' && <StationGroups />}
          {tab === 'settings' && <Settings />}
          {tab === 'users' && me.role === 'admin' && <Users me={me.username} />}
        </div>
      </main>
    </div>
  )
}
