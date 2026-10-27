import { useEffect, useRef, useState } from 'react'
import { Link, NavLink } from 'react-router-dom'
import { ICONS } from './icons'
import { HERO_ART } from './heroArt'
import { APP_VERSION } from '../../version'
import { USAGE, USAGE_TYPES } from '../../utils/usage'

export function Icon({ name, size = 16 }) {
  return (
    <svg className="icon" width={size} height={size} viewBox="0 0 14 14" aria-hidden="true" focusable="false">
      {ICONS[name].map((d, i) => <path key={i} d={d} />)}
    </svg>
  )
}

export function HeroArt({ signedOut = false }) {
  return (
    <div className="hero-art" aria-hidden="true">
      {HERO_ART.filter(([cls]) => signedOut || !cls.split(' ').includes('so')).map(([cls, style], i) => (
        <i key={i} className={cls} style={style} />
      ))}
    </div>
  )
}

export function Tag({ kind, children }) {
  return <span className={`tag tag-${kind}`}>{children}</span>
}

export const epochTag = (e) => e.cancelledAt
  ? <Tag kind="blocked">Cancelled</Tag>
  : e.isClosed ? <Tag kind="closed">Closed</Tag> : <Tag kind="open">Open</Tag>

// CAN Alert: no "Error:" or "Warning:" prefix; the why and the fix are separate body lines.
export function Alert({ kind = 'error', title, lines = [], children }) {
  return (
    <div className={`alert alert-${kind}`} role={kind === 'error' ? 'alert' : 'status'}>
      <Icon name={kind === 'error' ? 'warning-circle' : 'warning'} size={18} />
      <div>
        <p className="at">{title}</p>
        {lines.filter(Boolean).map((l, i) => <p key={i} className="ab">{l}</p>)}
        {children}
      </div>
    </div>
  )
}

const STEP_ICON = { done: 'check-circle', current: 'arrow-circle-right', waiting: 'clock', problem: 'warning-circle', later: 'circle' }
const STEP_WORD = { done: 'Done', current: 'Your next step', waiting: 'Waiting', problem: 'Problem', later: 'Not started' }
export const STEPS = ['Key-pair generation', 'Key exchange and shared secret generation', 'Data upload and verification', 'Masked data submission']

export function stepAnnouncement(strip) {
  const cur = strip.findIndex((k) => k !== 'done')
  return cur < 0 ? 'All steps done.' : `Step ${cur + 1}, ${STEPS[cur]}: ${STEP_WORD[strip[cur]]}.`
}

export function ProtocolStrip({ strip }) {
  return (
    <ol className="strip" aria-label="Protocol steps">
      {strip.map((k, i) => (
        <li key={i} className={`pstep ${k}`} aria-current={k === 'current' ? 'step' : undefined}>
          <div className="bar" />
          <div className="ptop">
            <span className="iline"><Icon name={STEP_ICON[k]} /><span className="sr-only">{STEP_WORD[k]}:</span></span>
            <span className="ptitle">{STEPS[i]}</span>
          </div>
        </li>
      ))}
    </ol>
  )
}

// Closes on outside click and on Escape (focus returns to the control).
function usePopover() {
  const [open, setOpen] = useState(false)
  const btnRef = useRef(null)
  const panelRef = useRef(null)
  useEffect(() => {
    if (!open) return
    const onClick = (e) => {
      if (panelRef.current?.contains(e.target) || btnRef.current?.contains(e.target)) return
      setOpen(false)
    }
    const onKey = (e) => {
      if (e.key !== 'Escape') return
      setOpen(false)
      btnRef.current?.focus()
    }
    document.addEventListener('click', onClick)
    document.addEventListener('keydown', onKey)
    return () => {
      document.removeEventListener('click', onClick)
      document.removeEventListener('keydown', onKey)
    }
  }, [open])
  return { open, setOpen, btnRef, panelRef }
}

export function Toggletip({ id, label, children }) {
  const { open, setOpen, btnRef, panelRef } = usePopover()
  return (
    <div className="tt-anchor">
      <button ref={btnRef} className="tt-btn" aria-expanded={open} aria-controls={id} aria-label={label} onClick={() => setOpen(!open)}>
        <Icon name="info" size={18} />
      </button>
      <div ref={panelRef} className="tt-panel" id={id} hidden={!open}>{children}</div>
    </div>
  )
}

export function Disclosure({ id, label, children }) {
  const [open, setOpen] = useState(false)
  return (
    <div className="disclosure">
      <button className="link" aria-expanded={open} aria-controls={id} onClick={() => setOpen(!open)}>
        {open ? 'Hide details' : label}
      </button>
      <div className="content" id={id} hidden={!open}>{children}</div>
    </div>
  )
}

// CAN Checkbox. `indeterminate` is a DOM property, so it is set on the element.
export function Checkbox({ checked, indeterminate = false, onChange, label, ariaLabel, disabled }) {
  const ref = useRef(null)
  useEffect(() => { if (ref.current) ref.current.indeterminate = indeterminate }, [indeterminate])
  return (
    <label className="cb">
      <input ref={ref} type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} aria-label={ariaLabel} disabled={disabled} />
      {label}
    </label>
  )
}

export function Nav({ signedIn, home = false, user, isAdmin = false, locked = false, onSignOut }) {
  const { open, setOpen, btnRef, panelRef } = usePopover()
  const cls = `nav wrap ${home ? 'transparent' : ''} on-dark`
  if (!signedIn) {
    return (
      <header className={cls}>
        <Link className="logo" to="/"><img src="/logo.png" alt="Development Data Partnership" /></Link>
      </header>
    )
  }
  // Quorum Check is the first nav item (2026-10-14); while a partner owes or awaits it, it is the only one
  const items = locked ? [['/quorum', 'Quorum Check']] : [
    ['/quorum', 'Quorum Check'],
    ['/', 'Home'],
    ...USAGE_TYPES.map((u) => [`/${u}`, USAGE[u].nav]),
    ...(isAdmin ? [['/admin', 'Admin']] : []),
  ]
  return (
    <header className={cls}>
      <Link className="logo" to={locked ? '/quorum' : '/'}><img src="/logo.png" alt="Development Data Partnership, Home" /></Link>
      <nav className="nav-links" aria-label="Main">
        {items.map(([to, label]) => (
          <NavLink key={to} to={to} end={to === '/'} className={`nav-link${to === '/admin' ? ' nav-admin' : ''}`}>{label}</NavLink>
        ))}
      </nav>
      <div className="auth">
        <div className="who"><b>{user?.name}</b>{user?.org && <span>{user.org}</span>}</div>
        <button className="btn btn-on-dark" onClick={onSignOut}>Sign out</button>
      </div>
      <button ref={btnRef} className="menu-control" aria-expanded={open} aria-controls="menu-panel" onClick={() => setOpen(!open)}>
        <Icon name="list" size={18} />Menu
      </button>
      <div ref={panelRef} className={`menu-panel${open ? ' open' : ''}`} id="menu-panel">
        {items.map(([to, label]) => (
          <NavLink key={to} to={to} end={to === '/'} onClick={() => setOpen(false)}>{label}</NavLink>
        ))}
        <hr />
        <button className="signout" onClick={onSignOut}>Sign out</button>
      </div>
    </header>
  )
}

export function Footer() {
  return (
    <footer className="footer wrap on-dark">
      <div className="cols">
        <div><h2>About</h2><p>The <a href="https://datapartnership.org/about" target="_blank" rel="noopener noreferrer">Development Data Partnership</a> is a collaboration between international organizations and technology companies, facilitating the efficient and responsible use of third-party data in international development.</p></div>
        <div><h2>Quick links</h2><a href="https://datapartnership.org/updates" target="_blank" rel="noopener noreferrer">Projects &amp; News</a></div>
        <div><h2>Contact</h2><a href="mailto:datapartnership@worldbank.org">datapartnership@worldbank.org</a></div>
        <div><h2>Legal</h2><a href="https://www.worldbank.org/en/about/legal" target="_blank" rel="noopener noreferrer">Legal Policies</a></div>
      </div>
      <hr />
      <small><span>© Development Data Partnership</span><span className="app-version">{APP_VERSION}</span></small>
    </footer>
  )
}

export const PROCESS = [
  [STEPS[0], 'Your browser creates an ML-KEM-768 key pair. The encapsulation key is registered with the Data Sharing API; the decapsulation key never leaves your browser.'],
  [STEPS[1], 'The keys are exchanged between every partner pair to derive a shared secret. This secret is never transmitted.'],
  [STEPS[2], 'You must upload the data as a CSV. Your browser checks your file against the pre-agreed format the epoch needs, and reports errors. Nothing is sent during the check. Once all errors are fixed, it moves to the next step.'],
  [STEPS[3], "Your browser adds noise derived from each shared secret to your values before submitting. When every partner's values are summed, the noise cancels out exactly."],
]

export function ProcessList() {
  return (
    <ol className="steps">
      {PROCESS.map(([t, d], i) => (
        <li key={i}><span className="num" aria-hidden="true">{i + 1}</span><div><div className="title">{t}</div><p className="desc">{d}</p></div></li>
      ))}
    </ol>
  )
}
