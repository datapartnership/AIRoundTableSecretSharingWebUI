import { Link } from 'react-router-dom'
import { PageNav, useShell } from '../components/Shell'
import { Alert, Disclosure, Footer, HeroArt, ProcessList, Toggletip, epochTag } from '../components/ui'
import useEpochs, { byStartDesc, contribution } from '../hooks/useEpochs'
import { USAGE, USAGE_TYPES, monthYear } from '../utils/usage'

const NOTE = { todo: 'Please share your data', done: 'Your masked data were submitted' }
const ROWS_PER_LIST = 2

export default function Home() {
  const shell = useShell()
  if (!shell.signedIn) return <HomeOut onSignIn={shell.signIn} />
  return <HomeIn />
}

function HomeOut({ onSignIn }) {
  return (
    <>
      <div className="home-top out">
        <PageNav home />
        <main id="main">
          <div className="hero-field">
            <HeroArt signedOut />
            <section className="hero hero-out wrap">
              <h1 tabIndex={-1}>Global AI Adoption Privacy-Preserving Data Sharing</h1>
              <p className="sub">Single Sign-On Platform for Participating Organization to share masked data submissions with pairwise noise cancellation for the Global AI Adoption Index</p>
              <button className="btn btn-white" onClick={onSignIn}>Sign in</button>
              <p className="note">Uses your organization&apos;s authentication. If your organization does not use Microsoft Authentication, you will receive an email on your registered email ID. In case of trouble logging in, please write to datapartnership@worldbank.org</p>
            </section>
          </div>
          <section className="process wrap" aria-labelledby="process-h">
            <h2 className="section-label" id="process-h">Secure data sharing process</h2>
            <ProcessList />
          </section>
        </main>
      </div>
      <Footer />
    </>
  )
}

const isOpen = (e) => !e.isClosed && !e.cancelledAt

function heroCopy({ loading, error, epochs, isAdmin }) {
  if (loading) return { h1: 'Checking your epochs', sub: 'This takes a moment.' }
  if (error) return { h1: 'Your epochs could not be loaded', sub: 'Your work is not affected. Try again, or contact the platform administrator if this continues.', retry: true }
  const open = epochs.filter(isOpen)
  if (open.length === 0) return { h1: 'No epoch is open right now', sub: 'The next epoch will appear below when it opens.' }
  const needs = isAdmin ? [] : open.filter((e) => contribution(e) === 'todo')
  if (needs.length >= 2) return { h1: 'Two epochs are open for your contribution', sub: 'An epoch will only be created if there are 3 contributing partners.' }
  if (needs.length === 1) return { h1: `${USAGE[needs[0].usageType].label}, ${monthYear(needs[0].startDate)} is open for your contribution` }
  return { h1: 'Nothing needs your contribution right now', sub: 'The open epochs are listed below, with your status for each.' }
}

function HomeIn() {
  const { isAdmin, registration, retryRegistration } = useShell()
  const { epochs: all, error, loading, retry } = useEpochs()
  // Partners see their own epochs; admins see every epoch
  const epochs = (all ?? []).filter((e) => isAdmin || e.isEligible)
  const hero = heroCopy({ loading, error, epochs, isAdmin })

  const skeleton = [0, 1].map((i) => (
    <li key={i} className="skel"><div className="bars"><i style={{ width: 160, height: 14 }} /><i style={{ width: 220, height: 11 }} /></div><i style={{ width: 90, height: 24 }} /></li>
  ))

  const list = (u) => {
    const rows = epochs.filter((e) => e.usageType === u).sort(byStartDesc).slice(0, ROWS_PER_LIST)
    return (
      <div className="elist" key={u}>
        <h2>{USAGE[u].label}</h2>
        <ul aria-busy={loading || undefined}>
          {loading ? skeleton : rows.length === 0
            ? <li className="erow"><p className="rnote">No epochs yet. The first epoch will appear here when it opens.</p></li>
            : rows.map((e) => {
              const mine = contribution(e)
              const month = monthYear(e.startDate)
              return (
                <li key={e.epochId} className="erow">
                  <div className="info">
                    <div className="name"><span className="month">{month}</span>{epochTag(e)}</div>
                    <p className="rnote">{e.cancelledAt ? 'Cancelled by an admin' : isAdmin ? `${e.producerCount} partners` : NOTE[mine]}</p>
                  </div>
                  {!isAdmin && isOpen(e) && mine === 'todo' && (
                    <Link className="btn btn-primary" to={`/${u}?epoch=${e.epochId}`} aria-label={`Continue ${USAGE[u].label}, ${month}`}>Continue</Link>
                  )}
                </li>
              )
            })}
        </ul>
      </div>
    )
  }

  return (
    <>
      <div className="home-top">
        <PageNav home />
        <main id="main">
          <div className="hero-field">
            <HeroArt />
            <section className="hero hero-in wrap">
              <h1 tabIndex={-1}>{hero.h1}</h1>
              {hero.sub && <p className="sub">{hero.sub}</p>}
              {hero.retry && <button className="btn btn-on-dark" onClick={retry}>Try again</button>}
            </section>
          </div>
          <section className="epochs wrap" aria-labelledby="epochs-h">
            <div className="heading-row">
              <h2 className="section-label" id="epochs-h">Your epochs</h2>
              <Toggletip id="tt" label="About publication thresholds">
                <p>A combined total is published only when enough partners report on it: at least 3, or half of the partners in the epoch, whichever is more.</p>
                <p>This is checked for each breakdown separately. For example, in an epoch with 10 partners, a country&apos;s breakdowns are published only if at least 5 partners report on that country.</p>
              </Toggletip>
            </div>
            {isAdmin && (
              <Alert kind="warning" title="Admins can't contribute data" lines={['Use Admin to open Quorum Check periods and epochs, and to follow submissions.']} />
            )}
            {registration.state === 'error' && (
              <Alert kind="error" title="Your organization could not be registered" lines={[`Server message: ${registration.error}`, 'Admins can add you to an epoch only after you are registered. Try again.']}>
                <div className="actions"><button className="btn btn-secondary" onClick={retryRegistration}>Try again</button></div>
              </Alert>
            )}
            {!error && (
              <>
                <div className="lists">{USAGE_TYPES.map(list)}</div>
                <div className="view-all"><Link className="link" to="/epochs">View all epochs →</Link></div>
              </>
            )}
          </section>
          <div className="how wrap">
            <Disclosure id="how" label="Secure data sharing process"><ProcessList /></Disclosure>
          </div>
        </main>
      </div>
      <Footer />
    </>
  )
}
