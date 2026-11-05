import { createContext, useContext } from 'react'
import { Nav, Footer } from './ui'

export const ShellContext = createContext(null)
export const useShell = () => useContext(ShellContext)

/** The navy nav bar for the signed-in or signed-out state; transparent over the Home hero. */
export function PageNav({ home = false }) {
  const shell = useShell()
  return <Nav signedIn={shell.signedIn} home={home} user={shell.user} isAdmin={shell.isAdmin} locked={!!shell.quorum?.locked} onSignOut={shell.signOut} />
}

/** Nav, page body, and footer for every page except Home, which puts the nav over its hero. */
export function Page({ children }) {
  return (
    <>
      <PageNav />
      {children}
      <Footer />
    </>
  )
}

export { announce, download } from '../utils/dom'
