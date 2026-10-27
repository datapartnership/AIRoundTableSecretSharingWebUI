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

/** Reads a message to screen readers through the shared live region. */
export function announce(text) {
  const live = document.getElementById('live')
  if (live) live.textContent = text
}

export function download(name, text, type = 'text/csv') {
  const a = document.createElement('a')
  a.href = URL.createObjectURL(new Blob([text], { type }))
  a.download = name
  a.click()
  setTimeout(() => URL.revokeObjectURL(a.href), 1000)
}
