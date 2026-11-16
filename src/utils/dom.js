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
