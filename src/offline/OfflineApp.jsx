import { useRef, useState } from 'react'
import { Alert } from '../components/ui'
import { FileCheck, Sec, ValidFile } from '../components/FileCheckView'
import { announce, download } from '../utils/dom'
import { HEADERS, isCsvFile, parseAndValidateCsv, requiredCells, valueRule } from '../utils/csvUpload'
import { buildFileCheck } from '../utils/fileCheck'
import { buildMaskedFile, openKeyFile, readKeyFileHeader } from '../utils/offlineBundle'
import { APP_VERSION } from '../version'
import { longDate, monthRange } from '../utils/usage'

const SEC_KEY = 'Step 1: Open your key file'
const SEC_FILE = 'Step 2: Upload and verify data'
const SEC_MASK = 'Step 3: Save the masked file'


/** Rows for every cell the key file expects; metrics that must be 0 are filled in, the rest are left empty. */
function templateCsv(kf) {
  const rows = requiredCells(kf.months, kf.quorumRules).map((c) =>
    [c.month, c.country, c.indicator, c.segment, valueRule(kf.quorumRules, c.country, c.indicator) === 'zero' ? '0' : ''].join(','))
  return `${[HEADERS.join(','), ...rows].join('\n')}\n`
}

export default function OfflineApp() {
  // Keys stay in memory only: closing or reloading this page forgets them
  const [keyText, setKeyText] = useState(null)        // { name, text, header }
  const [keyError, setKeyError] = useState(null)
  const [passphrase, setPassphrase] = useState('')
  const [opening, setOpening] = useState(false)
  const [keyFile, setKeyFile] = useState(null)
  const [upload, setUpload] = useState({ phase: 'idle' })
  const [saved, setSaved] = useState(null)
  const keyInput = useRef(null)
  const csvInput = useRef(null)

  const onKeyFile = async (event) => {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    setKeyError(null)
    setKeyFile(null)
    setUpload({ phase: 'idle' })
    setSaved(null)
    try {
      const text = await file.text()
      setKeyText({ name: file.name, text, header: readKeyFileHeader(text) })
    } catch (e) {
      setKeyText(null)
      setKeyError(e.message)
    }
  }

  const openKey = async (event) => {
    event.preventDefault()
    setOpening(true)
    setKeyError(null)
    try {
      const kf = await openKeyFile(keyText.text, passphrase)
      setKeyFile(kf)
      setPassphrase('')
      announce('Key file opened.')
    } catch (e) {
      setKeyError(e.message)
    } finally {
      setOpening(false)
    }
  }

  const forget = () => {
    setKeyText(null)
    setKeyFile(null)
    setPassphrase('')
    setUpload({ phase: 'idle' })
    setSaved(null)
    setKeyError(null)
  }

  const onCsv = async (event) => {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    setSaved(null)
    setUpload({ phase: 'checking', fileName: file.name })
    const at = new Date()
    let result
    if (!isCsvFile(file)) result = { fileError: { kind: 'not-csv' } }
    else {
      let text = null
      try { text = await file.text() } catch { result = { fileError: { kind: 'unreadable' } } }
      if (text !== null) {
        try { result = parseAndValidateCsv(text, keyFile.months, keyFile.quorumRules, keyFile.maskBits) } catch (e) {
          console.error('[offline] csv check failed', e)
          result = { fileError: { kind: 'platform' } }
        }
      }
    }
    setUpload({ phase: 'checked', fileName: file.name, at, result, view: buildFileCheck(result, keyFile.months, file.name) })
  }

  const saveMasked = () => {
    const masked = buildMaskedFile(keyFile, upload.result.rows, APP_VERSION)
    const name = `airt-epoch-${keyFile.epochId}-masked.json`
    download(name, JSON.stringify(masked, null, 2), 'application/json')
    setSaved({ name, count: masked.values.length, at: new Date() })
    announce(`Masked file ${name} saved.`)
  }

  const chooseCsv = () => csvInput.current?.click()
  const downloadTemplate = () => download(`template-${keyFile.months[0]}.csv`, templateCsv(keyFile))
  const epochName = (kf) => kf.epochLabel ?? `Epoch ${kf.epochId}`

  let keySec
  if (keyFile) {
    keySec = (
      <Sec h={SEC_KEY}>
        <p>Key file for the <b>{epochName(keyFile)}</b> epoch is open. Reporting months: {monthRange(`${keyFile.months[0]}-01`, keyFile.months.length)}. Exported on {longDate(keyFile.createdAt)}.</p>
        <div className="actions"><button className="btn btn-secondary" onClick={forget}>Close the key file</button></div>
      </Sec>
    )
  } else {
    keySec = (
      <Sec h={SEC_KEY}>
        <p>Choose the key file you exported from the epoch page on the computer with internet access, then enter its passphrase.</p>
        {keyError && <Alert kind="error" title="The key file could not be opened" lines={[keyError]} />}
        {keyText ? (
          <form className="offline-form" onSubmit={openKey}>
            <div className="file">
              <div className="fname">{keyText.name}</div>
              <div className="fmeta">{keyText.header.epochLabel ? `${keyText.header.epochLabel} epoch` : `Epoch ${keyText.header.epochId}`} | Exported on {longDate(keyText.header.createdAt)}</div>
            </div>
            <label className="field">Passphrase
              <input type="password" autoComplete="off" value={passphrase} onChange={(e) => setPassphrase(e.target.value)} autoFocus />
            </label>
            <div className="actions">
              <button className="btn btn-primary" type="submit" disabled={!passphrase || opening}>{opening ? 'Opening' : 'Open key file'}</button>
              <button className="btn btn-secondary" type="button" onClick={() => keyInput.current?.click()} disabled={opening}>Choose a different key file</button>
            </div>
          </form>
        ) : (
          <div className="actions"><button className="btn btn-primary" onClick={() => keyInput.current?.click()}>Choose key file</button></div>
        )}
      </Sec>
    )
  }

  let fileSec
  if (!keyFile) {
    fileSec = (
      <Sec h={SEC_FILE}>
        <p>Available when your key file is open.</p>
      </Sec>
    )
  } else if (upload.phase === 'checked' && upload.result.ok) {
    fileSec = <ValidFile upload={upload} onChoose={chooseCsv} h={SEC_FILE} />
  } else if (upload.phase === 'checked') {
    fileSec = <FileCheck upload={upload} onChoose={chooseCsv} onSample={downloadTemplate} h={SEC_FILE} labels={{ sample: 'Download template.csv' }} />
  } else {
    const ignored = [...keyFile.quorumRules.ignoredPairs].map((k) => k.replace('|', ' ')).join(', ')
    fileSec = (
      <Sec h={SEC_FILE}>
        <p>Your CSV needs {requiredCells(keyFile.months, keyFile.quorumRules).length} rows for {monthRange(`${keyFile.months[0]}-01`, keyFile.months.length)}.{keyFile.quorumRules.answers ? ' Report values for the metrics you ticked in the Quorum Check, and enter 0 for the others.' : ''}{ignored ? ` Leave out the metrics below the quorum: ${ignored}.` : ''} The file is checked in this browser only.</p>
        <div className="actions">
          <button className="btn btn-primary" onClick={chooseCsv} disabled={upload.phase === 'checking'}>Upload file</button>
          <button className="btn btn-secondary" onClick={downloadTemplate}>Download template.csv</button>
        </div>
      </Sec>
    )
  }

  const ready = keyFile && upload.phase === 'checked' && upload.result.ok
  const maskSec = (
    <Sec h={SEC_MASK}>
      <p>This tool masks each value with your key file and saves only the masked values. Copy the masked file to the computer with internet access, open the epoch page and choose <b>Submit masked file</b>.</p>
      {saved && (
        <Alert kind="success" title={`Saved ${saved.name}`} lines={[
          `${saved.count} masked values. Check your downloads folder.`,
          'When your contribution is submitted, delete the key file from both computers. Anyone with the key file and its passphrase can unmask your values.',
        ]} />
      )}
      <div className="actions">
        <button className="btn btn-primary" onClick={saveMasked} disabled={!ready} aria-describedby={ready ? undefined : 'why-mask'}>{saved ? 'Save the masked file again' : 'Save masked file'}</button>
        {!ready && <span className="why" id="why-mask">Available when the file is valid</span>}
      </div>
    </Sec>
  )

  return (
    <>
      <header className="nav wrap on-dark"><span className="offline-title">AI Roundtable offline masking tool</span></header>
      <main id="main" className="epoch wrap">
        <div className="epoch-head">
          <p className="usage">Works without internet access</p>
          <h1 tabIndex={-1}>Mask your data offline</h1>
          <p className="meta">Nothing you open here leaves this computer. This page makes no network requests and forgets your key when you close it.</p>
        </div>
        <div className="main">
          {keySec}
          {fileSec}
          {maskSec}
        </div>
      </main>
      <footer className="footer wrap on-dark"><small><span>© Development Data Partnership</span><span className="app-version">{APP_VERSION}</span></small></footer>
      <input ref={keyInput} type="file" accept=".json,application/json" onChange={onKeyFile} hidden />
      <input ref={csvInput} type="file" accept=".csv,text/csv" onChange={onCsv} hidden />
      <div id="live" className="sr-only" role="status" aria-live="polite" />
    </>
  )
}
