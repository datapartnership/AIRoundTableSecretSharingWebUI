import { useState } from 'react'
import { Tag } from './ui'
import { announce, download } from '../utils/dom'
import { HEADERS } from '../utils/csvUpload'
import { problemReportCsv } from '../utils/fileCheck'
import { clockTime } from '../utils/usage'

// The checked-file views of the epoch page, shared with the offline masking tool
export const SEC_FILE = 'Upload and verify data'
export const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`
export const Sec = ({ h, children }) => <section className="sec"><h2 tabIndex={-1}>{h}</h2>{children}</section>

export function FileInfo({ upload, meta }) {
  return (
    <div className="file">
      <div className="fname">{upload.fileName}</div>
      <div className="fmeta">{meta ?? `${upload.result.dataRowCount} rows | Checked in this browser at ${clockTime(upload.at)}`}</div>
    </div>
  )
}

export function ValidFile({ upload, busy = false, onChoose, h = SEC_FILE }) {
  const rows = upload.result.rows.filter((r) => r.value > 0n)
  const countries = new Set(rows.map((r) => r.country)).size
  const access = new Set(rows.filter((r) => r.indicator === 'token_access').map((r) => r.segment)).size
  const regions = new Set(rows.filter((r) => r.indicator === 'model_region_of_origin').map((r) => r.segment)).size
  return (
    <Sec h={h}>
      <FileInfo upload={upload} />
      <div className="summary"><span className="big">All {upload.result.dataRowCount} rows are valid</span></div>
      <p>Your contribution: {plural(countries, 'country', 'countries')}, {plural(access, 'model access type', 'model access types')} and {plural(regions, 'model region of origin', 'model regions of origin')}.</p>
      <div className="actions"><button className="btn btn-secondary" onClick={onChoose} disabled={busy}>Upload a different file</button></div>
    </Sec>
  )
}

function ExampleTable({ g, limit }) {
  const rows = limit ? g.rows.slice(0, limit) : g.rows
  return (
    <div className="table-scroll">
      <table className="ex">
        <thead><tr>{g.cols.map((h) => <th key={h} scope="col">{h}</th>)}</tr></thead>
        <tbody>{rows.map((r, i) => <tr key={i}>{r.map((v, j) => <td key={j}>{j === g.bad ? <mark>{v}</mark> : v}</td>)}</tr>)}</tbody>
      </table>
    </div>
  )
}

function ProblemGroup({ g, i, compact }) {
  const [all, setAll] = useState(false)
  const more = g.rows.length > 3
  const copyHeader = () => navigator.clipboard?.writeText(HEADERS.join(',')).then(() => announce('Accepted header copied.'), () => {})
  return (
    <li className="prob">
      <span className="pn">{i + 1}</span>
      <div className="pc">
        <p className="what">{g.what}</p>
        <p className="where">{g.where}</p>
        {!compact && g.accepted && <p>{g.accepted}</p>}
        {(g.code ?? []).map(([l, v]) => [<p key={`${l}l`}>{l}</p>, <p key={`${l}v`} className="code">{v}</p>])}
        {g.copyHeader && <div><button className="btn btn-secondary" onClick={copyHeader}>Copy accepted header</button></div>}
        <p>{g.fix}</p>
        {g.rows.length > 0 && (compact
          ? <details className="exrows"><summary className="link">Show example rows</summary><ExampleTable g={g} limit={3} /></details>
          : <ExampleTable g={g} limit={all ? 0 : 3} />)}
        {!compact && more && <div><button className="link" onClick={() => setAll(!all)}>{all ? 'Show fewer rows' : `Show all ${g.rows.length} rows`}</button></div>}
      </div>
    </li>
  )
}

// `labels` overrides action labels, e.g. the offline tool offers a template instead of sample.csv
export function FileCheck({ upload, onChoose, onSample, h = SEC_FILE, labels = {} }) {
  const v = upload.view
  const acts = (v.acts ?? [['choose', 'Upload corrected file'], ['report', 'Download problem report (.csv)']])
    .map(([a, label]) => [a, labels[a] ?? label])
  const run = {
    choose: onChoose,
    sample: onSample,
    report: () => download('problem-report.csv', problemReportCsv(v)),
    reload: () => window.location.reload(),
  }
  return (
    <Sec h={h}>
      <FileInfo upload={upload} meta={v.blocked ? v.meta : undefined} />
      <div className="summary">
        <div className="line">{v.blocked ? <Tag kind="blocked">Blocked</Tag> : <Tag kind="fix">Needs correction</Tag>}<span className="big">{v.big}</span></div>
        {v.sub && <p className="subl">{v.sub}</p>}
      </div>
      {(v.text ?? []).map((p, i) => <p key={i}>{p}</p>)}
      {v.groups?.length > 0 && <ol className="problems">{v.groups.map((g, i) => <ProblemGroup key={i} g={g} i={i} compact={v.compact} />)}</ol>}
      <div className="actions">
        {acts.map(([a, label], i) => <button key={a} className={`btn ${i ? 'btn-secondary' : 'btn-primary'}`} onClick={run[a]}>{label}</button>)}
      </div>
    </Sec>
  )
}
