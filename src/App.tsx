import { useEffect, useMemo, useState } from 'react'
import { useIsAuthenticated, useMsal } from '@azure/msal-react'
import {
  AlertCircle,
  ArrowUpRight,
  BarChart3,
  Building2,
  ChevronDown,
  CircleDollarSign,
  Database,
  LayoutDashboard,
  LoaderCircle,
  LogIn,
  LogOut,
  PencilLine,
  RefreshCw,
  Search,
  ShieldCheck,
  Users,
} from 'lucide-react'
import { loadDirectory, updateClient, updateClientProfile, writeClientLog } from './api'
import { isMsalConfigured, loginRequest } from './authConfig'
import {
  emptyProfile,
  mergeIndividuals,
  type ClientProfile,
  type ClientRecord,
  type HouseRecord,
  type Individual,
} from './data'

type Tab = 'dashboard' | 'editor'
const currency = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 })

function App() {
  const { instance, accounts } = useMsal()
  const authenticated = useIsAuthenticated()
  const [loginError, setLoginError] = useState('')

  const signIn = async () => {
    setLoginError('')
    try {
      const result = await instance.loginPopup(loginRequest)
      instance.setActiveAccount(result.account)
    } catch (error) {
      setLoginError(error instanceof Error ? error.message : 'Microsoft sign-in failed.')
    }
  }

  if (!isMsalConfigured) return <WelcomeScreen onSignIn={signIn} error="Microsoft Entra settings are missing from .env." disabled />
  if (!authenticated) return <WelcomeScreen onSignIn={signIn} error={loginError} />
  return <AuthenticatedApp accountName={accounts[0]?.name || accounts[0]?.username || 'C-NTA user'} />
}

function WelcomeScreen({ onSignIn, error, disabled = false }: { onSignIn: () => void; error?: string; disabled?: boolean }) {
  return <main className="welcome-screen">
    <section className="welcome-card">
      <div className="brand-mark large">C</div>
      <p className="eyebrow">C-nta · Residential services</p>
      <h1>Individual Dashboard</h1>
      <p>Sign in with your C-NTA Microsoft account to view the live client directory and residential dashboard.</p>
      <div className="welcome-security"><ShieldCheck size={17} /><span>Organizational access through Microsoft Entra ID</span></div>
      {error && <div className="inline-error"><AlertCircle size={16} />{error}</div>}
      <button className="microsoft-button" onClick={onSignIn} disabled={disabled}><LogIn size={18} /> Sign in with Microsoft</button>
    </section>
  </main>
}

function AuthenticatedApp({ accountName }: { accountName: string }) {
  const [activeTab, setActiveTab] = useState<Tab>('dashboard')
  const [clients, setClients] = useState<ClientRecord[]>([])
  const [houses, setHouses] = useState<HouseRecord[]>([])
  const [profiles, setProfileState] = useState<Record<string, ClientProfile>>({})
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const { instance, accounts } = useMsal()

  const refresh = async () => {
    setLoading(true)
    setError('')
    try {
      const directory = await loadDirectory()
      setClients(directory.clients)
      setHouses(directory.houses)
      setProfileState(directory.profiles)
    } catch (loadError) {
      setError(loadError instanceof Error ? loadError.message : 'Unable to load the C-NTA directory.')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => { void refresh() }, [])
  const individuals = useMemo(() => mergeIndividuals(clients, houses, profiles), [clients, houses, profiles])

  const saveIndividual = async (draft: Individual, previous: Individual) => {
    const { id, _id: _recordId, home, county, gender, dateOfBirth, needsGroup, diagnosisCode, weeklyRevenue, notes, ...clientPayload } = draft
    const profile: ClientProfile = { county, gender, dateOfBirth, needsGroup, diagnosisCode, weeklyRevenue, notes }
    const [apiResult, savedProfile] = await Promise.all([
      updateClient(id, clientPayload),
      updateClientProfile(id, profile),
    ])
    setClients((current) => current.map((client) => client._id === id ? { ...client, ...clientPayload, ...apiResult, _id: id } : client))
    const nextProfiles = { ...profiles, [id]: { ...profile, ...savedProfile } }
    setProfileState(nextProfiles)

    void writeClientLog({
      action: 'UPDATE',
      username: accountName,
      timestamp: new Date().toISOString(),
      clientId: id,
      previousData: previous,
      newData: draft,
      source: 'CID',
    }).catch(() => undefined)
  }

  if (loading) return <StatusScreen icon={<LoaderCircle className="spin" />} title="Loading live C-NTA data" message="Connecting to clients and houses…" />
  if (error) return <StatusScreen icon={<AlertCircle />} title="The directory could not be loaded" message={error} action={<button className="primary-button" onClick={() => void refresh()}><RefreshCw size={15} /> Try again</button>} />

  return <div className="app-shell">
    <aside className="sidebar">
      <div className="brand"><div className="brand-mark">C</div><div><strong>C-nta</strong><span>Individual Dashboard</span></div></div>
      <nav aria-label="Main navigation">
        <button className={activeTab === 'dashboard' ? 'nav-item active' : 'nav-item'} onClick={() => setActiveTab('dashboard')}><LayoutDashboard size={19} /> Dashboard</button>
        <button className={activeTab === 'editor' ? 'nav-item active' : 'nav-item'} onClick={() => setActiveTab('editor')}><PencilLine size={19} /> Individual editor</button>
      </nav>
      <div className="sidebar-foot">
        <div className="live-pill"><Database size={14} /> Live C-NTA data</div>
        <div className="user-card"><div className="avatar">{accountName[0]?.toUpperCase()}</div><div><strong>{accountName}</strong><span>{accounts[0]?.username}</span></div></div>
        <button className="auth-button" onClick={() => void instance.logoutPopup({ account: accounts[0] })}><LogOut size={16} /> Sign out</button>
      </div>
    </aside>
    <main>{activeTab === 'dashboard'
      ? <Dashboard individuals={individuals} houses={houses} onRefresh={() => void refresh()} />
      : <Editor individuals={individuals} onSave={saveIndividual} />}
    </main>
  </div>
}

function StatusScreen({ icon, title, message, action }: { icon: React.ReactNode; title: string; message: string; action?: React.ReactNode }) {
  return <main className="status-screen"><div className="status-icon">{icon}</div><h1>{title}</h1><p>{message}</p>{action}</main>
}

function getAge(date: string) {
  if (!date) return null
  const birth = new Date(`${date}T00:00:00`)
  const now = new Date()
  let age = now.getFullYear() - birth.getFullYear()
  if (now < new Date(now.getFullYear(), birth.getMonth(), birth.getDate())) age--
  return age
}

function Dashboard({ individuals, houses, onRefresh }: { individuals: Individual[]; houses: HouseRecord[]; onRefresh: () => void }) {
  const revenue = individuals.reduce((sum, person) => sum + Number(person.weeklyRevenue || 0), 0)
  const ages = individuals.map((person) => getAge(person.dateOfBirth)).filter((age): age is number => age !== null)
  const averageAge = ages.length ? Math.round(ages.reduce((sum, age) => sum + age, 0) / ages.length) : null
  const assigned = individuals.filter((person) => person.home !== 'Unassigned').length
  const profiled = individuals.filter((person) => person.county || person.needsGroup || person.diagnosisCode).length
  const homeStats = houses.map((house) => {
    const residents = individuals.filter((person) => person.home === house.houseName)
    return { name: house.houseName, count: residents.length, revenue: residents.reduce((sum, person) => sum + Number(person.weeklyRevenue || 0), 0) }
  }).filter((home) => home.count > 0).sort((a, b) => b.count - a.count).slice(0, 7)
  const hasRevenue = homeStats.some((home) => home.revenue > 0)

  return <div className="page dashboard-page">
    <header className="page-header"><div><p className="eyebrow">Residential services · Live directory</p><h1>Good morning.</h1><p>A clear view of the people, homes, and services in your care.</p></div><button className="period-button" onClick={onRefresh}><RefreshCw size={16} /> Refresh data</button></header>
    <section className="metric-grid" aria-label="Overview metrics">
      <Metric icon={<Building2 />} label="Homes in directory" value={String(houses.length)} trend={`${homeStats.length} with assignments`} />
      <Metric icon={<Users />} label="Individuals served" value={String(individuals.length)} trend={`${assigned} assigned to homes`} positive />
      <Metric icon={<CircleDollarSign />} label="Profiled weekly revenue" value={currency.format(revenue)} trend={`${profiled} CID profiles started`} />
      <Metric icon={<BarChart3 />} label="Average documented age" value={averageAge === null ? '—' : String(averageAge)} suffix={averageAge === null ? undefined : 'years'} trend={`${ages.length} birth dates entered`} />
    </section>
    <section className="dashboard-grid">
      <article className="panel revenue-panel"><div className="panel-heading"><div><span className="panel-kicker">Residential overview</span><h2>{hasRevenue ? 'Weekly revenue by home' : 'Individuals by home'}</h2></div><span className="quiet">Live assignments</span></div><HomeChart homes={homeStats} useRevenue={hasRevenue} /></article>
      <article className="panel mix-panel"><div className="panel-heading"><div><span className="panel-kicker">People</span><h2>Individuals by county</h2></div></div><CountyChart individuals={individuals} /></article>
      <article className="panel distribution-panel"><div className="panel-heading"><div><span className="panel-kicker">Support profile</span><h2>Needs group distribution</h2></div><span className="quiet">CID profiles</span></div><Distribution individuals={individuals} field="needsGroup" /></article>
      <article className="panel distribution-panel"><div className="panel-heading"><div><span className="panel-kicker">Clinical profile</span><h2>Diagnosis code distribution</h2></div><span className="quiet">CID profiles</span></div><Distribution individuals={individuals} field="diagnosisCode" /></article>
    </section>
  </div>
}

function Metric({ icon, label, value, suffix, trend, positive }: { icon: React.ReactNode; label: string; value: string; suffix?: string; trend: string; positive?: boolean }) {
  return <article className="metric-card"><div className="metric-top"><span className="metric-icon">{icon}</span><span className={positive ? 'trend up' : 'trend'}>{positive && <ArrowUpRight size={15} />}{trend}</span></div><span>{label}</span><strong>{value} {suffix && <small>{suffix}</small>}</strong></article>
}

function HomeChart({ homes, useRevenue }: { homes: { name: string; count: number; revenue: number }[]; useRevenue: boolean }) {
  if (!homes.length) return <EmptyChart message="No house assignments were returned by the API." />
  const max = Math.max(...homes.map((home) => useRevenue ? home.revenue : home.count), 1)
  return <div className="revenue-chart">{homes.map((home) => { const value = useRevenue ? home.revenue : home.count; return <div className="bar-row" key={home.name}><span>{home.name}</span><div className="bar-track"><div className="bar current-bar" style={{ width: `${(value / max) * 100}%` }} /></div><strong>{useRevenue ? currency.format(value) : value}</strong></div> })}</div>
}

function CountyChart({ individuals }: { individuals: Individual[] }) {
  const documented = individuals.filter((person) => person.county)
  if (!documented.length) return <EmptyChart message="Add county details in the Individual editor to populate this chart." />
  const cambria = documented.filter((person) => person.county === 'Cambria').length
  const percent = Math.round((cambria / documented.length) * 100)
  return <div className="donut-wrap"><div className="donut" style={{ background: `conic-gradient(var(--teal) 0 ${percent}%, var(--gold) ${percent}% 100%)` }}><div><strong>{documented.length}</strong><span>profiled</span></div></div><div className="donut-legend"><p><i className="cambria" /><span>Cambria County</span><strong>{cambria}</strong></p><p><i className="somerset" /><span>Somerset County</span><strong>{documented.length - cambria}</strong></p></div></div>
}

function Distribution({ individuals, field }: { individuals: Individual[]; field: 'needsGroup' | 'diagnosisCode' }) {
  const data = Object.entries(individuals.reduce<Record<string, number>>((acc, person) => person[field] ? ({ ...acc, [person[field]]: (acc[person[field]] || 0) + 1 }) : acc, {})).sort(([a], [b]) => a.localeCompare(b))
  if (!data.length) return <EmptyChart message={`Add ${field === 'needsGroup' ? 'needs groups' : 'diagnosis codes'} in the Individual editor to populate this chart.`} />
  const max = Math.max(...data.map(([, count]) => count))
  return <div className="distribution">{data.map(([name, count]) => <div className="distribution-row" key={name}><span>{name}</span><div><i style={{ width: `${(count / max) * 100}%` }} /></div><strong>{count}</strong></div>)}</div>
}

function EmptyChart({ message }: { message: string }) { return <div className="empty-chart"><Database size={24} /><p>{message}</p></div> }

function Editor({ individuals, onSave }: { individuals: Individual[]; onSave: (draft: Individual, previous: Individual) => Promise<void> }) {
  const [query, setQuery] = useState('')
  const [selectedId, setSelectedId] = useState(individuals[0]?.id || '')
  const selected = individuals.find((person) => person.id === selectedId) || individuals[0]
  const [draft, setDraft] = useState<Individual>(selected)
  const [saveState, setSaveState] = useState<'idle' | 'saving' | 'saved' | 'error'>('idle')
  const [saveError, setSaveError] = useState('')
  const filtered = useMemo(() => individuals.filter((person) => `${person.initials} ${person.home} ${person.id}`.toLowerCase().includes(query.toLowerCase())), [individuals, query])

  useEffect(() => { if (selected && (!draft || draft.id !== selected.id)) setDraft(selected) }, [selected, draft])
  if (!draft) return <EmptyChart message="No active client records were returned by the API." />

  const select = (person: Individual) => { setSelectedId(person.id); setDraft(person); setSaveState('idle'); setSaveError('') }
  const update = (field: keyof Individual, value: unknown) => { setDraft((current) => ({ ...current, [field]: value })); setSaveState('idle') }
  const save = async () => {
    setSaveState('saving'); setSaveError('')
    try { await onSave(draft, selected); setSaveState('saved') }
    catch (error) { setSaveError(error instanceof Error ? error.message : 'Unable to save changes.'); setSaveState('error') }
  }

  return <div className="page editor-page">
    <header className="page-header"><div><p className="eyebrow">Live client records</p><h1>Individual editor</h1><p>Update existing service fields and complete each person’s CID profile.</p></div><span className="record-count">{individuals.length} records</span></header>
    <section className="editor-layout">
      <aside className="people-list"><label className="search-box"><Search size={18} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search initials or home" /></label><div className="people-scroll">{filtered.map((person) => <button key={person.id} className={selectedId === person.id ? 'person-row selected' : 'person-row'} onClick={() => select(person)}><span className="person-avatar">{person.initials.slice(0, 2)}</span><span><strong>{person.initials}</strong><small>{person.home}</small></span><i className="status active" /></button>)}</div></aside>
      <form className="edit-form" onSubmit={(event) => { event.preventDefault(); void save() }}>
        <div className="form-title"><div className="large-avatar">{draft.initials.slice(0, 2)}</div><div><span>Client record</span><h2>{draft.initials}</h2><p>{draft.home} · ID ending {draft.id.slice(-6)}</p></div><span className="status-chip active">Live API</span></div>
        <div className="profile-notice"><Database size={16} /><span>Service fields save to the existing client record. CID profile fields save to the shared ClientProfiles store.</span></div>
        <div className="form-section"><h3>Client identity</h3><div className="field-grid">
          <Field label="Initials"><input value={draft.initials} onChange={(event) => update('initials', event.target.value.toUpperCase())} /></Field>
          <Field label="Assigned home"><input value={draft.home} disabled /></Field>
          <Field label="Address"><input value={String(draft.Address || '')} onChange={(event) => update('Address', event.target.value)} /></Field>
          <Field label="County"><select value={draft.county} onChange={(event) => update('county', event.target.value)}><option value="">Not entered</option><option>Cambria</option><option>Somerset</option></select></Field>
          <Field label="Date of birth"><input type="date" value={draft.dateOfBirth} onChange={(event) => update('dateOfBirth', event.target.value)} /></Field>
          <Field label="Gender"><select value={draft.gender} onChange={(event) => update('gender', event.target.value)}><option value="">Not entered</option><option>Female</option><option>Male</option><option>Nonbinary</option></select></Field>
        </div></div>
        <div className="form-section"><h3>Service details</h3><div className="field-grid">
          <Field label="Community code"><input value={String(draft.communityCode || draft.CommunityCode || '')} onChange={(event) => update('communityCode', event.target.value)} /></Field>
          <Field label="Community units"><input type="number" min="0" value={String(draft.CommunityUnits || '')} onChange={(event) => update('CommunityUnits', event.target.value)} /></Field>
          <Field label="Facility code"><input value={String(draft.facilityCode || draft.FacilityCode || '')} onChange={(event) => update('facilityCode', event.target.value)} /></Field>
          <Field label="Facility units"><input type="number" min="0" value={String(draft.FacilityUnits || '')} onChange={(event) => update('FacilityUnits', event.target.value)} /></Field>
          <Field label="Transportation code"><input value={String(draft.transportationCode || draft.TransportationCode || '')} onChange={(event) => update('transportationCode', event.target.value)} /></Field>
          <Field label="Transportation units"><input type="number" min="0" value={String(draft.TransportationUnits || '')} onChange={(event) => update('TransportationUnits', event.target.value)} /></Field>
          <Field label="Needs group"><select value={draft.needsGroup} onChange={(event) => update('needsGroup', event.target.value)}><option value="">Not entered</option>{['NG 1', 'NG 2', 'NG 3', 'NG 4', 'NG 5'].map((group) => <option key={group}>{group}</option>)}</select></Field>
          <Field label="Primary diagnosis code"><input value={draft.diagnosisCode} onChange={(event) => update('diagnosisCode', event.target.value)} /></Field>
          <Field label="Weekly revenue"><div className="money-input"><span>$</span><input type="number" min="0" value={draft.weeklyRevenue || ''} onChange={(event) => update('weeklyRevenue', Number(event.target.value))} /></div></Field>
        </div></div>
        <div className="form-section"><h3>Schedule and support</h3><div className="field-grid">
          <Field label="Default schedule"><input value={String(draft.DefaultSchedule || '')} onChange={(event) => update('DefaultSchedule', event.target.value)} /></Field>
          <Field label="Start time"><input type="time" value={String(draft.StartTime || '')} onChange={(event) => update('StartTime', event.target.value)} /></Field>
          <Field label="End time"><input type="time" value={String(draft.EndTime || '')} onChange={(event) => update('EndTime', event.target.value)} /></Field>
          <Field label="Scheduled days"><input value={(draft.ScheduledDays || []).join(', ')} onChange={(event) => update('ScheduledDays', event.target.value.split(',').map((day) => day.trim()).filter(Boolean))} placeholder="Monday, Tuesday" /></Field>
        </div><div className="toggle-row"><label><input type="checkbox" checked={Boolean(draft.transportation)} onChange={(event) => update('transportation', event.target.checked)} /> Transportation required</label><label><input type="checkbox" checked={Boolean(draft.AdditionalSupport)} onChange={(event) => update('AdditionalSupport', event.target.checked)} /> Additional support</label></div></div>
        <div className="form-section"><h3>Additional notes</h3><Field label="CID notes"><textarea rows={4} value={draft.notes} onChange={(event) => update('notes', event.target.value)} /></Field></div>
        <div className="form-actions">{saveState === 'saved' && <span className="save-message visible">Changes saved.</span>}{saveState === 'error' && <span className="form-error">{saveError}</span>}<button type="submit" className="primary-button" disabled={saveState === 'saving'}>{saveState === 'saving' && <LoaderCircle className="spin" size={15} />}{saveState === 'saving' ? 'Saving…' : 'Save changes'}</button></div>
      </form>
    </section>
  </div>
}

function Field({ label, children }: { label: string; children: React.ReactNode }) { return <label className="field"><span>{label}</span>{children}</label> }

export default App
