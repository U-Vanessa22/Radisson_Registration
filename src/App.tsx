import { useEffect, useMemo, useState } from 'react'
import type { ChangeEvent, FormEvent } from 'react'
import * as XLSX from 'xlsx'
import './App.css'
import {
  ensureWorkspaceMembership,
  isSupabaseEnabled,
  getCurrentSession,
  loadWorkspaceSnapshot,
  onSupabaseAuthChange,
  signInWithPassword,
  signOutSupabaseUser,
  saveWorkspaceEmployees,
  saveWorkspacePreferences,
  saveWorkspaceRegistrations,
} from './lib/supabase'

const departments = [
  'Front Office',
  'Housekeeping',
  'Food & Beverage',
  'Kitchen',
  'Finance',
  'People and Culture',
  'Sales & Marketing',
  'Engineering',
  'Security',
  'Spa & Leisure',
  'Casuals and Trainees',
  'JOTETE',
  'PTS',
  'FCC',
  'OTIS',
  'IT',
] as const

type Department = (typeof departments)[number]
type Page = 'home' | 'registrations' | 'settings'
type Theme = 'light' | 'dark'

type Registration = {
  id: string
  employeeId?: string
  fullName: string
  department: Department
  phone: string
  date: string
  registeredAt: string
  /* role: string */
}

type RegistrationForm = Omit<Registration, 'id' | 'registeredAt'>

type Employee = {
  employeeId: string
  fullName: string
  department: Department
  /* role: string */
  phone: string
  email: string
  status: string
}

type Preferences = {
  theme: Theme
  officerName: string
  officerRole: string
  hotelName: string
  dailyTarget: number
  workspaceKey: string
}

const employeeBase: Record<Department, number> = {
  'Front Office': 30,
  Housekeeping: 60,
  'Food & Beverage': 80,
  Kitchen: 90,
  Finance: 20,
  'People and Culture': 6,
  'Sales & Marketing': 14,
  Engineering: 50,
  Security: 45,
  'Spa & Leisure': 7,
  'Casuals and Trainees': 15,
  'JOTETE': 10,
  'PTS': 25,
  'FCC': 20,
  'OTIS': 15,
  'IT': 12,
}

const defaultPreferences: Preferences = {
  theme: 'light',
  officerName: 'HR Officer',
  officerRole: 'People & Culture',
  hotelName: 'Radisson Blu Hotel',
  dailyTarget: 18,
  workspaceKey: 'radisson-registration',
}

const emptyForm = (date: string): RegistrationForm => ({
  fullName: '',
  department: 'Front Office',
 /* role: '',*/
  phone: '',
  date,
})

function getTodayInput() {
  const now = new Date()
  const local = new Date(now.getTime() - now.getTimezoneOffset() * 60000)
  return local.toISOString().slice(0, 10)
}

function getCurrentTime() {
  return new Intl.DateTimeFormat('en', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).format(new Date())
}

function formatDate(value: string) {
  return new Intl.DateTimeFormat('en', {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
  }).format(new Date(`${value}T00:00:00`))
}

function escapeHtml(value: string) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
}

function escapePdfText(value: string) {
  return value.replaceAll('\\', '\\\\').replaceAll('(', '\\(').replaceAll(')', '\\)')
}

function loadRegistrations() {
  const saved = localStorage.getItem('radisson-daily-registrations')
  if (!saved) {
    return []
  }

  try {
    const parsed = JSON.parse(saved) as Registration[]
    return Array.isArray(parsed) ? parsed : []
  } catch {
    return []
  }
}

function loadPreferences() {
  const saved = localStorage.getItem('radisson-registration-preferences')
  if (!saved) {
    return defaultPreferences
  }

  try {
    return { ...defaultPreferences, ...(JSON.parse(saved) as Partial<Preferences>) }
  } catch {
    return defaultPreferences
  }
}

function loadEmployees() {
  const saved = localStorage.getItem('radisson-employee-directory')
  if (!saved) {
    return []
  }

  try {
    const parsed = JSON.parse(saved) as Employee[]
    return Array.isArray(parsed) ? parsed : []
  } catch {
    return []
  }
}

function normalizeColumn(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]/g, '')
}

function getCell(row: Record<string, unknown>, columns: string[]) {
  const wanted = columns.map(normalizeColumn)
  const key = Object.keys(row).find((item) => wanted.includes(normalizeColumn(item)))
  return key ? String(row[key] ?? '').trim() : ''
}

function normalizeDepartment(value: string): Department {
  const found = departments.find(
    (department) => normalizeColumn(department) === normalizeColumn(value),
  )
  return found ?? 'Front Office'
}

function parseEmployeeRows(rows: Record<string, unknown>[]) {
  return rows
    .map((row, index): Employee | null => {
      const fullName = getCell(row, ['Full Name', 'Name', 'Employee Name', 'Employee'])
      if (!fullName) {
        return null
      }

      return {
        employeeId:
          getCell(row, ['Employee ID', 'EmployeeId', 'Staff ID', 'StaffId', 'ID', 'Code']) ||
          `EMP-${String(index + 1).padStart(4, '0')}`,
        fullName,
        department: normalizeDepartment(getCell(row, ['Department', 'Dept', 'Division'])),
        /*role: getCell(row, ['Role', 'Position', 'Job Title', 'Title']) || 'Team Member',*/
        phone: getCell(row, ['Phone', 'Telephone', 'Mobile', 'Contact']),
        email: getCell(row, ['Email', 'Email Address', 'Work Email']),
        status: getCell(row, ['Status', 'Employee Status']) || 'Active',
      }
    })
    .filter((employee): employee is Employee => employee !== null)
}

function downloadBlob(filename: string, content: BlobPart, type: string) {
  const blob = new Blob([content], { type })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = filename
  document.body.append(anchor)
  anchor.click()
  anchor.remove()
  URL.revokeObjectURL(url)
}

function buildExcelReport(registrations: Registration[], date: string, preferences: Preferences) {
  const rows = registrations
    .map(
      (item, index) => `
        <tr>
          <td>${index + 1}</td>
          <td>${escapeHtml(item.employeeId ?? '')}</td>
          <td>${escapeHtml(item.fullName)}</td>
          <td>${escapeHtml(item.department)}</td>
          <td>${escapeHtml(item.phone)}</td>
          <td>${escapeHtml(item.registeredAt)}</td>
        </tr>`,
    )
    .join('')

  return `
    <html>
      <head>
        <meta charset="utf-8" />
        <style>
          body { font-family: Arial, sans-serif; }
          h1 { color: #13294b; }
          table { border-collapse: collapse; width: 100%; }
          th { background: #13294b; color: #fff; }
          th, td { border: 1px solid #d9dee8; padding: 8px; text-align: left; }
        </style>
      </head>
      <body>
        <h1>${escapeHtml(preferences.hotelName)} Daily Registration Report</h1>
        <p>Date: ${escapeHtml(formatDate(date))}</p>
        <p>Prepared by: ${escapeHtml(preferences.officerName)}</p>
        <p>Total registrations: ${registrations.length}</p>
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Employee ID</th>
              <th>Full name</th>
              <th>Department</th>
              <th>Phone</th>
              <th>Registered at</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </body>
    </html>`
}

function buildPdfReport(registrations: Registration[], date: string, preferences: Preferences) {
  const lines = [
    `${preferences.hotelName} Daily Registration Report`,
    `Date: ${formatDate(date)}`,
    `Prepared by: ${preferences.officerName}`,
    `Total registrations: ${registrations.length}`,
    '',
    'Employee ID | Name | Department | Time',
    ...registrations.map(
      (item) =>
        `${item.employeeId ?? '-'} | ${item.fullName} | ${item.department} || 'Team Member'} | ${item.registeredAt}`,
    ),
  ]

  const pageLines = 35
  const pages = Array.from(
    { length: Math.max(1, Math.ceil(lines.length / pageLines)) },
    (_, index) => lines.slice(index * pageLines, (index + 1) * pageLines),
  )

  const objects: string[] = []
  const addObject = (body: string) => {
    objects.push(body)
    return objects.length
  }

  const fontObject = addObject('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')
  const pageObjectIds: number[] = []

  pages.forEach((page, pageIndex) => {
    const stream = [
      'BT',
      '/F1 13 Tf',
      '54 548 Td',
      ...page.map((line, lineIndex) => {
        const prefix = lineIndex === 0 ? '' : '0 -15 Td '
        return `${prefix}(${escapePdfText(line.slice(0, 92))}) Tj`
      }),
      '0 -28 Td',
      `/F1 9 Tf (Page ${pageIndex + 1} of ${pages.length}) Tj`,
      'ET',
    ].join('\n')
    const contentObject = addObject(
      `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    )
    const pageObject = addObject(
      `<< /Type /Page /Parent 0 0 R /MediaBox [0 0 792 612] /Resources << /Font << /F1 ${fontObject} 0 R >> >> /Contents ${contentObject} 0 R >>`,
    )
    pageObjectIds.push(pageObject)
  })

  const pagesObject = addObject(
    `<< /Type /Pages /Kids [${pageObjectIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${
      pageObjectIds.length
    } >>`,
  )
  const catalogObject = addObject(`<< /Type /Catalog /Pages ${pagesObject} 0 R >>`)

  pageObjectIds.forEach((pageId) => {
    objects[pageId - 1] = objects[pageId - 1].replace('/Parent 0 0 R', `/Parent ${pagesObject} 0 R`)
  })

  const chunks = ['%PDF-1.4\n']
  const offsets = [0]
  objects.forEach((object, index) => {
    offsets.push(chunks.join('').length)
    chunks.push(`${index + 1} 0 obj\n${object}\nendobj\n`)
  })
  const xrefOffset = chunks.join('').length
  chunks.push(`xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`)
  offsets.slice(1).forEach((offset) => {
    chunks.push(`${String(offset).padStart(10, '0')} 00000 n \n`)
  })
  chunks.push(
    `trailer\n<< /Size ${objects.length + 1} /Root ${catalogObject} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`,
  )

  return chunks.join('')
}

function App() {
  const today = getTodayInput()
  const [page, setPage] = useState<Page>('home')
  const [selectedDate, setSelectedDate] = useState(today)
  const [registrations, setRegistrations] = useState<Registration[]>(loadRegistrations)
  const [employees, setEmployees] = useState<Employee[]>(loadEmployees)
  const [form, setForm] = useState<RegistrationForm>(emptyForm(today))
  const [preferences, setPreferences] = useState<Preferences>(loadPreferences)
  const [isSignedIn, setIsSignedIn] = useState(() => localStorage.getItem('radisson-session') === 'active')
  const [loginName, setLoginName] = useState(preferences.officerName)
  const [loginEmail, setLoginEmail] = useState('')
  const [loginCode, setLoginCode] = useState('')
  const [employeeSearch, setEmployeeSearch] = useState('')
  const [employeeSearchType, setEmployeeSearchType] = useState<'all' | 'name' | 'department'>('all')
  const [selectedEmployeeId, setSelectedEmployeeId] = useState('')
  const [registrationNotice, setRegistrationNotice] = useState('')
  const [importNotice, setImportNotice] = useState('')
  const [syncNotice, setSyncNotice] = useState('')
  const [syncState, setSyncState] = useState(isSupabaseEnabled() ? 'Connecting to Supabase...' : 'Offline mode')

  useEffect(() => {
    let cancelled = false

    const bootstrapAuth = async () => {
      if (!isSupabaseEnabled()) {
        setIsSignedIn(true)
        return
      }

      const session = await getCurrentSession()
      if (cancelled) {
        return
      }

      setIsSignedIn(Boolean(session))
    }

    const hydrateWorkspace = async () => {
      if (!isSupabaseEnabled()) {
        setSyncState('Offline mode')
        return
      }

      setSyncState('Loading workspace data from Supabase...')

      try {
        const snapshot = await loadWorkspaceSnapshot(preferences.workspaceKey)
        if (cancelled || !snapshot) {
          return
        }

        if (snapshot.preferences) {
          const nextPreferences = { ...preferences, ...snapshot.preferences } as Preferences
          setPreferences(nextPreferences)
          localStorage.setItem('radisson-registration-preferences', JSON.stringify(nextPreferences))
          setLoginName(nextPreferences.officerName)
        }

        if (snapshot.employees.length > 0) {
          setEmployees(snapshot.employees as Employee[])
          localStorage.setItem('radisson-employee-directory', JSON.stringify(snapshot.employees))
        }

        if (snapshot.registrations.length > 0) {
          setRegistrations(snapshot.registrations as Registration[])
          localStorage.setItem('radisson-daily-registrations', JSON.stringify(snapshot.registrations))
        }

        setSyncState('Connected to Supabase')
      } catch (error) {
        if (cancelled) {
          return
        }

        console.error('Supabase sync failed', error)
        setSyncState('Supabase sync failed, using local storage')
      }
    }

    void bootstrapAuth()
    void hydrateWorkspace()

    return () => {
      cancelled = true
    }
  }, [])

  useEffect(() => {
    if (!isSupabaseEnabled()) {
      return
    }

    const subscription = onSupabaseAuthChange((_event, session) => {
      setIsSignedIn(Boolean(session))
    })

    return () => {
      subscription.unsubscribe()
    }
  }, [])

  const dailyRegistrations = useMemo(
    () => registrations.filter((item) => item.date === selectedDate),
    [registrations, selectedDate],
  )

  const departmentTotals = useMemo(
    () =>
      departments.map((department) => ({
        department,
        total:
          employeeBase[department] +
          registrations.filter((item) => item.department === department).length,
      })),
    [registrations],
  )

  const dailyDepartmentTotals = useMemo(
    () =>
      departments.map((department) => ({
        department,
        total: dailyRegistrations.filter((item) => item.department === department).length,
      })),
    [dailyRegistrations],
  )

  const filteredEmployees = useMemo(() => {
    const query = employeeSearch.trim().toLowerCase()
    if (!query) {
      return employees.slice(0, 5)
    }

    if (employeeSearchType === 'name') {
      return employees
        .filter((employee) => employee.fullName.toLowerCase().includes(query))
        .slice(0, 6)
    }

    if (employeeSearchType === 'department') {
      return employees
        .filter((employee) => employee.department.toLowerCase().includes(query))
        .slice(0, 6)
    }

    return employees
      .filter((employee) =>
        [employee.employeeId, employee.fullName, employee.department, {/*employee.role*/}]
          .join(' ')
          .toLowerCase()
          .includes(query),
      )
      .slice(0, 6)
  }, [employeeSearch, employeeSearchType, employees])

  const totalEmployees = departmentTotals.reduce((sum, item) => sum + item.total, 0)
  const busiestDepartment = departmentTotals.reduce((winner, item) =>
    item.total > winner.total ? item : winner,
  )
  const maxDepartmentTotal = Math.max(...departmentTotals.map((item) => item.total))
  const todayTotal = dailyRegistrations.length
  const targetProgress = Math.min(100, Math.round((todayTotal / preferences.dailyTarget) * 100))

  const pieGradient = useMemo(() => {
    if (todayTotal === 0) {
      return 'var(--chart-empty)'
    }

    let cursor = 0
    const colors = ['#13294b', '#c69c50', '#007f7a', '#b83a4b', '#5f6f89', '#243b67']
    const stops = dailyDepartmentTotals
      .filter((item) => item.total > 0)
      .map((item, index) => {
        const start = cursor
        cursor += (item.total / todayTotal) * 100
        const color = colors[index % colors.length]
        return `${color} ${start}% ${cursor}%`
      })

    return `conic-gradient(${stops.join(', ')})`
  }, [dailyDepartmentTotals, todayTotal])

  const pageTitle = {
    home: 'Hotel staff overview',
    registrations: 'Registration desk',
    settings: 'System settings',
  }[page]

  const saveRegistrations = (nextRegistrations: Registration[]) => {
    setRegistrations(nextRegistrations)
    localStorage.setItem('radisson-daily-registrations', JSON.stringify(nextRegistrations))
    void saveWorkspaceRegistrations(preferences.workspaceKey, nextRegistrations).catch((error) => {
      console.error('Failed to sync registrations to Supabase', error)
      setSyncNotice('Registrations saved locally. Supabase sync failed.')
    })
  }

  const saveEmployees = (nextEmployees: Employee[], sourceFile?: File, sourceFileName?: string) => {
    setEmployees(nextEmployees)
    localStorage.setItem('radisson-employee-directory', JSON.stringify(nextEmployees))
    void saveWorkspaceEmployees(preferences.workspaceKey, nextEmployees, sourceFileName, sourceFile).catch((error) => {
      console.error('Failed to sync employees to Supabase', error)
      setSyncNotice('Employees saved locally. Supabase sync failed.')
    })
  }

  const savePreferences = (nextPreferences: Preferences) => {
    setPreferences(nextPreferences)
    localStorage.setItem('radisson-registration-preferences', JSON.stringify(nextPreferences))
    void saveWorkspacePreferences(nextPreferences).catch((error) => {
      console.error('Failed to sync preferences to Supabase', error)
      setSyncNotice('Preferences saved locally. Supabase sync failed.')
    })
  }

  const updatePreference = <Key extends keyof Preferences>(key: Key, value: Preferences[Key]) => {
    savePreferences({ ...preferences, [key]: value })
  }

  const applyEmployee = (employee: Employee) => {
    setSelectedEmployeeId(employee.employeeId)
    setEmployeeSearch(`${employee.employeeId} - ${employee.fullName}`)
    setForm({
      ...form,
      employeeId: employee.employeeId,
      fullName: employee.fullName,
      department: employee.department,
      /*role: employee.role,*/
      phone: employee.phone,
    })
    setRegistrationNotice('')
  }

  const handleSearchClick = () => {
    setEmployeeSearch((s) => s.trim())
  }

  const handleEmployeeImport = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }

    const data = await file.arrayBuffer()
    const workbook = XLSX.read(data, { type: 'array' })
    const rows = workbook.SheetNames.flatMap((sheetName) => {
      const sheet = workbook.Sheets[sheetName]
      return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
        defval: '',
      })
    })
    const importedEmployees = parseEmployeeRows(rows)

    if (importedEmployees.length === 0) {
      const firstSheetName = workbook.SheetNames[0] ?? 'Unknown sheet'
      const firstSheetRows = workbook.SheetNames.length > 0
        ? XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[firstSheetName], {
            defval: '',
          })
        : []
      const detectedColumns = firstSheetRows[0] ? Object.keys(firstSheetRows[0]).join(', ') : 'No columns detected'

      setImportNotice(
        `No employees were found. We looked through ${workbook.SheetNames.length} sheet(s). Detected columns: ${detectedColumns}. Expected a name column such as Full Name or Employee Name.`,
      )
      event.target.value = ''
      return
    }

    saveEmployees(importedEmployees, file, file.name)
    setImportNotice(`${importedEmployees.length} employees imported from ${file.name}.`)
    event.target.value = ''
  }

  const handleLogin = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const signIn = async () => {
      if (!isSupabaseEnabled()) {
        const nextPreferences = {
          ...preferences,
          officerName: loginName.trim() || 'HR Officer',
        }
        savePreferences(nextPreferences)
        localStorage.setItem('radisson-session', 'active')
        setIsSignedIn(true)
        setLoginCode('')
        return
      }

      const nextEmail = loginEmail.trim().toLowerCase()
      if (!nextEmail || !loginCode.trim()) {
        setSyncNotice('Enter both email and password to sign in.')
        return
      }

      const result = await signInWithPassword(nextEmail, loginCode)
      if (result.error) {
        throw result.error
      }

      await ensureWorkspaceMembership(preferences.workspaceKey, preferences.hotelName)
      setLoginCode('')
      setIsSignedIn(true)
    }

    void signIn().catch((error) => {
      console.error('Sign in failed', error)
      setSyncNotice('Sign in failed. Check your Supabase email and password.')
    })
  }

  const handleLogout = () => {
    void signOutSupabaseUser()
      .catch((error) => {
        console.error('Sign out failed', error)
      })
      .finally(() => {
        localStorage.removeItem('radisson-session')
        setIsSignedIn(false)
        setPage('home')
      })
  }

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const duplicateEmployee = registrations.some(
      (item) => {
        if (item.date !== form.date) {
          return false
        }

        return (
          (selectedEmployeeId !== '' && item.employeeId === selectedEmployeeId) ||
          item.fullName.toLowerCase() === form.fullName.trim().toLowerCase()
        )
      },
    )

    if (duplicateEmployee) {
      setRegistrationNotice('This employee is already registered for the selected date.')
      return
    }

    const nextRegistration: Registration = {
      ...form,
      employeeId: selectedEmployeeId || form.employeeId,
      id: crypto.randomUUID(),
      registeredAt: getCurrentTime(),
    }
    saveRegistrations([nextRegistration, ...registrations])
    setSelectedDate(form.date)
    setForm(emptyForm(form.date))
    setSelectedEmployeeId('')
    setEmployeeSearch('')
    setRegistrationNotice('Registration saved successfully.')
  }

  const removeRegistration = (id: string) => {
    saveRegistrations(registrations.filter((item) => item.id !== id))
  }

  const downloadExcel = () => {
    downloadBlob(
      `radisson-daily-registration-${selectedDate}.xls`,
      buildExcelReport(dailyRegistrations, selectedDate, preferences),
      'application/vnd.ms-excel;charset=utf-8',
    )
  }

  const downloadPdf = () => {
    downloadBlob(
      `radisson-daily-registration-${selectedDate}.pdf`,
      buildPdfReport(dailyRegistrations, selectedDate, preferences),
      'application/pdf',
    )
  }

  if (!isSignedIn) {
    return (
      <main className="auth-page" data-theme={preferences.theme}>
        <section className="auth-hero">
          <div className="brand-lockup" aria-label="Radisson Blu HR Registration">
            <img src="/Logo.jpeg" className="brand-mark" alt="Radisson Blu logo" />
            <div>
              <strong>{preferences.hotelName}</strong>
              <span>HR Registration</span>
            </div>
          </div>
          <div className="auth-copy">
            <p className="eyebrow">Daily people operations</p>
            <h1>Welcome back to the registration desk</h1>
            <p>
              Sign in to record daily employees, monitor department totals, and prepare the
              report for HR leadership.
            </p>
          </div>
        </section>

        <form className="login-panel" onSubmit={handleLogin}>
          <div>
            <p className="eyebrow">Secure access</p>
            <h2>Officer login</h2>
          </div>
          {isSupabaseEnabled() ? (
            <label>
              Email
              <input
                required
                type="email"
                value={loginEmail}
                onChange={(event) => setLoginEmail(event.target.value)}
                placeholder="name@hotel.com"
              />
            </label>
          ) : null}
          <label>
            Officer name
            <input
              required={!isSupabaseEnabled()}
              value={loginName}
              onChange={(event) => setLoginName(event.target.value)}
              placeholder="Enter your name"
            />
          </label>
          <label>
            {isSupabaseEnabled() ? 'Password' : 'Access code'}
            <input
              required
              type="password"
              value={loginCode}
              onChange={(event) => setLoginCode(event.target.value)}
              placeholder={isSupabaseEnabled() ? 'Enter password' : 'Enter access code'}
            />
          </label>
          <button type="submit" className="primary-button">
            Sign in
          </button>
          <p className="login-note">
            {isSupabaseEnabled() ? syncState : 'Offline mode'}
          </p>
          {/*<p className="login-note">Demo mode accepts any access code.</p>*/}
        </form>
      </main>
    )
  }

  return (
    <main className="app-shell" data-theme={preferences.theme}>
      <aside className="sidebar">
        <div className="brand-lockup" aria-label="Radisson Blu HR Registration">
          <img src="/Logo.jpeg" className="brand-mark" alt="Radisson Blu logo" />
          <div>
            <strong>{preferences.hotelName}</strong>
            <span>HR Registration</span>
          </div>
        </div>

        <nav className="nav-tabs" aria-label="Application pages">
          <button
            className={page === 'home' ? 'active' : ''}
            type="button"
            onClick={() => setPage('home')}
          >
            Dashboard
          </button>
          <button
            className={page === 'registrations' ? 'active' : ''}
            type="button"
            onClick={() => setPage('registrations')}
          >
            Registrations
          </button>
          <button
            className={page === 'settings' ? 'active' : ''}
            type="button"
            onClick={() => setPage('settings')}
          >
            Settings
          </button>
        </nav>

        <div className="sidebar-note">
          <span>Today</span>
          <strong>{formatDate(today)}</strong>
          <p>{todayTotal} registrations selected for the daily report.</p>
        </div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">Daily employee registration</p>
            <h1>{pageTitle}</h1>
          </div>
          <div className="topbar-actions">
            <div className="date-filter">
              <label htmlFor="report-date">Report date</label>
              <input
                id="report-date"
                type="date"
                value={selectedDate}
                onChange={(event) => setSelectedDate(event.target.value)}
              />
            </div>
            <div className="user-badge">
              <span>{preferences.officerName.slice(0, 1).toUpperCase()}</span>
              <div>
                <strong>{preferences.officerName}</strong>
                {/*<p>{preferences.officerRole}</p>*/}
              </div>
            </div>
            <button type="button" className="quiet-button" onClick={handleLogout}>
              Sign out
            </button>
          </div>
        </header>

        {syncNotice ? <p className="workspace-sync-notice">{syncNotice}</p> : null}

        {page === 'home' ? (
          <section className="page-grid">
            <div className="metric-card dark-card">
              <span>Total employees</span>
              <strong>{totalEmployees}</strong>
              <p>Across {departments.length} hotel departments</p>
            </div>
            <div className="metric-card">
              <span>Daily registrations</span>
              <strong>{todayTotal}</strong>
              <p>{formatDate(selectedDate)} report total</p>
            </div>
            <div className="metric-card">
              <span>Daily target</span>
              <strong>{targetProgress}%</strong>
              <p>{todayTotal} of {preferences.dailyTarget} expected entries</p>
            </div>

            <section className="panel wide-panel">
              <div className="panel-heading">
                <div>
                  <p className="eyebrow">Database by department</p>
                  <h2>Employee distribution</h2>
                </div>
                <button type="button" className="ghost-button" onClick={() => setPage('registrations')}>
                  + Add person
                </button>
              </div>

              <div className="bar-chart" aria-label="Employee totals by department">
                {departmentTotals.map((item) => (
                  <div className="bar-row" key={item.department}>
                    <span>{item.department}</span>
                    <div className="bar-track">
                      <div
                        className="bar-fill"
                        style={{ width: `${(item.total / maxDepartmentTotal) * 100}%` }}
                      />
                    </div>
                    <strong>{item.total}</strong>
                  </div>
                ))}
              </div>
            </section>

            <section className="panel pie-panel">
              <div className="panel-heading">
                <div>
                  <p className="eyebrow">Daily registration</p>
                  <h2>Today by department</h2>
                </div>
              </div>
              <div className="pie-chart" style={{ background: pieGradient }}>
                <div>
                  <strong>{todayTotal}</strong>
                  <span>Total</span>
                </div>
              </div>
              <div className="legend-list">
                {dailyDepartmentTotals
                  .filter((item) => item.total > 0)
                  .slice(0, 6)
                  .map((item, index) => (
                    <div className="legend-item" key={item.department}>
                      <span className={`legend-dot dot-${index + 1}`} />
                      <p>{item.department}</p>
                      <strong>{item.total}</strong>
                    </div>
                  ))}
                {todayTotal === 0 ? <p className="empty-note">No registration yet for this date.</p> : null}
              </div>
            </section>

            <section className="panel insight-panel">
              <div>
                <p className="eyebrow">Quick insight</p>
                <h2>{busiestDepartment.department}</h2>
                <p>
                  This is currently the largest department in the employee database with
                  {` ${busiestDepartment.total} `}people.
                </p>
              </div>
              <div className="progress-track">
                <div style={{ width: `${targetProgress}%` }} />
              </div>
            </section>
          </section>
        ) : null}

        {page === 'registrations' ? (
          <section className="registration-layout">
            <form className="panel registration-form" onSubmit={handleSubmit}>
              <div className="form-title">
                <div className="plus-badge">+</div>
                <div>
                  <p className="eyebrow">New daily entry</p>
                  <h2>Add a person</h2>
                </div>
              </div>

              <div className="employee-picker">
                <label>
                  Search employee database
                  <div className="search-row">
                    <input
                      value={employeeSearch}
                      onChange={(event) => {
                        setEmployeeSearch(event.target.value)
                        setSelectedEmployeeId('')
                      }}
                      placeholder="Search by name, ID, department"
                    />
                    <select
                      value={employeeSearchType}
                      onChange={(e) => setEmployeeSearchType(e.target.value as any)}
                      aria-label="Search by"
                    >
                      <option value="all">All</option>
                      <option value="name">Display name</option>
                      <option value="department">Department</option>
                    </select>
                    <button type="button" className="quiet-button" onClick={handleSearchClick}>
                      Search
                    </button>
                  </div>
                </label>
                <div className="employee-results">
                  {filteredEmployees.map((employee, idx) => {
                    const resultNumber = idx + 1
                    const isRegistered = registrations.some(
                      (r) => r.date === selectedDate && r.employeeId === employee.employeeId,
                    )

                    return (
                      <div key={employee.employeeId} className="employee-result-row">
                        <label>
                          <input
                            type="checkbox"
                            checked={isRegistered}
                            onChange={() => {
                              // toggle registration for this employee for the selected date
                              if (isRegistered) {
                                const remaining = registrations.filter(
                                  (r) => !(r.date === selectedDate && r.employeeId === employee.employeeId),
                                )
                                saveRegistrations(remaining)
                              } else {
                                const nextRegistration: Registration = {
                                  id: crypto.randomUUID(),
                                  employeeId: employee.employeeId,
                                  fullName: employee.fullName,
                                  department: employee.department,
                                 /* role: employee.role,*/
                                  phone: employee.phone,
                                  date: selectedDate,
                                  registeredAt: getCurrentTime(),
                                }
                                saveRegistrations([nextRegistration, ...registrations])
                              }
                            }}
                          />
                        </label>

                        <div className="result-index">{resultNumber}</div>
                        <button type="button" onClick={() => applyEmployee(employee)}>
                          <strong>{employee.fullName}</strong>
                          <small>{employee.department}</small>
                        </button>
                      </div>
                    )
                  })}
                </div>
              </div>

              <label>
                Full name
                <input
                  required
                  value={form.fullName}
                  onChange={(event) => setForm({ ...form, fullName: event.target.value })}
                  placeholder="Enter employee name"
                />
              </label>

              <label>
                Department
                <select
                  value={form.department}
                  onChange={(event) =>
                    setForm({ ...form, department: event.target.value as Department })
                  }
                >
                  {departments.map((department) => (
                    <option key={department}>{department}</option>
                  ))}
                </select>
              </label>

              {/*<label>
                Role
                <input
                  required
                  value={form.role}
                  onChange={(event) => setForm({ ...form, role: event.target.value })}
                  placeholder="Example: Receptionist"
                />
              </label>*/}

              <div className="form-row">
                <label>
                  Phone
                  <input
                    value={form.phone}
                    onChange={(event) => setForm({ ...form, phone: event.target.value })}
                    placeholder="+250 ..."
                  />
                </label>
              </div>

              <label>
                Registration date
                <input
                  type="date"
                  value={form.date}
                  onChange={(event) => setForm({ ...form, date: event.target.value })}
                />
              </label>

              <button type="submit" className="primary-button">
                + Save registration
              </button>
              {registrationNotice ? <p className="form-notice">{registrationNotice}</p> : null}
            </form>

            <section className="panel report-panel">
              <div className="panel-heading">
                <div>
                  <p className="eyebrow">Daily report</p>
                  <h2>{formatDate(selectedDate)}</h2>
                </div>
                <div className="export-actions">
                  <button type="button" onClick={downloadExcel}>
                    Excel
                  </button>
                  <button type="button" onClick={downloadPdf}>
                    PDF
                  </button>
                </div>
              </div>

              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>ID</th>
                      <th>Name</th>
                      <th>Department</th>
                      {/*<th>Role</th>*/}
                      <th>Time</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dailyRegistrations.map((item) => (
                      <tr key={item.id}>
                        <td>{item.employeeId ?? '-'}</td>
                        <td>{item.fullName}</td>
                        <td>{item.department}</td>
                        {/*<td>{item.role}</td>*/}
                        <td>{item.registeredAt}</td>
                        <td>
                          <button
                            type="button"
                            className="text-button"
                            onClick={() => removeRegistration(item.id)}
                          >
                            Remove
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {dailyRegistrations.length === 0 ? (
                  <div className="empty-state">
                    <strong>No registrations for this date</strong>
                    <p>Add a person or choose another report date.</p>
                  </div>
                ) : null}
              </div>
            </section>
          </section>
        ) : null}

        {page === 'settings' ? (
          <section className="settings-layout">
            <section className="panel settings-panel">
              <div className="panel-heading">
                <div>
                  <p className="eyebrow">Appearance</p>
                  <h2>Theme mode</h2>
                </div>
              </div>
              <div className="theme-switcher" aria-label="Theme mode">
                <button
                  type="button"
                  className={preferences.theme === 'light' ? 'active' : ''}
                  onClick={() => updatePreference('theme', 'light')}
                >
                  Light
                </button>
                <button
                  type="button"
                  className={preferences.theme === 'dark' ? 'active' : ''}
                  onClick={() => updatePreference('theme', 'dark')}
                >
                  Dark
                </button>
              </div>
              <p className="settings-copy">
                Your theme is saved on this computer, so the app opens with the same look next
                time.
              </p>
            </section>

            <section className="panel settings-panel">
              <div className="panel-heading">
                <div>
                  <p className="eyebrow">Officer profile</p>
                  <h2>Report identity</h2>
                </div>
              </div>
              <div className="settings-form">
                <label>
                  Hotel name
                  <input
                    value={preferences.hotelName}
                    onChange={(event) => updatePreference('hotelName', event.target.value)}
                  />
                </label>
                <label>
                  Officer name
                  <input
                    value={preferences.officerName}
                    onChange={(event) => updatePreference('officerName', event.target.value)}
                  />
                </label>
                {/*<label>
                  Officer role
                  <input
                    value={preferences.officerRole}
                    onChange={(event) => updatePreference('officerRole', event.target.value)}
                  />
                </label>*/ }
                <label>
                  Daily target
                  <input
                    min="1"
                    type="number"
                    value={preferences.dailyTarget}
                    onChange={(event) =>
                      updatePreference('dailyTarget', Math.max(1, Number(event.target.value)))
                    }
                  />
                </label>
                <label>
                  Workspace code
                  <input
                    value={preferences.workspaceKey}
                    onChange={(event) => updatePreference('workspaceKey', event.target.value.trim())}
                    placeholder="Shared hotel workspace code"
                  />
                </label>
              </div>
            </section>

            <section className="panel settings-panel directory-panel">
              <div className="panel-heading">
                <div>
                  <p className="eyebrow">Employee database</p>
                  <h2>Import directory</h2>
                </div>
                <strong className="directory-count">{employees.length} employees</strong>
              </div>
              <label className="import-box">
                Upload Excel or CSV
                <input type="file" accept=".xlsx,.xls,.csv" onChange={handleEmployeeImport} />
              </label>
              <p className="settings-copy">
                Expected columns include Employee ID, Full Name, Department, Role, Phone,
                Email, and Status. Similar column names also work.
              </p>
              {importNotice ? <p className="form-notice">{importNotice}</p> : null}
              <div className="directory-preview">
                {employees.slice(0, 5).map((employee) => (
                  <div key={employee.employeeId}>
                    <span>{employee.employeeId}</span>
                    <strong>{employee.fullName}</strong>
                    {/*<p>{employee.department} - {employee.role}</p>*/}
                  </div>
                ))}
              </div>
            </section>

            <section className="panel settings-panel accent-panel">
              <div>
                <p className="eyebrow">Data tools</p>
                <h2>Demo records</h2>
                <p className="settings-copy">
                  Keep the dashboard empty until you add registrations or import employees.
                </p>
              </div>
            </section>
          </section>
        ) : null}
      </section>
    </main>
  )
}

export default App
