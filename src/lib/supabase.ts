import { createClient } from '@supabase/supabase-js'
import type { AuthChangeEvent, Session, Subscription } from '@supabase/supabase-js'

type EmployeeRecord = {
  workspace_key: string
  employee_id: string
  full_name: string
  department: string
  role: string
  phone: string
  email: string
  status: string
  source_file_name: string
  source_file_path: string
}

type RegistrationRecord = {
  workspace_key: string
  id: string
  employee_id: string | null
  full_name: string
  department: string
  role: string
  phone: string
  date: string
  registered_at: string
}

type PreferenceRecord = {
  workspace_key: string
  theme: string
  officer_name: string
  officer_role: string
  hotel_name: string
  daily_target: number
  workspace_label: string
}

type UploadRecord = {
  workspace_key: string
  file_name: string
  bucket_path: string
  row_count: number
}

type EmployeeInput = {
  employeeId: string
  fullName: string
  department: string
  role: string
  phone: string
  email: string
  status: string
}

type RegistrationInput = {
  id: string
  employeeId?: string
  fullName: string
  department: string
  role: string
  phone: string
  date: string
  registeredAt: string
}

type PreferencesInput = {
  theme: string
  officerName: string
  officerRole: string
  hotelName: string
  dailyTarget: number
  workspaceKey: string
}

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL?.trim() ?? ''
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY?.trim() ?? ''
const uploadBucket = import.meta.env.VITE_SUPABASE_BUCKET?.trim() || 'employee-imports'

const hasSupabaseConfig = Boolean(supabaseUrl && supabaseAnonKey)

const supabase = hasSupabaseConfig ? createClient(supabaseUrl, supabaseAnonKey) : null

export function isSupabaseEnabled() {
  return supabase !== null
}

export async function getCurrentSession() {
  if (!supabase) {
    return null
  }

  const result = await supabase.auth.getSession()
  if (result.error) {
    throw result.error
  }

  return result.data.session
}

export function onSupabaseAuthChange(
  listener: (event: AuthChangeEvent, session: Session | null) => void | Promise<void>,
): Subscription {
  if (!supabase) {
    return {
      unsubscribe: () => undefined,
    } as Subscription
  }

  const { data } = supabase.auth.onAuthStateChange(listener)
  return data.subscription
}

export async function signInWithPassword(email: string, password: string) {
  if (!supabase) {
    throw new Error('Supabase is not configured')
  }

  return supabase.auth.signInWithPassword({ email, password })
}

export async function signOutSupabaseUser() {
  if (!supabase) {
    return
  }

  const result = await supabase.auth.signOut()
  if (result.error) {
    throw result.error
  }
}

export async function ensureWorkspaceMembership(workspaceKey: string, workspaceName: string) {
  if (!supabase) {
    return
  }

  const result = await supabase.rpc('ensure_workspace_membership', {
    workspace_code: workspaceKey,
    workspace_name: workspaceName,
  })

  if (result.error) {
    throw result.error
  }
}

export async function loadWorkspaceSnapshot(workspaceKey: string) {
  if (!supabase) {
    return null
  }

  const [employeesResult, registrationsResult, preferencesResult] = await Promise.all([
    supabase
      .from('employee_directory')
      .select('*')
      .eq('workspace_key', workspaceKey)
      .order('full_name', { ascending: true }),
    supabase
      .from('registrations')
      .select('*')
      .eq('workspace_key', workspaceKey)
      .order('registered_at', { ascending: false }),
    supabase
      .from('app_preferences')
      .select('*')
      .eq('workspace_key', workspaceKey)
      .maybeSingle(),
  ])

  if (employeesResult.error) {
    throw employeesResult.error
  }

  if (registrationsResult.error) {
    throw registrationsResult.error
  }

  if (preferencesResult.error) {
    throw preferencesResult.error
  }

  return {
    employees: (employeesResult.data ?? []).map((item) => ({
      employeeId: item.employee_id as string,
      fullName: item.full_name as string,
      department: item.department as string,
      role: item.role as string,
      phone: item.phone as string,
      email: item.email as string,
      status: item.status as string,
    })),
    registrations: (registrationsResult.data ?? []).map((item) => ({
      id: item.id as string,
      employeeId: (item.employee_id as string | null) ?? undefined,
      fullName: item.full_name as string,
      department: item.department as string,
      role: item.role as string,
      phone: item.phone as string,
      date: item.date as string,
      registeredAt: item.registered_at as string,
    })),
    preferences: preferencesResult.data
      ? {
          theme: preferencesResult.data.theme as string,
          officerName: preferencesResult.data.officer_name as string,
          officerRole: preferencesResult.data.officer_role as string,
          hotelName: preferencesResult.data.hotel_name as string,
          dailyTarget: Number(preferencesResult.data.daily_target),
          workspaceKey: preferencesResult.data.workspace_label as string,
        }
      : null,
  }
}

export async function saveWorkspacePreferences(preferences: PreferencesInput) {
  if (!supabase) {
    return
  }

  const record: PreferenceRecord = {
    workspace_key: preferences.workspaceKey,
    theme: preferences.theme,
    officer_name: preferences.officerName,
    officer_role: preferences.officerRole,
    hotel_name: preferences.hotelName,
    daily_target: preferences.dailyTarget,
    workspace_label: preferences.workspaceKey,
  }

  const result = await supabase.from('app_preferences').upsert(record, {
    onConflict: 'workspace_key',
  })

  if (result.error) {
    throw result.error
  }
}

export async function saveWorkspaceEmployees(
  workspaceKey: string,
  employees: EmployeeInput[],
  sourceFileName?: string,
  sourceFile?: File,
) {
  if (!supabase) {
    return
  }

  const records: EmployeeRecord[] = employees.map((employee) => ({
    workspace_key: workspaceKey,
    employee_id: employee.employeeId,
    full_name: employee.fullName,
    department: employee.department,
    role: employee.role,
    phone: employee.phone,
    email: employee.email,
    status: employee.status,
    source_file_name: sourceFileName ?? '',
    source_file_path: '',
  }))

  const employeeResult = await supabase.from('employee_directory').upsert(records, {
    onConflict: 'workspace_key,employee_id',
  })

  if (employeeResult.error) {
    throw employeeResult.error
  }

  if (sourceFile) {
    const bucketPath = `${workspaceKey}/employee-imports/${Date.now()}-${sourceFile.name}`
    const uploadResult = await supabase.storage.from(uploadBucket).upload(bucketPath, sourceFile, {
      upsert: true,
      contentType: sourceFile.type || 'application/octet-stream',
    })

    if (!uploadResult.error) {
      const uploadRecord: UploadRecord = {
        workspace_key: workspaceKey,
        file_name: sourceFile.name,
        bucket_path: uploadResult.data?.path ?? bucketPath,
        row_count: employees.length,
      }

      const auditResult = await supabase.from('employee_uploads').insert(uploadRecord)
      if (auditResult.error) {
        throw auditResult.error
      }
    }
  }
}

export async function saveWorkspaceRegistrations(workspaceKey: string, registrations: RegistrationInput[]) {
  if (!supabase) {
    return
  }

  const records: RegistrationRecord[] = registrations.map((registration) => ({
    workspace_key: workspaceKey,
    id: registration.id,
    employee_id: registration.employeeId ?? null,
    full_name: registration.fullName,
    department: registration.department,
    role: registration.role,
    phone: registration.phone,
    date: registration.date,
    registered_at: registration.registeredAt,
  }))

  const result = await supabase.from('registrations').upsert(records, {
    onConflict: 'id',
  })

  if (result.error) {
    throw result.error
  }
}
