import { emptyProfile, type ClientProfile, type ClientRecord, type HouseRecord } from './data'

const configuredBaseUrl = import.meta.env.VITE_API_BASE_URL?.trim()
const apiBaseUrl = configuredBaseUrl && !configuredBaseUrl.includes('your-existing-api')
  ? configuredBaseUrl.replace(/\/$/, '')
  : '/cntat-api'

async function apiRequest<T>(path: string, options: RequestInit = {}): Promise<T> {
  const response = await fetch(`${apiBaseUrl}${path}`, {
    ...options,
    headers: {
      Accept: 'application/json',
      ...(options.body ? { 'Content-Type': 'application/json' } : {}),
      ...(options.headers || {}),
    },
  })

  const body = await response.json().catch(() => null)
  if (!response.ok) {
    const message = body && typeof body === 'object' && 'error' in body ? String(body.error) : `C-NTA API returned ${response.status}`
    throw new Error(message)
  }
  return body as T
}

export async function loadDirectory() {
  const [clients, houses, profileRecords] = await Promise.all([
    apiRequest<ClientRecord[]>('/clients'),
    apiRequest<HouseRecord[]>('/houses'),
    apiRequest<Array<Partial<ClientProfile> & { clientId: string }>>('/client-profiles'),
  ])
  const profiles = profileRecords.reduce<Record<string, ClientProfile>>((result, profile) => {
    result[profile.clientId] = {
      ...emptyProfile,
      county: profile.county || '',
      gender: profile.gender || '',
      dateOfBirth: profile.dateOfBirth || '',
      needsGroup: profile.needsGroup || '',
      diagnosisCode: profile.diagnosisCode || '',
      weeklyRevenue: Number(profile.weeklyRevenue || 0),
      notes: profile.notes || '',
    }
    return result
  }, {})
  return { clients, houses, profiles }
}

export function updateClient(id: string, client: Omit<ClientRecord, '_id'>) {
  return apiRequest<Partial<ClientRecord>>(`/clients/${encodeURIComponent(id)}`, {
    method: 'PUT',
    body: JSON.stringify(client),
  })
}

export function updateClientProfile(clientId: string, profile: ClientProfile) {
  return apiRequest<ClientProfile & { clientId: string }>(`/client-profiles/${encodeURIComponent(clientId)}`, {
    method: 'PUT',
    body: JSON.stringify(profile),
  })
}

export function writeClientLog(entry: Record<string, unknown>) {
  return apiRequest('/logs', {
    method: 'POST',
    body: JSON.stringify(entry),
  })
}
