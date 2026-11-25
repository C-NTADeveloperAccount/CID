export type ClientRecord = {
  _id: string
  initials: string
  requiredTrainings?: string[]
  Address?: string
  CommunityCode?: string
  CommunityUnits?: string | number
  FacilityCode?: string
  FacilityUnits?: string | number
  TransportationCode?: string
  TransportationUnits?: string | number
  Units?: string | number
  DefaultSchedule?: string
  StartTime?: string
  EndTime?: string
  transportation?: boolean
  AdditionalSupport?: boolean
  ScheduledDays?: string[]
  communityCodeId?: string
  communityCode?: string
  facilityCodeId?: string
  facilityCode?: string
  transportationCodeId?: string
  transportationCode?: string
  active?: boolean
  isActive?: boolean
  inactive?: boolean
  [key: string]: unknown
}

export type HouseRecord = {
  _id: string
  houseName: string
  clients: string[]
  squad?: string
}

export type ClientProfile = {
  county: '' | 'Cambria' | 'Somerset'
  gender: '' | 'Female' | 'Male' | 'Nonbinary'
  dateOfBirth: string
  needsGroup: '' | 'NG 1' | 'NG 2' | 'NG 3' | 'NG 4' | 'NG 5'
  diagnosisCode: string
  weeklyRevenue: number
  notes: string
}

export type Individual = ClientRecord & ClientProfile & {
  id: string
  home: string
}

export const emptyProfile: ClientProfile = {
  county: '',
  gender: '',
  dateOfBirth: '',
  needsGroup: '',
  diagnosisCode: '',
  weeklyRevenue: 0,
  notes: '',
}

export function isActiveClient(client: ClientRecord) {
  return client.active !== false && client.isActive !== false && client.inactive !== true
}

export function mergeIndividuals(clients: ClientRecord[], houses: HouseRecord[], profiles: Record<string, ClientProfile>): Individual[] {
  return clients.filter(isActiveClient).map((client) => ({
    ...client,
    ...emptyProfile,
    ...(profiles[client._id] || {}),
    id: client._id,
    home: houses.find((house) => house.clients?.includes(client._id))?.houseName || 'Unassigned',
    initials: client.initials || 'Unknown',
  }))
}
