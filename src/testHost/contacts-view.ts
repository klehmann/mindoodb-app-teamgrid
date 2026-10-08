// A configured view for the test host, as Haven delivers one in the launch
// context: contacts by company, so view sheets (Data tab → MindooDB view)
// have something to show. The documents live in the teamgrid database next
// to the workbooks; the workbook list ignores them (other form).
import {
  createViewLanguage,
  type MindooDBAppResolvedViewDefinition,
  type MindooDBAppViewDefinition,
  type MindooDBAppViewNavigator,
  type MindooDBAppViewNavigatorOpenOptions,
} from 'mindoodb-app-sdk'
import { createEvaluatingViewNavigator, type MockSeedDocument } from 'mindoodb-app-sdk/testing'

export const CONTACTS: MockSeedDocument[] = [
  ['ACME', 'Alice Adams', 'Vertrieb', '2024-03-01', 12000],
  ['ACME', 'Bob Brown', 'Einkauf', '2025-01-15', 8000],
  ['Globex', 'Carla Cruz', 'Geschäftsführung', '2023-11-20', 25000],
  ['Globex', 'Dan Dorn', 'IT', '2026-02-02', 4000],
  ['Initech', 'Eva Ernst', 'Vertrieb', '2025-07-07', 9500],
].map(([company, name, department, since, revenue], index) => ({
  id: `contact_${index + 1}`,
  data: { form: 'contact', company, name, department, since, revenue },
}))

const v = createViewLanguage()
const COLUMNS = [
  { name: 'company', title: 'Firma', role: 'category', field: 'company' },
  { name: 'name', title: 'Name', role: 'display', field: 'name' },
  { name: 'department', title: 'Abteilung', role: 'display', field: 'department' },
  { name: 'since', title: 'Kunde seit', role: 'display', field: 'since' },
  { name: 'revenue', title: 'Umsatz', role: 'display', field: 'revenue' },
] as const

const VIEW_ID = 'contacts_by_company'

const DEFINITION: MindooDBAppViewDefinition = {
  id: VIEW_ID,
  title: 'Kontakte nach Firma',
  defaultExpand: 'expanded',
  columns: COLUMNS.map((column) => ({
    name: column.name,
    title: column.title,
    role: column.role,
    expression: v.field(column.field),
    sorting: 'ascending',
  })),
  filter: { mode: 'expression', expression: v.eq(v.field('form'), 'contact') },
} as MindooDBAppViewDefinition

export const CONTACTS_VIEW: MindooDBAppResolvedViewDefinition = {
  id: VIEW_ID,
  description: 'Kontakte nach Firma',
  categorizationStyle: 'category_then_document',
  previewMode: 'table',
  sources: [
    { origin: 'teamgrid', databaseId: 'teamgrid', title: 'teamgrid', targetMode: 'local', tenantId: 'test', databaseName: 'teamgrid' },
  ],
  filter: { mode: 'rules', match: 'all', rules: [{ id: 'form', field: 'form', operator: 'eq', value: 'contact' }] },
  columns: COLUMNS.map((column) => ({
    id: column.name,
    name: column.name,
    title: column.title,
    role: column.role,
    expression: { mode: 'field', field: column.field },
    sorting: 'ascending',
    totalMode: 'none',
    hidden: false,
  })),
}

/**
 * Opens the view over the seeded contacts. The mock session hands out the
 * same navigator for every later open of this view id, so disposing it must
 * not end it here (Haven opens a fresh one each time).
 */
export async function openContactsView(options?: MindooDBAppViewNavigatorOpenOptions): Promise<MindooDBAppViewNavigator> {
  const navigator = await createEvaluatingViewNavigator(
    { databaseIds: ['teamgrid'], definition: DEFINITION, categorizationStyle: 'category_then_document', ...(options ? { options } : {}) },
    CONTACTS.map((contact) => ({ origin: 'teamgrid', docId: contact.id, data: contact.data })),
  )
  if (!navigator) throw new Error('mindoodb is not installed: the test view cannot be evaluated')
  return new Proxy(navigator, {
    get: (target, property) => (property === 'dispose' ? async () => {} : Reflect.get(target, property)),
  })
}
