/**
 * Export — RBAC (export module + report / upload / export gates)
 */

export type ModuleKey = 'export';

export type UserRecord = {
  username: string;
  password: string;
  role: string;
  fullName?: string;
  modules?: ModuleKey[];
};

export type RoleDefinition = {
  id: string;
  name: string;
  modules: ModuleKey[];
  permissions: string[];
};

export const DEFAULT_PERMISSIONS_MAP: Record<string, string[]> = {
  EXP_VIEW: ['admin', 'export_officer', 'sales_manager', 'export_finance', 'compliance', 'warehouse', 'viewer'],
  EXP_EDIT: ['admin', 'export_officer', 'sales_manager'],
  CONTRACT_VIEW: ['admin', 'export_officer', 'sales_manager', 'export_finance', 'compliance'],
  CONTRACT_EDIT: ['admin', 'export_officer', 'sales_manager'],
  ROLE_CRUD: ['admin'],
  CAN_ADD_RECORD: ['admin', 'export_officer', 'sales_manager'],
  CAN_EDIT_RECORD: ['admin', 'export_officer', 'sales_manager'],
  CAN_DELETE_RECORD: ['admin', 'export_officer'],

  /* data encoder / Excel upload for coffee & sesame */
  UPLOAD_DATA: ['admin', 'export_officer', 'sales_manager'],

  /* report tabs */
  REPORT_COFFEE: ['admin', 'export_officer', 'sales_manager', 'export_finance', 'compliance', 'warehouse', 'viewer'],
  REPORT_SESAME: ['admin', 'export_officer', 'sales_manager', 'export_finance', 'compliance', 'warehouse', 'viewer'],
  REPORT_CONTRACT: ['admin', 'export_officer', 'sales_manager', 'export_finance', 'compliance'],
  REPORT_BOARD: ['admin', 'export_officer', 'sales_manager', 'export_finance'],

  /* file exports */
  EXPORT_EXCEL: ['admin', 'export_officer', 'sales_manager', 'export_finance', 'compliance', 'warehouse'],
  EXPORT_PDF: ['admin', 'export_officer', 'sales_manager', 'export_finance', 'compliance'],
  EXPORT_WORD: ['admin', 'export_officer', 'sales_manager'],
  EXPORT_CSV: ['admin', 'export_officer', 'sales_manager', 'export_finance'],
  EXPORT_BOARD: ['admin', 'export_officer', 'sales_manager', 'export_finance'],
};

export const MODULE_ACCESS: Record<ModuleKey, string[]> = {
  export: ['admin', 'export_officer', 'sales_manager', 'export_finance', 'compliance', 'warehouse', 'viewer'],
};

export const BASE_ROLES: RoleDefinition[] = [
  {
    id: 'admin',
    name: 'Admin',
    modules: ['export'],
    permissions: Object.keys(DEFAULT_PERMISSIONS_MAP),
  },
  {
    id: 'export_officer',
    name: 'Export Officer',
    modules: ['export'],
    permissions: [
      'EXP_VIEW', 'EXP_EDIT', 'CONTRACT_VIEW', 'CONTRACT_EDIT',
      'CAN_ADD_RECORD', 'CAN_EDIT_RECORD', 'CAN_DELETE_RECORD',
      'UPLOAD_DATA',
      'REPORT_COFFEE', 'REPORT_SESAME', 'REPORT_CONTRACT', 'REPORT_BOARD',
      'EXPORT_EXCEL', 'EXPORT_PDF', 'EXPORT_WORD', 'EXPORT_CSV', 'EXPORT_BOARD',
    ],
  },
  {
    id: 'sales_manager',
    name: 'Sales Manager',
    modules: ['export'],
    permissions: [
      'EXP_VIEW', 'EXP_EDIT', 'CONTRACT_VIEW', 'CONTRACT_EDIT',
      'CAN_ADD_RECORD', 'CAN_EDIT_RECORD',
      'UPLOAD_DATA',
      'REPORT_COFFEE', 'REPORT_SESAME', 'REPORT_CONTRACT', 'REPORT_BOARD',
      'EXPORT_EXCEL', 'EXPORT_PDF', 'EXPORT_WORD', 'EXPORT_CSV', 'EXPORT_BOARD',
    ],
  },
  {
    id: 'export_finance',
    name: 'Export Finance',
    modules: ['export'],
    permissions: [
      'EXP_VIEW', 'CONTRACT_VIEW',
      'REPORT_COFFEE', 'REPORT_SESAME', 'REPORT_CONTRACT', 'REPORT_BOARD',
      'EXPORT_EXCEL', 'EXPORT_PDF', 'EXPORT_CSV', 'EXPORT_BOARD',
    ],
  },
  {
    id: 'compliance',
    name: 'Compliance',
    modules: ['export'],
    permissions: [
      'EXP_VIEW', 'CONTRACT_VIEW',
      'REPORT_COFFEE', 'REPORT_SESAME', 'REPORT_CONTRACT',
      'EXPORT_EXCEL', 'EXPORT_PDF',
    ],
  },
  {
    id: 'warehouse',
    name: 'Warehouse',
    modules: ['export'],
    permissions: [
      'EXP_VIEW',
      'REPORT_COFFEE', 'REPORT_SESAME',
      'EXPORT_EXCEL',
    ],
  },
  {
    id: 'viewer',
    name: 'Viewer',
    modules: ['export'],
    permissions: [
      'EXP_VIEW',
      'REPORT_COFFEE', 'REPORT_SESAME',
    ],
  },
];

export const DEFAULT_USERS: UserRecord[] = [
  { username: 'admin', password: 'password', role: 'admin', fullName: 'System Administrator', modules: ['export'] },
  { username: 'export_officer', password: 'password', role: 'export_officer', fullName: 'Export Officer', modules: ['export'] },
  { username: 'sales_manager', password: 'password', role: 'sales_manager', fullName: 'Sales Manager', modules: ['export'] },
  { username: 'export_finance', password: 'password', role: 'export_finance', fullName: 'Export Finance', modules: ['export'] },
  { username: 'compliance', password: 'password', role: 'compliance', fullName: 'Compliance Officer', modules: ['export'] },
  { username: 'warehouse', password: 'password', role: 'warehouse', fullName: 'Warehouse Officer', modules: ['export'] },
  { username: 'viewer', password: 'password', role: 'viewer', fullName: 'Read-only Viewer', modules: ['export'] },
];

export function normalizeRoleName(role?: string | null): string {
  return String(role || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_');
}

export function canAccessModule(
  role: string,
  moduleId: ModuleKey,
  userModules?: ModuleKey[] | null,
): boolean {
  const r = normalizeRoleName(role);
  if (r === 'admin') return true;
  if (Array.isArray(userModules) && userModules.map(String).includes(moduleId)) return true;
  const allowed = MODULE_ACCESS[moduleId] || [];
  return allowed.map(normalizeRoleName).includes(r);
}

export function getRolePermissions(role: string, roles: RoleDefinition[] = BASE_ROLES): string[] {
  const r = normalizeRoleName(role);
  if (r === 'admin') return Object.keys(DEFAULT_PERMISSIONS_MAP);

  const fromCatalog = roles.find(
    (row) => normalizeRoleName(row.id) === r || normalizeRoleName(row.name) === r,
  );

  if (fromCatalog) {
    return Array.isArray(fromCatalog.permissions) ? [...fromCatalog.permissions] : [];
  }

  return Object.entries(DEFAULT_PERMISSIONS_MAP)
    .filter(([, list]) => list.map(normalizeRoleName).includes(r))
    .map(([key]) => key);
}

export function hasPermission(
  role: string,
  permissionKey: string,
  roles: RoleDefinition[] = BASE_ROLES,
): boolean {
  const r = normalizeRoleName(role);
  if (r === 'admin') return true;

  const fromCatalog = roles.find(
    (row) => normalizeRoleName(row.id) === r || normalizeRoleName(row.name) === r,
  );

  if (fromCatalog) {
    const perms = Array.isArray(fromCatalog.permissions) ? fromCatalog.permissions : [];
    return perms.includes('*') || perms.includes(permissionKey);
  }

  const mapRoles = DEFAULT_PERMISSIONS_MAP[permissionKey] || [];
  return mapRoles.map(normalizeRoleName).includes(r);
}
