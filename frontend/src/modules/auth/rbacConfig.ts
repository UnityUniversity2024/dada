/**
 * Export — RBAC (export module only)
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
  EXP_VIEW: ['admin', 'export_officer', 'sales_manager', 'export_finance', 'compliance'],
  EXP_EDIT: ['admin', 'export_officer', 'sales_manager'],
  CONTRACT_VIEW: ['admin', 'export_officer', 'sales_manager', 'export_finance', 'compliance'],
  CONTRACT_EDIT: ['admin', 'export_officer', 'sales_manager'],
  ROLE_CRUD: ['admin'],
  CAN_ADD_RECORD: ['admin', 'export_officer', 'sales_manager'],
  CAN_EDIT_RECORD: ['admin', 'export_officer', 'sales_manager'],
  CAN_DELETE_RECORD: ['admin', 'export_officer'],
};

export const MODULE_ACCESS: Record<ModuleKey, string[]> = {
  export: ['admin', 'export_officer', 'sales_manager', 'export_finance', 'compliance'],
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
    ],
  },
  {
    id: 'sales_manager',
    name: 'Sales Manager',
    modules: ['export'],
    permissions: [
      'EXP_VIEW', 'EXP_EDIT', 'CONTRACT_VIEW', 'CONTRACT_EDIT',
      'CAN_ADD_RECORD', 'CAN_EDIT_RECORD',
    ],
  },
  {
    id: 'export_finance',
    name: 'Export Finance',
    modules: ['export'],
    permissions: ['EXP_VIEW', 'CONTRACT_VIEW'],
  },
  {
    id: 'compliance',
    name: 'Compliance',
    modules: ['export'],
    permissions: ['EXP_VIEW', 'CONTRACT_VIEW'],
  },
];

export const DEFAULT_USERS: UserRecord[] = [
  { username: 'admin', password: 'password', role: 'admin', fullName: 'System Administrator', modules: ['export'] },
  { username: 'export_officer', password: 'password', role: 'export_officer', fullName: 'Export Officer', modules: ['export'] },
  { username: 'sales_manager', password: 'password', role: 'sales_manager', fullName: 'Sales Manager', modules: ['export'] },
  { username: 'export_finance', password: 'password', role: 'export_finance', fullName: 'Export Finance', modules: ['export'] },
  { username: 'compliance', password: 'password', role: 'compliance', fullName: 'Compliance Officer', modules: ['export'] },
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
  if (fromCatalog?.permissions?.length) return [...fromCatalog.permissions];

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
  const perms = getRolePermissions(role, roles);
  if (perms.includes('*') || perms.includes(permissionKey)) return true;
  const mapRoles = DEFAULT_PERMISSIONS_MAP[permissionKey] || [];
  return mapRoles.map(normalizeRoleName).includes(r);
}
