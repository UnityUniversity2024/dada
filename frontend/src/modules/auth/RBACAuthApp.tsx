import React, { useEffect, useMemo, useState } from 'react';
import ExportApp from '../export/ExportApp';
import {
  BASE_ROLES,
  DEFAULT_PERMISSIONS_MAP,
  DEFAULT_USERS,
  canAccessModule,
  getRolePermissions,
  hasPermission,
  normalizeRoleName,
  type ModuleKey,
  type RoleDefinition,
  type UserRecord,
} from './rbacConfig';
import ExportApp from '../export/ExportApp';
import './rbac.css';
const USERS_KEY = 'unity_export_users';
const ROLES_KEY = 'unity_export_roles';
const SESSION_KEY = 'unity_export_session';

const tabList: Array<{ id: 'export' | 'settings'; label: string }> = [
  { id: 'export', label: 'Export' },
  { id: 'settings', label: 'RBAC Settings' },
];

const readJson = <T,>(key: string, fallback: T): T => {
  try {
    const value = window.localStorage.getItem(key);
    return value ? (JSON.parse(value) as T) : fallback;
  } catch {
    return fallback;
  }
};

const writeJson = <T,>(key: string, value: T) => {
  window.localStorage.setItem(key, JSON.stringify(value));
};

const defaultRoleChoices = Object.keys(DEFAULT_PERMISSIONS_MAP);

export default function RBACAuthApp(): JSX.Element {
  const [users, setUsers] = useState<UserRecord[]>(() => readJson<UserRecord[]>(USERS_KEY, DEFAULT_USERS));
  const [roles, setRoles] = useState<RoleDefinition[]>(() => readJson<RoleDefinition[]>(ROLES_KEY, BASE_ROLES));
  // Always start on login
  const [session, setSession] = useState<UserRecord | null>(null);
  const [activeTab, setActiveTab] = useState<'export' | 'settings'>('export');
  const [loginMode, setLoginMode] = useState<'login' | 'register'>('login');
  const [loginForm, setLoginForm] = useState({ username: 'admin', password: 'password' });
  const [registerForm, setRegisterForm] = useState({
    username: '',
    password: 'password',
    role: 'export_officer',
    fullName: '',
    modules: ['export'] as ModuleKey[],
  });
  const [notice, setNotice] = useState('');

  useEffect(() => {
    writeJson(USERS_KEY, users);
  }, [users]);

  useEffect(() => {
    writeJson(ROLES_KEY, roles);
  }, [roles]);

  useEffect(() => {
    if (session) {
      writeJson(SESSION_KEY, session);
    } else {
      window.localStorage.removeItem(SESSION_KEY);
    }
  }, [session]);

  const permittedModules = useMemo(() => {
    if (!session) return [] as ModuleKey[];
    const roleModules = roles
      .filter(
        (roleRecord) =>
          normalizeRoleName(roleRecord.id) === normalizeRoleName(session.role) ||
          normalizeRoleName(roleRecord.name) === normalizeRoleName(session.role),
      )
      .flatMap((roleRecord) => roleRecord.modules);

    const merged = [...new Set([...(session.modules ?? []), ...roleModules])];
    return merged.filter((module) => canAccessModule(session.role, module, session.modules));
  }, [roles, session]);

  const currentPermissions = useMemo(() => {
    if (!session) return [] as string[];
    return getRolePermissions(session.role, roles);
  }, [roles, session]);

  const loginUser = (username: string, password: string) => {
    const user = users.find(
      (entry) => entry.username.toLowerCase() === username.toLowerCase() && entry.password === password,
    );

    if (!user) {
      setNotice('Invalid username or password. Use the demo credentials from the RBAC setup.');
      return;
    }

    setSession(user);
    setActiveTab('export');
    setNotice(`Logged in as ${user.username} (${user.role}).`);
  };

  const handleLoginSubmit = (event: React.FormEvent) => {
    event.preventDefault();
    loginUser(loginForm.username, loginForm.password);
  };

  const handleRegisterSubmit = (event: React.FormEvent) => {
    event.preventDefault();
    const username = registerForm.username.trim();
    const password = registerForm.password.trim();
    const fullName = registerForm.fullName.trim() || username;

    if (!username || !password) {
      setNotice('Username and password are required for registration.');
      return;
    }

    if (users.some((user) => user.username.toLowerCase() === username.toLowerCase())) {
      setNotice('This username already exists. Please choose another one.');
      return;
    }

    const newUser: UserRecord = {
      username,
      password,
      role: registerForm.role,
      modules: ['export'],
      fullName,
    };

    setUsers((previous) => [...previous, newUser]);
    setRegisterForm({ username: '', password: 'password', role: 'export_officer', fullName: '', modules: ['export'] });
    setLoginMode('login');
    setLoginForm({ username, password });
    setNotice(`User ${username} registered successfully. You can log in now.`);
  };

  const removeUser = (username: string) => {
    if (!session || session.username === username) {
      setNotice('You cannot remove the active session user.');
      return;
    }

    setUsers((previous) => previous.filter((user) => user.username !== username));
    setNotice(`User ${username} removed.`);
  };

  const removeRole = (roleId: string) => {
    if (roleId === 'admin') {
      setNotice('The admin role cannot be removed.');
      return;
    }

    setRoles((previous) => previous.filter((role) => role.id !== roleId));
    setUsers((previous) =>
      previous.map((user) =>
        user.role === roleId ? { ...user, role: 'export_officer', modules: ['export'] } : user,
      ),
    );
    setNotice(`Role ${roleId} removed and affected users were reset to export_officer.`);
  };

  const addRole = (event: React.FormEvent) => {
    event.preventDefault();
    const form = new FormData(event.currentTarget as HTMLFormElement);
    const id = String(form.get('roleId') || '').trim();
    const name = String(form.get('roleName') || '').trim();
    const permissions = (form.getAll('permissions') as string[]).filter(Boolean);

    if (!id || !name) {
      setNotice('Role ID and role name are required.');
      return;
    }

    if (
      roles.some(
        (role) =>
          role.id.toLowerCase() === id.toLowerCase() || role.name.toLowerCase() === name.toLowerCase(),
      )
    ) {
      setNotice('That role ID or name already exists.');
      return;
    }

    setRoles((previous) => [...previous, { id, name, modules: ['export'], permissions }]);
    setNotice(`Role ${name} created.`);
    (event.currentTarget as HTMLFormElement).reset();
  };

  const isAdmin = session?.role === 'admin' || normalizeRoleName(session?.role || '') === 'admin';

  if (!session) {
    return (
      <div className="rbac-page-shell">
        <div className="rbac-card">
          <div className="rbac-card-header">
            <div>
              <span className="rbac-kicker">Export</span>
              <h1>Login</h1>
            </div>
            <div className="rbac-tabs">
              <button
                type="button"
                className={loginMode === 'login' ? 'rbac-tab active' : 'rbac-tab'}
                onClick={() => setLoginMode('login')}
              >
                Login
              </button>
              <button
                type="button"
                className={loginMode === 'register' ? 'rbac-tab active' : 'rbac-tab'}
                onClick={() => setLoginMode('register')}
              >
                Register
              </button>
            </div>
          </div>

          {loginMode === 'login' ? (
            <form onSubmit={handleLoginSubmit} className="rbac-form">
              <label>
                Username
                <input
                  value={loginForm.username}
                  onChange={(event) => setLoginForm((prev) => ({ ...prev, username: event.target.value }))}
                />
              </label>
              <label>
                Password
                <input
                  type="password"
                  value={loginForm.password}
                  onChange={(event) => setLoginForm((prev) => ({ ...prev, password: event.target.value }))}
                />
              </label>
              <button type="submit" className="rbac-primary-button">
                Sign in
              </button>
            </form>
          ) : (
            <form onSubmit={handleRegisterSubmit} className="rbac-form">
              <label>
                Full name
                <input
                  value={registerForm.fullName}
                  onChange={(event) => setRegisterForm((prev) => ({ ...prev, fullName: event.target.value }))}
                />
              </label>
              <label>
                Username
                <input
                  value={registerForm.username}
                  onChange={(event) => setRegisterForm((prev) => ({ ...prev, username: event.target.value }))}
                />
              </label>
              <label>
                Password
                <input
                  type="password"
                  value={registerForm.password}
                  onChange={(event) => setRegisterForm((prev) => ({ ...prev, password: event.target.value }))}
                />
              </label>
              <label>
                Role
                <select
                  value={registerForm.role}
                  onChange={(event) =>
                    setRegisterForm((prev) => ({
                      ...prev,
                      role: event.target.value,
                      modules: ['export'],
                    }))
                  }
                >
                  {roles.map((role) => (
                    <option key={role.id} value={role.id}>
                      {role.name}
                    </option>
                  ))}
                </select>
              </label>
              <button type="submit" className="rbac-primary-button">
                Create account
              </button>
            </form>
          )}

          <div className="rbac-demo-box">
            <strong>Demo users</strong>
            <div className="rbac-demo-grid">
              {DEFAULT_USERS.map((user) => (
                <div key={user.username} className="rbac-demo-item">
                  <span>{user.username}</span>
                  <small>{user.role}</small>
                </div>
              ))}
            </div>
            <p>
              Password for all demo accounts: <strong>password</strong>
            </p>
          </div>

          {notice && <div className="rbac-notice">{notice}</div>}
        </div>
      </div>
    );
  }

  return (
    <div className="rbac-shell">
      <header className="rbac-topbar">
        <div>
          <span className="rbac-kicker">Export</span>
          <h2>Export operations</h2>
        </div>
        <div className="rbac-user-panel">
          <div>
            <small>Signed in</small>
            <strong>{session.fullName || session.username}</strong>
          </div>
          <span className="rbac-role-badge">{session.role}</span>
          <button type="button" className="rbac-secondary-button" onClick={() => setSession(null)}>
            Logout
          </button>
        </div>
      </header>

      <nav className="rbac-module-nav">
        {tabList.map((tab) => {
          const disabled =
            tab.id === 'export'
              ? !canAccessModule(session.role, 'export', session.modules)
              : !isAdmin && !hasPermission(session.role, 'ROLE_CRUD', roles);
          return (
            <button
              key={tab.id}
              type="button"
              className={activeTab === tab.id ? 'rbac-module-button active' : 'rbac-module-button'}
              onClick={() => {
                if (disabled) return;
                setActiveTab(tab.id);
              }}
              disabled={disabled}
            >
              {tab.label}
            </button>
          );
        })}
      </nav>

      {notice && <div className="rbac-notice">{notice}</div>}

      {activeTab === 'settings' ? (
        <div className="rbac-settings-grid">
          <div className="rbac-panel">
            <h3>Role manager</h3>
            <form onSubmit={addRole} className="rbac-form compact-form">
              <label>
                Role ID
                <input name="roleId" placeholder="export_viewer" />
              </label>
              <label>
                Role name
                <input name="roleName" placeholder="Export Viewer" />
              </label>
              <label>
                Permissions
                <div className="rbac-checkbox-row wrap-row">
                  {defaultRoleChoices.map((permissionKey) => (
                    <label key={permissionKey} className="rbac-checkbox-item">
                      <input type="checkbox" name="permissions" value={permissionKey} />
                      {permissionKey}
                    </label>
                  ))}
                </div>
              </label>
              <button type="submit" className="rbac-primary-button">
                Add role
              </button>
            </form>

            <div className="rbac-table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Role</th>
                    <th>Modules</th>
                    <th>Permissions</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {roles.map((role) => (
                    <tr key={role.id}>
                      <td>{role.name}</td>
                      <td>{role.modules.join(', ') || '—'}</td>
                      <td>{role.permissions.slice(0, 4).join(', ') || 'No permissions'}</td>
                      <td>
                        {role.id !== 'admin' && (
                          <button type="button" className="rbac-danger-button" onClick={() => removeRole(role.id)}>
                            Remove
                          </button>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="rbac-panel">
            <h3>User manager</h3>
            <form className="rbac-form compact-form">
              <label>
                Username
                <input
                  value={registerForm.username}
                  onChange={(event) => setRegisterForm((prev) => ({ ...prev, username: event.target.value }))}
                />
              </label>
              <label>
                Password
                <input
                  type="password"
                  value={registerForm.password}
                  onChange={(event) => setRegisterForm((prev) => ({ ...prev, password: event.target.value }))}
                />
              </label>
              <label>
                Role
                <select
                  value={registerForm.role}
                  onChange={(event) =>
                    setRegisterForm((prev) => ({
                      ...prev,
                      role: event.target.value,
                      modules: ['export'],
                    }))
                  }
                >
                  {roles.map((role) => (
                    <option key={role.id} value={role.id}>
                      {role.name}
                    </option>
                  ))}
                </select>
              </label>
              <button
                type="button"
                className="rbac-primary-button"
                onClick={() => {
                  if (!registerForm.username.trim() || !registerForm.password.trim()) {
                    setNotice('A username and password are required.');
                    return;
                  }
                  if (users.some((user) => user.username.toLowerCase() === registerForm.username.trim().toLowerCase())) {
                    setNotice('That user already exists.');
                    return;
                  }
                  const newUser: UserRecord = {
                    username: registerForm.username.trim(),
                    password: registerForm.password,
                    role: registerForm.role,
                    modules: ['export'],
                    fullName: registerForm.fullName || registerForm.username,
                  };
                  setUsers((previous) => [...previous, newUser]);
                  setRegisterForm({
                    username: '',
                    password: 'password',
                    role: 'export_officer',
                    fullName: '',
                    modules: ['export'],
                  });
                  setNotice(`User ${newUser.username} created successfully.`);
                }}
              >
                Add user
              </button>
            </form>

            <div className="rbac-table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>User</th>
                    <th>Role</th>
                    <th>Modules</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {users.map((user) => (
                    <tr key={user.username}>
                      <td>{user.fullName || user.username}</td>
                      <td>{user.role}</td>
                      <td>{(user.modules || []).join(', ') || '—'}</td>
                      <td>
                        <button type="button" className="rbac-danger-button" onClick={() => removeUser(user.username)}>
                          Remove
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      ) : canAccessModule(session.role, 'export', session.modules) ? (
        <div className="rbac-dashboard-shell">
          <ExportApp />
          <div className="rbac-analytics-grid" style={{ marginTop: 18 }}>
            <div className="rbac-panel">
              <h3>My permissions</h3>
              <div className="rbac-chip-row">
                {currentPermissions.length > 0 ? (
                  currentPermissions.map((permission) => (
                    <span key={permission} className="rbac-chip">
                      {permission}
                    </span>
                  ))
                ) : (
                  <span className="rbac-empty">No permissions assigned.</span>
                )}
              </div>
            </div>
            <div className="rbac-panel">
              <h3>Module access</h3>
              <div className="rbac-chip-row">
                {permittedModules.length > 0 ? (
                  permittedModules.map((module) => (
                    <span key={module} className="rbac-chip module-chip">
                      {module}
                    </span>
                  ))
                ) : (
                  <span className="rbac-empty">No module access assigned.</span>
                )}
              </div>
            </div>
          </div>
        </div>
      ) : (
        <div className="rbac-denied">Access denied for Export.</div>
      )}
    </div>
  );
}

