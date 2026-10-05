const DEFAULT_BASE = 'http://localhost:3002/api';

const baseUrl = import.meta?.env?.VITE_EXPORT_API_BASE_URL || DEFAULT_BASE;

export const exportApi = {
  async bootstrap() {
    const response = await fetch(`${baseUrl}/bootstrap`);
    if (!response.ok) throw new Error('Bootstrap failed');
    return response.json();
  },

  async login(username, password) {
    const response = await fetch(`${baseUrl}/auth/login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password }),
    });

    if (!response.ok) throw new Error('Login failed');
    return response.json();
  },

  async listCoffee() {
    const response = await fetch(`${baseUrl}/coffee`, {
      headers: { Authorization: `Bearer demo-token` },
    });
    if (!response.ok) throw new Error('Coffee fetch failed');
    return response.json();
  },

  async listSesame() {
    const response = await fetch(`${baseUrl}/sesame`, {
      headers: { Authorization: `Bearer demo-token` },
    });
    if (!response.ok) throw new Error('Sesame fetch failed');
    return response.json();
  },
};
