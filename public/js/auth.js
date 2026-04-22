const TOKEN_KEY = 'friday_crm_token';
const USER_KEY  = 'friday_crm_user';

function getToken() { return sessionStorage.getItem(TOKEN_KEY); }
function getUser()  { return JSON.parse(sessionStorage.getItem(USER_KEY) || 'null'); }
function setAuth(token, user) {
  sessionStorage.setItem(TOKEN_KEY, token);
  sessionStorage.setItem(USER_KEY, JSON.stringify(user));
}
function clearAuth() {
  sessionStorage.removeItem(TOKEN_KEY);
  sessionStorage.removeItem(USER_KEY);
}

function requireLogin(role) {
  const token = getToken();
  const user  = getUser();
  if (!token || !user) { location.href = '/index.html'; return false; }
  if (role && user.role !== role) { location.href = user.role === 'manager' ? '/manager.html' : '/map.html'; return false; }
  return true;
}

async function apiFetch(path, options = {}) {
  const token = getToken();
  const res = await fetch(path, {
    ...options,
    headers: {
      ...(options.headers || {}),
      ...(token ? { 'Authorization': `Bearer ${token}` } : {}),
      ...(options.body instanceof FormData ? {} : { 'Content-Type': 'application/json' })
    }
  });
  if (res.status === 401) { clearAuth(); location.href = '/index.html'; return null; }
  return res;
}

function logout() { clearAuth(); location.href = '/index.html'; }
