import { createClient } from '@supabase/supabase-js';
import { SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY } from './config.js';

export const supabase = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY);

/** Current session, or null. Also picks up invite/reset tokens from the URL. */
export async function getSession() {
  const { data } = await supabase.auth.getSession();
  return data.session;
}

/** The signed-in worker's profile row, or null if they haven't picked a username yet. */
export async function getProfile(userId) {
  const { data, error } = await supabase
    .from('profiles')
    .select('username')
    .eq('id', userId)
    .maybeSingle();
  if (error) throw error;
  return data;
}

/**
 * Log in by username. The Netlify function swaps the username for the
 * worker's email server-side, signs in, and hands back the session tokens.
 */
export async function signInWithUsername(username, password) {
  const res = await fetch('/api/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username, password }),
  });
  const body = await res.json().catch(() => ({}));
  if (res.status === 429) {
    // Netlify's rate limiter answers before our function, with no JSON body.
    throw new Error(body.error || 'Too many login attempts. Wait a minute and try again.');
  }
  if (!res.ok) {
    throw new Error(body.error || 'Login failed. Please try again.');
  }
  const { error } = await supabase.auth.setSession({
    access_token: body.access_token,
    refresh_token: body.refresh_token,
  });
  if (error) throw error;
}

export async function signOut() {
  await supabase.auth.signOut();
  window.location.replace('/login.html');
}

/**
 * Page guard for worker-only pages. Redirects to login if signed out, or to
 * account setup if the worker hasn't chosen a username yet.
 * NOTE: this only controls what the page shows. Real data protection comes
 * from row-level security on every table workers can read.
 */
export async function requireWorker() {
  const session = await getSession();
  if (!session) {
    window.location.replace('/login.html');
    return null;
  }
  const profile = await getProfile(session.user.id);
  if (!profile) {
    window.location.replace('/setup.html');
    return null;
  }
  return { session, profile };
}

/** Show a message in a form's status element. */
export function showMessage(el, text, kind = 'error') {
  el.textContent = text;
  el.dataset.kind = kind;
  el.hidden = !text;
}
