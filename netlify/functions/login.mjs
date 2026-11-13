// POST /api/login  { username, password }
// Looks up the worker's email by username (server-side, so emails stay
// private), signs in with Supabase, and returns the session tokens.
import { createClient } from '@supabase/supabase-js';
import {
  SUPABASE_URL,
  SUPABASE_PUBLISHABLE_KEY,
  USERNAME_PATTERN,
} from '../../src/config.js';

const GENERIC_ERROR = 'Incorrect username or password.';

// Netlify blocks an IP with a 429 after 10 login attempts per minute, before
// this function runs, so spam never reaches Supabase.
// https://docs.netlify.com/manage/security/secure-access-to-sites/rate-limiting/
export const config = {
  path: '/api/login',
  rateLimit: {
    windowLimit: 10,
    windowSize: 60,
    aggregateBy: ['ip', 'domain'],
  },
};

// Every failed login takes at least this long. It slows down password guessing
// and hides whether the username existed (an unknown username would otherwise
// fail noticeably faster than a wrong password).
const MIN_FAILURE_MS = 1000;
const MAX_BODY_BYTES = 1024;
const MAX_PASSWORD_LENGTH = 72; // Supabase/bcrypt ignores anything longer

const json = (status, body) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });

export default async (req) => {
  const started = Date.now();
  const fail = async (status, body) => {
    const wait = MIN_FAILURE_MS - (Date.now() - started);
    if (wait > 0) await new Promise((r) => setTimeout(r, wait));
    return json(status, body);
  };

  if (req.method !== 'POST') return json(405, { error: 'Method not allowed' });
  if (Number(req.headers.get('content-length') || 0) > MAX_BODY_BYTES) {
    return json(413, { error: 'Invalid request.' });
  }

  const secretKey = process.env.SUPABASE_SECRET_KEY;
  if (!secretKey) {
    console.error('SUPABASE_SECRET_KEY is not set');
    return json(500, { error: 'Login is not configured yet.' });
  }

  let username, password;
  try {
    const raw = await req.text();
    if (raw.length > MAX_BODY_BYTES) return json(413, { error: 'Invalid request.' });
    ({ username, password } = JSON.parse(raw));
  } catch {
    return json(400, { error: 'Invalid request.' });
  }
  username = String(username ?? '').trim().toLowerCase();
  password = String(password ?? '');
  if (
    !USERNAME_PATTERN.test(username) ||
    !password ||
    password.length > MAX_PASSWORD_LENGTH
  ) {
    return fail(401, { error: GENERIC_ERROR });
  }

  const noPersist = { auth: { persistSession: false, autoRefreshToken: false } };
  const admin = createClient(SUPABASE_URL, secretKey, noPersist);

  const { data: profile, error: profileError } = await admin
    .from('profiles')
    .select('id')
    .eq('username', username)
    .maybeSingle();
  if (profileError) {
    console.error('profile lookup failed', profileError);
    return fail(500, { error: 'Login failed. Please try again.' });
  }
  // Same message whether the username or the password is wrong, so the form
  // can't be used to discover which usernames exist.
  if (!profile) return fail(401, { error: GENERIC_ERROR });

  const { data: userData, error: userError } =
    await admin.auth.admin.getUserById(profile.id);
  if (userError || !userData?.user?.email) {
    console.error('user lookup failed', userError);
    return fail(401, { error: GENERIC_ERROR });
  }

  const client = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, noPersist);
  const { data, error } = await client.auth.signInWithPassword({
    email: userData.user.email,
    password,
  });
  if (error || !data.session) {
    if (error?.status === 429) {
      return fail(429, { error: 'Too many attempts. Wait a minute and try again.' });
    }
    return fail(401, { error: GENERIC_ERROR });
  }

  return json(200, {
    access_token: data.session.access_token,
    refresh_token: data.session.refresh_token,
  });
};
