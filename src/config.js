// Public Supabase settings. These are safe to ship in browser code — access is
// controlled by the row-level security rules in supabase/schema.sql.
// The SECRET key must never go here; it lives only in Netlify's env settings.
export const SUPABASE_URL = 'https://fmmilmstqovaivpgpcah.supabase.co';
export const SUPABASE_PUBLISHABLE_KEY =
  'sb_publishable_EqojWYMZZ_aZuolIUHHF7Q_GzlQWD0q';

export const USERNAME_PATTERN = /^[a-z0-9_]{3,24}$/;
