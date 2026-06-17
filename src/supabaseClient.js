import { createClient } from "@supabase/supabase-js";

// Project URL is not secret. The anon key below is the "public" key from
// Supabase Settings -> API -> Project API keys -> anon/public. It is safe
// to ship in frontend code: Row Level Security policies on the database
// (see schema.sql) are what actually restrict each user to their own data,
// not secrecy of this key.
//
// Both values are read from environment variables (see .env.example) so
// project-specific config never has to be committed to source control.
// Vite only exposes variables prefixed with VITE_ to the browser bundle.
const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!SUPABASE_URL || !SUPABASE_ANON_KEY) {
  throw new Error(
    "Missing Supabase config. Copy .env.example to .env and fill in " +
      "VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY from your Supabase " +
      "project's Settings -> API page, then restart the dev server."
  );
}

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

