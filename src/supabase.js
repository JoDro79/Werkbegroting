import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL ||
  "https://hjizveuklymmobqzsglb.supabase.co";

const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY ||
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImhqaXp2ZXVrbHltbW9icXpzZ2xiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzcxNTkyNjAsImV4cCI6MjA5MjczNTI2MH0.PRyNwCicqg1IaYe7-QlGIGmHljFpaYdel4tmLVGD5vU";

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
