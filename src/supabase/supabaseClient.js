// ============================================================================
// CLIENTE DE SUPABASE — S&S IA
// ============================================================================
// Conexión a la base de datos real y compartida — reemplaza localStorage
// (que solo vivía en un navegador) por datos que todos los dispositivos
// conectados ven y actualizan entre sí, en tiempo real.
//
// La URL y la clave "anon" (pública) de tu proyecto de Supabase NO son
// secretas en el sentido tradicional — están pensadas para vivir en el
// código del frontend, igual que el clientId de Azure. La seguridad real
// vive en las políticas de RLS (Row Level Security) definidas en
// supabase/schema.sql — ajústalas si más adelante necesitas restringir
// quién puede leer/escribir qué.
// ============================================================================

import { createClient } from "@supabase/supabase-js";

// AJUSTA ESTOS DOS VALORES con los de tu propio proyecto de Supabase:
// Panel de Supabase → tu proyecto → Project Settings → API.
const SUPABASE_URL = "https://mijoktgrcirvfirkwlzc.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1pam9rdGdyY2lydmZpcmt3bHpjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODgwMDk3NzMsImV4cCI6MjEwMzU4NTc3M30.OnPj8LlpedHs0A7foB45-Ok5tE3_zvNkstWlzRHq7ko";

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
