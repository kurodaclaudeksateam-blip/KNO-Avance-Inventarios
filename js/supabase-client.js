// Proyecto Supabase dedicado a KNO-Avance-Inventarios (no se comparte con otros proyectos).
// La clave es la "publishable key" pública (segura de exponer en el cliente); el acceso real
// se controla con las políticas RLS definidas en las tablas `cargas` y `areas_almacen`.
const SUPABASE_URL = 'https://jtgungaomigzfzvxwrtc.supabase.co';
const SUPABASE_PUBLISHABLE_KEY = 'sb_publishable__gkn7xXxu0eEMm3xQuB4qQ_rz3Hn2O5';

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY);
