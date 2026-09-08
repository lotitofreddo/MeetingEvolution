import { supabase } from "./supabaseClient";

// Legge un valore salvato (per chiave). Ritorna null se non esiste ancora.
export async function storageGet(key) {
  const { data, error } = await supabase
    .from("app_data")
    .select("value")
    .eq("key", key)
    .maybeSingle();
  if (error) throw error;
  return data ? data.value : null;
}

// Salva (o aggiorna) un valore per una data chiave.
export async function storageSet(key, value) {
  const { error } = await supabase
    .from("app_data")
    .upsert({ key, value, updated_at: new Date().toISOString() });
  if (error) throw error;
}
