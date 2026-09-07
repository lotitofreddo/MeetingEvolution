-- Esegui questo script nell'SQL Editor di Supabase (una sola volta)

create table if not exists app_data (
  key text primary key,
  value jsonb,
  updated_at timestamptz default now()
);

-- Abilita la sicurezza a livello di riga
alter table app_data enable row level security;

-- Consente lettura e scrittura tramite la chiave "anon"
-- (va bene per iniziare da soli; vedi il README per aggiungere un vero login più avanti)
create policy "Consenti lettura pubblica" on app_data
  for select using (true);

create policy "Consenti scrittura pubblica" on app_data
  for insert with check (true);

create policy "Consenti aggiornamento pubblico" on app_data
  for update using (true);
