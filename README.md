# Meeting Club · Gestionale Scuola Nuoto

Web app per il registro presenze, corsi, anagrafica allievi, recuperi, valutazioni e pagamenti.
Questa è la versione **esportabile**: gira su un tuo hosting reale (Vercel), con un database vero (Supabase), indipendente da Claude.

---

## 1. Cosa ti serve (tutto gratuito)

- Un account [GitHub](https://github.com) — per salvare il codice
- Un account [Supabase](https://supabase.com) — il database
- Un account [Vercel](https://vercel.com) — l'hosting pubblico
- [Node.js](https://nodejs.org) installato sul tuo computer (versione 18 o superiore) — solo per testare in locale prima di pubblicare

---

## 2. Crea il database su Supabase

1. Vai su [supabase.com](https://supabase.com) → crea un account gratuito → **New Project**.
2. Dagli un nome (es. "scuola-nuoto"), scegli una password per il database (salvala da parte) e la regione più vicina (es. Frankfurt).
3. Attendi 1-2 minuti che il progetto sia pronto.
4. Vai nella sezione **SQL Editor** (menu a sinistra) → **New query**.
5. Apri il file `supabase-setup.sql` incluso in questo progetto, copia tutto il contenuto, incollalo nell'editor e premi **Run**.
   - Questo crea la tabella `app_data` dove verranno salvati tutti i dati dell'app (allievi, corsi, presenze, pagamenti, valutazioni).
6. Vai in **Project Settings → API**. Ti servono due valori:
   - **Project URL** (es. `https://xxxxx.supabase.co`)
   - **anon public key** (una chiave lunga)
   Tienili a portata di mano, ti serviranno tra poco.

---

## 3. Prova l'app sul tuo computer (opzionale ma consigliato)

1. Scompatta la cartella del progetto che ti ho fornito.
2. Apri il terminale dentro quella cartella.
3. Copia il file d'esempio delle variabili d'ambiente:
   ```
   cp .env.example .env
   ```
4. Apri `.env` con un editor di testo e incolla i due valori di Supabase:
   ```
   VITE_SUPABASE_URL=https://xxxxx.supabase.co
   VITE_SUPABASE_ANON_KEY=la-tua-chiave-anon-public
   ```
5. Installa le dipendenze:
   ```
   npm install
   ```
6. Avvia l'app in locale:
   ```
   npm run dev
   ```
7. Apri il link che appare nel terminale (di solito `http://localhost:5173`). Dovresti vedere l'app funzionante, con i dati salvati davvero su Supabase.

---

## 4. Pubblica il codice su GitHub

1. Crea un nuovo repository su [github.com](https://github.com/new) (puoi tenerlo privato).
2. Nel terminale, dentro la cartella del progetto:
   ```
   git init
   git add .
   git commit -m "Prima versione gestionale scuola nuoto"
   git branch -M main
   git remote add origin https://github.com/TUO-UTENTE/NOME-REPO.git
   git push -u origin main
   ```
   (Il file `.env` non verrà caricato: è escluso di proposito per non pubblicare le tue chiavi.)

---

## 5. Metti l'app online con Vercel

1. Vai su [vercel.com](https://vercel.com) → accedi con GitHub.
2. **Add New → Project** → seleziona il repository appena creato.
3. Vercel riconosce automaticamente che è un progetto Vite: lascia le impostazioni di default.
4. Prima di premere "Deploy", apri **Environment Variables** e aggiungi le stesse due variabili del file `.env`:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
5. Premi **Deploy**. Dopo circa un minuto avrai un link pubblico tipo `https://scuola-nuoto.vercel.app`, funzionante da qualsiasi dispositivo, senza passare da Claude.

Da questo momento, ogni volta che vuoi aggiornare l'app basta caricare le modifiche su GitHub (`git push`) e Vercel la ripubblica automaticamente in pochi secondi.

---

## 6. Un punto importante sulla sicurezza (da leggere)

Per farti partire subito, il database è configurato con accesso aperto (chiunque abbia il link dell'app può leggere e scrivere i dati). Va bene per iniziare o per un uso interno con link non condiviso pubblicamente, ma **non è un vero controllo accessi**.

Quando sei pronto per un uso strutturato con account separati per Admin e Segreteria, i prossimi passi consigliati sono:
1. Attivare **Supabase Auth** (login con email/password per staff).
2. Restringere le policy della tabella `app_data` (nel file SQL) in modo che solo utenti autenticati possano scrivere.
3. Distinguere i permessi Admin/Segreteria non solo nell'interfaccia (come ora) ma anche lato database.

Posso aiutarti a fare questo passaggio quando vuoi: basta chiedermelo.

---

## 7. Struttura del progetto

```
scuola-nuoto-app/
├── src/
│   ├── App.jsx           → tutta la logica e l'interfaccia dell'app
│   ├── main.jsx          → punto di ingresso React
│   ├── storage.js        → funzioni di lettura/scrittura verso Supabase
│   ├── supabaseClient.js → connessione a Supabase
│   └── index.css         → stili Tailwind
├── supabase-setup.sql    → script da eseguire una volta su Supabase
├── .env.example          → modello per le tue chiavi Supabase
├── package.json
└── README.md             → questo file
```

## 8. Come continuare a modificarla con Claude

Puoi sempre tornare in chat e chiedermi modifiche (nuove funzioni, campi, colori, logiche). Ti aggiorno il file `src/App.jsx`, tu lo sostituisci nella cartella del progetto e fai `git push` — Vercel pubblica da solo la nuova versione. I dati salvati su Supabase non vengono mai toccati da un aggiornamento del codice.
