# FantaSIN

Fantacongresso per il 56° Congresso di Neurologia, Firenze, 24–27 ottobre 2026.

| File | Cosa fa |
|---|---|
| `index.html` | App dei giocatori (GitHub Pages): squadra, classifica live, scommesse, relatori, regole |
| `dati.json` | Relatori, sessioni con sala, interventi e poster estratti dal programma |
| `Code.gs` | Backend Apps Script del Foglio Google: classifica in JSON, invio squadre e scommesse, app arbitri |
| `Arbitro.html` | App arbitri per telefono, servita da Apps Script (`.../exec?page=arbitro`) |

## Come funziona
- **Fogli**: Config, Relatori, Sessioni, Programma, Poster, Eventi, Arbitri, Log, Squadre, Scommesse, Puntate, Classifica.
- **Punti**: gli arbitri registrano dall'app; il Log tiene ora, sala, arbitro, relatore, evento e punti. Il tetto giornaliero (foglio Eventi) si applica per relatore, per tipo di evento, per giorno.
- **Classifica**: formule nel Foglio (capitano ×2, controllo budget, scommesse).
- **CORS**: GET semplice per leggere; POST con `Content-Type: text/plain` per scrivere (nessuna preflight). Il backend accetta anche `?callback=` (JSONP) come piano B.

Link giocatori: `https://fantasin.it/`
Link arbitri: `https://script.google.com/macros/s/<ID>/exec?page=arbitro`
