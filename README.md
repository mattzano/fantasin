# FantaSIN

Fantacongresso per il 56° Congresso di Neurologia, Firenze, 24–27 ottobre 2026.

- `index.html`: la web app (squadre, classifica, pannello arbitro), pubblicata con GitHub Pages.
- `Code.gs`: il backend, da incollare in Apps Script dentro un Foglio Google (Estensioni → Apps Script) e distribuire come App web con accesso "Chiunque". L'URL `/exec` va in `API_URL` dentro `index.html`.

Link giocatori: `https://mattzano.github.io/fantasin/`
Link arbitri: `https://mattzano.github.io/fantasin/?arbitro=CODICE` (lo stesso `REF_CODE` scritto in `Code.gs`).
