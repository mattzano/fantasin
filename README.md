# FantaSIN

Fantacongresso per il 56° Congresso di Neurologia, Firenze, 24–27 ottobre 2026.

- `index.html`: la web app (squadre, classifica, pannello arbitro). Serve un progetto Firebase (Firestore + accesso anonimo): incolla la configurazione in `FIREBASE_CONFIG` dentro `index.html`.
- `firestore.rules`: regole di sicurezza da incollare in Firestore → Regole.

Link arbitri: `https://mattzano.github.io/fantasin/?arbitro=CODICE`, dove `CODICE` è l'ID di un documento nella raccolta `refkeys`.
