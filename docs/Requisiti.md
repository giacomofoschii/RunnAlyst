# 2. Analisi dei Requisiti

Questo documento definisce i requisiti funzionali e non funzionali per la Versione 1.0 di RunnAlyst. Questi requisiti servono come guida per la progettazione dell'architettura e lo sviluppo delle funzionalità.

---

## 2.1 Modulo 1: Gestione Utenti e Autenticazione

Requisiti relativi all'identità, ai permessi e al profilo dell'utente.

* **RF-U1 (Login):** Il sistema deve fornire funzionalità di registrazione (nome, email, password) e login per gli utenti.
* **RF-U2 (Ruoli Utente):** Devono esistere tre livelli di autorizzazione:
    * **Atleta:** Accesso completo e modifica solo sui propri dati personali e sulle proprie attività.
    * **Coach:** Accesso in lettura ai dati (attività, profilo, calendario) di tutti gli atleti assegnati al suo team.
    * **Admin:** Accesso totale al sistema per la gestione degli utenti e delle configurazioni globali.
* **RF-U3 (Profilo Personale):** L'utente deve disporre di una schermata di profilo dove può modificare i propri dati anagrafici e caricare/aggiornare una foto profilo.
* **RF-U4 (Record Personali):** Il profilo utente deve includere una sezione dedicata dove l'atleta può inserire e tracciare manualmente i propri Record Personali (PB) su distanze standard (es. 400m, 1km, 5km, 10km, etc.) e altre "milestone".

---

## 2.2 Modulo 2: Gestione e Analisi delle Attività

Requisiti relativi al cuore dell'applicazione: l'importazione e l'analisi dei dati di corsa.

* **RF-A1 (Caricamento Manuale):** L'interfaccia deve permettere all'utente di caricare uno o più file `.fit` contemporaneamente.
* **RF-A2 (Classificazione Iniziale):** Durante il processo di caricamento, l'utente deve poter associare a ciascun file i seguenti metadati:
    * `Nome Attività` (Testo libero)
    * `Descrizione` (Testo libero)
    * `Tipo di Corsa` (Menu a tendina: Corsa, Corsa su Pista, Trail Running, etc.)
* **RF-A3 (Dashboard Attività):** Deve esistere una dashboard principale che elenca tutte le attività dell'utente in ordine cronologico inverso.
* **RF-A4 (Vista di Dettaglio):** Cliccando su un'attività, l'utente accede a una pagina di analisi dettagliata che mostra:
    * Un pannello di riepilogo con le metriche "effettive" (RF-A5).
    * Un pannello con le "Dinamiche di Corsa" avanzate (TCS, Oscillazione Verticale, etc.), se presenti nel file.
    * Una tabella interattiva dei `lap` (RF-A6).
* **RF-A5 (Metriche "Effettive"):** Il sistema deve calcolare e mostrare le metriche di riepilogo (Passo, FC Media, Cadenza Media) basandosi **esclusivamente** sui lap classificati come `Corsa` o `Recupero Attivo`, escludendo quindi i periodi di `Riscaldamento`, `Defaticamento` e `Recupero Passivo`.
* **RF-A6 (Classificazione Lap Interattiva):**
    * **RF-A6.1 (Classificazione Automatica):** Al caricamento, il sistema esegue una prima classificazione automatica dei lap basata su regole predefinite (es. passo > 6:00/km per il primo/ultimo lap = `Riscaldamento`/`Defaticamento`; passo > 8:30/km = `Recupero Passivo`; tutto il resto = `Corsa`).
    * **RF-A6.2 (Modifica Manuale):** L'utente deve avere la possibilità di **modificare e sovrascrivere** il tipo di ogni singolo lap tramite un'interfaccia (es. menu a tendina) scegliendo da una lista predefinita: `Corsa`, `Recupero Attivo`, `Recupero Passivo`, `Riscaldamento`, `Defaticamento`.
    * **RF-A6.3 (Salvataggio e Ricalcolo):** La modifica manuale del tipo di lap deve essere **salvata permanentemente** nel database per quella specifica attività e deve **ricalcolare istantaneamente** le metriche "effettive" (RF-A5).
* **RF-A7 (Autovalutazione):** L'utente deve poter associare a ogni attività un'autovalutazione soggettiva del proprio stato di forma (es. scala da 1 a 5, "Molto Debole" -> "Molto Forte").
* **RF-A8 (Esportazione Dati):** Dalla vista di dettaglio, l'utente deve poter:
    * Esportare i dati dell'attività (inclusi i lap classificati) in formato `CSV`.
    * Scaricare il file `.fit` originale che aveva caricato.

---

## 2.3 Modulo 3: Funzionalità di Team e Coaching

Requisiti che abilitano la collaborazione tra atleti e coach.

* **RF-S1 (Condivisione Atleta):** Un atleta deve poter impostare la visibilità di una propria attività su "condivisa", rendendola visibile agli altri membri del suo team.
* **RF-S2 (Confronto Attività):** Il sistema deve fornire una funzionalità per selezionare due o più attività (proprie o condivise) e visualizzarle in una vista di confronto dedicata (affiancando metriche e grafici).
* **RF-C1 (Calendario):** L'applicazione deve includere una vista calendario settimanale scorrevole.
* **RF-C2 (Disponibilità Atleta):** L'atleta deve poter usare il calendario per indicare i propri giorni di disponibilità all'allenamento per la settimana.
* **RF-C3 (Pianificazione Coach):** Il coach deve poter usare il calendario per assegnare allenamenti futuri a singoli atleti o a gruppi, visualizzando al contempo le loro disponibilità (RF-C2).
* **RF-P1 (Previsione Performance):** (Funzionalità avanzata) Quando il coach assegna un allenamento, il sistema deve mostrare un "range di performance atteso" per l'atleta, calcolato tramite un modello che considera lo storico delle performance, i cicli di carico/scarico e l'autovalutazione (RF-A7).

---

## 2.4 Modulo 4: Riepiloghi e Statistiche

Requisiti per l'analisi dei trend a lungo termine.

* **RF-R1 (Dashboard Riepiloghi):** Una sezione dedicata deve mostrare i dati aggregati (km totali, tempo totale, numero sessioni, etc.) su base settimanale, mensile e annuale, supportata da grafici di andamento.

---

## 2.5 Modulo 5: Funzionalità Future (Post v1.0)

Requisiti pianificati per le versioni successive.

* **RF-F1 (Sincronizzazione API):** Implementare un sistema di sincronizzazione API (es. tramite un pulsante "Sincronizza") per importare le attività da piattaforme esterne.
* **RF-F1.1 (Priorità Strava):** L'integrazione con l'API di Strava è considerata prioritaria, data la sua gratuità e la capacità di fornire i dati dei `lap` manuali originali, necessari per le nostre analisi (RF-A6).