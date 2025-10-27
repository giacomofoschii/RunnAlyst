# 1. Introduzione

## 1.1 Autore del Progetto

RunnAlyst è sviluppato da Giacomo Foschi.

## 1.2 Credits e Ringraziamenti

Questo progetto è stato sviluppato con il supporto e l'assistenza di **Google Gemini**, un partner AI che ha contribuito alla definizione dell'architettura, alla stesura del codice e al debugging delle funzionalità.

## 1.3 Scopo del Progetto

**RunnAlyst** è una piattaforma web di analisi dati progettata specificamente per corridori e coach che desiderano un controllo approfondito e personalizzato sulle proprie performance. L'obiettivo primario è trasformare i dati grezzi provenienti dai file `.fit` dei dispositivi GPS in insight strategici, aggregati e visualizzati in modo chiaro e intuitivo.

A differenza delle piattaforme commerciali esistenti, RunnAlyst mira a fornire:
* **Analisi Dettagliata:** Focus su metriche specifiche per la corsa, come l'analisi dei lap distinguendo tra fasi attive e di recupero.
* **Personalizzazione:** Possibilità per l'utente di classificare i propri allenamenti e i singoli lap, garantendo l'accuratezza dei calcoli.
* **Funzionalità di Coaching e Team:** Strumenti dedicati ai coach per monitorare gli atleti, pianificare allenamenti condivisi, analizzare i trend di performance e facilitare la condivisione e il confronto dei dati all'interno della squadra.
* **Controllo dei Dati:** Un ambiente in cui l'utente ha la piena proprietà e il controllo dei propri dati di allenamento.

## 1.4 Contesto e Motivazioni

Il panorama attuale delle piattaforme di analisi sportiva (es. Garmin Connect, Strava, TrainingPeaks) offre ottimi strumenti, ma spesso presenta limitazioni per un'analisi tecnica approfondita o per le specifiche esigenze di un team di atletica:
* **Analisi dei Lap:** La gestione e la classificazione dei lap manuali (fondamentali per le ripetute) è spesso limitata o assente.
* **Personalizzazione:** L'utente ha scarso controllo su come i dati vengono processati e classificati.
* **Funzionalità di Squadra:** Gli strumenti per la pianificazione centralizzata e l'analisi comparativa specifica per un team sono spesso carenti o legati a costi elevati.
* **Costi:** Le funzionalità avanzate sono frequentemente legate a piani di abbonamento.

RunnAlyst nasce per colmare queste lacune, offrendo uno strumento focalizzato, potente e adattabile alle esigenze specifiche di atleti e allenatori, con un'architettura che permette un'evoluzione continua.

## 1.5 Architettura e Infrastruttura Previste

Per realizzare gli obiettivi del progetto, è stata definita la seguente architettura tecnologica:
* **Backend:** **Python** con un web framework moderno (es. Flask o FastAPI), scelto per il suo potente ecosistema di analisi dati (Pandas, NumPy).
* **Frontend:** Un framework **JavaScript** moderno (es. React/Next.js o Vue/Nuxt.js) abbinato a **Tailwind CSS** per creare un'interfaccia utente reattiva, accattivante e interattiva.
* **Database:** Un database **relazionale (SQL)**, utilizzando **SQLite** per lo sviluppo locale e prevedendo la migrazione a **PostgreSQL** per l'ambiente di produzione online.
* **Archiviazione File:** Un servizio di **Object Storage cloud** (con **Cloudflare R2** come opzione iniziale preferita per il suo piano gratuito) per salvare i file `.fit` originali e le immagini del profilo.
* **Hosting:** Piattaforme cloud con piani gratuiti o a basso costo (es. Render, Vercel, PythonAnywhere) per la messa online iniziale.

## 1.6 Glossario Essenziale

* **FIT (Flexible and Interoperable Data Transfer):** Formato file binario standard, utilizzato dalla maggior parte dei dispositivi GPS (Garmin, Wahoo, etc.) per registrare i dati di un'attività sportiva.
* **Lap:** Un singolo intervallo (giro) registrato durante un'attività, sia automaticamente (es. ogni km) che manually dall'utente.
* **Metriche "Effettive":** Valori calcolati escludendo i periodi identificati come recupero passivo, riscaldamento o defaticamento (es. Passo Medio Effettivo).
* **API (Application Programming Interface):** Interfaccia software che permette a diverse applicazioni di comunicare tra loro (es. per sincronizzare dati con Strava).
* **CI/CD (Continuous Integration / Continuous Delivery):** Pratiche di automazione per garantire la qualità del codice e facilitare il rilascio di nuove versioni.
* **Framework Web:** Insieme di librerie e strumenti che facilitano lo sviluppo di applicazioni web (es. Flask, Django per il backend; React, Vue per il frontend).
* **Database Relazionale (SQL):** Sistema per memorizzare dati in tabelle strutturate con relazioni definite tra loro (es. PostgreSQL, SQLite).
* **Object Storage:** Servizio cloud per archiviare file (es. i file `.fit` originali, le foto profilo).