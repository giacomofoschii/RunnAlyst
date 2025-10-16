# RunnAlyst 🏃‍♂️📊

**RunnAlyst** è uno strumento di analisi per runner che trasforma i file `.fit` grezzi, esportati da orologi GPS, in un report Excel dettagliato e facile da consultare.

Questo progetto nasce dall'esigenza di avere un controllo completo sui propri dati di allenamento, aggregandoli in un unico posto per un'analisi offline, personalizzata e senza dipendere da piattaforme esterne, con la possibilità di confrontare una o più corse tra loro.

---

### ✨ Funzionalità Principali

* **Parsing Automatico dei File `.fit`**: Legge tutti i file `.fit` presenti in una struttura di cartelle, anche suddivise per anno.
* **Riconoscimento del Tipo di Corsa**: Distingue automaticamente tra **Corsa**, **Corsa su Pista**, **Trail Running** e **Ultra Run** leggendo i metadati specifici del file.
* **Report Excel Multi-Foglio**: Genera un file `.xlsx` organizzato in tre sezioni:
    1.  **Storico Allenamenti**: Una dashboard riassuntiva con i dati chiave di ogni sessione (data, distanza, durata, passo medio, FC media).
    2.  **Dettaglio Lap**: Un'analisi approfondita di ogni singolo giro, fondamentale per visualizzare i dettagli di ripetute e allenamenti intervallati.
    3.  **Riepiloghi**: Calcoli automatici dei totali e delle medie su base **settimanale, mensile e annuale** per monitorare i carichi di lavoro e i progressi.
* **Log Incrementale e Non Distruttivo**: Lo script riconosce le attività già analizzate e aggiunge solo quelle nuove, preservando eventuali note manuali inserite dall'utente.
* **Campi Personalizzabili**: Include colonne vuote (`Tipo di Allenamento`, `Descrizione`) pensate per essere compilate manualmente, trasformando il report in un vero e proprio diario di allenamento.

---

### 🛠️ Stack Tecnologico

* **Python 3**
* `pandas`: Per la manipolazione dei dati e la creazione del file Excel.
* `fitparse`: Per la lettura e l'interpretazione del formato `.fit`.
* `openpyxl`: Come motore per la scrittura dei file `.xlsx`.
