from fitparse import FitFile
from datetime import timedelta
import pandas as pd

# Le funzioni di formattazione sono di supporto al parser
def formatta_passo(passo_decimale):
    # ... (copia la funzione da analizza_corse.py) ...
    if pd.isna(passo_decimale) or passo_decimale == float('inf'):
        return "00:00"
    minuti = int(passo_decimale)
    secondi = int((passo_decimale * 60) % 60)
    return f"{minuti:02d}:{secondi:02d}"

def formatta_durata(secondi_totali):
    # ... (copia la funzione da analizza_corse.py) ...
    if pd.isna(secondi_totali):
        return "00:00:00"
    return str(timedelta(seconds=int(secondi_totali)))

def analizza_file_fit(percorso_file):
    # ... (copia tutta la logica della funzione analizza_file_fit da analizza_corse.py) ...
    # Assicurati che alla fine restituisca i due valori:
    # return riga_storico, righe_lap
    try:
        fitfile = FitFile(percorso_file)
        
        dati_attivita = {}
        lista_lap = []

        for record in fitfile.get_messages(['session', 'lap', 'event', 'file_id']):
            if record.name == 'session':
                for field in record.fields:
                    dati_attivita[field.name] = field.value
            if record.name == 'lap':
                lap_data = {}
                for field in record.fields:
                    lap_data[field.name] = field.value
                lista_lap.append(lap_data)
            # ... resto del codice della funzione ...
        
        sport = dati_attivita.get('sport')
        sub_sport = dati_attivita.get('sub_sport')
        # ... resto del codice della funzione ...

        riga_storico = {
            # ...
        }
        righe_lap = df_lap[['Data_Attività', 'Numero_Lap', 'tipo_lap', 'Distanza_Lap (m)', 'Tempo_Lap', 'Andatura_Lap (min/km)', 'FC_Media_Lap (bpm)']]

        return riga_storico, righe_lap

    except Exception as e:
        print(f"ERRORE CRITICO durante l'analisi: {e}")
        return None, None