import os
import pandas as pd
from fitparse import FitFile
from datetime import timedelta

def formatta_passo(passo_decimale):
    """Converte il passo da minuti decimali a formato MM:SS."""
    if pd.isna(passo_decimale) or passo_decimale == float('inf'):
        return "00:00"
    minuti = int(passo_decimale)
    secondi = int((passo_decimale * 60) % 60)
    return f"{minuti:02d}:{secondi:02d}"

def formatta_durata(secondi_totali):
    """Converte i secondi totali in formato HH:MM:SS."""
    if pd.isna(secondi_totali):
        return "00:00:00"
    return str(timedelta(seconds=int(secondi_totali)))

def analizza_file_fit(percorso_file):
    """Estrae tutti i dati necessari da un singolo file .fit."""
    try:
        fitfile = FitFile(percorso_file) # <-- ECCO DOVE VIENE USATO "FitFile"
        
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
            if record.name == 'event' and record.get_value('event_type') == 'activity':
                 dati_attivita['activity_name'] = record.get_value('data')
            if record.name == 'file_id' and 'product_name' in [f.name for f in record.fields]:
                 dati_attivita['activity_name'] = record.get_value('product_name')

        start_time = dati_attivita.get('start_time')
        distanza_totale_m = dati_attivita.get('total_distance', 0)
        durata_totale_s = dati_attivita.get('total_elapsed_time', 0)
        
        sport = dati_attivita.get('sport')
        sub_sport = dati_attivita.get('sub_sport')
        sport_tradotto = "Sconosciuto"

        if sport == 'running':
            if sub_sport == 'track': sport_tradotto = 'Corsa su Pista'
            else: sport_tradotto = 'Corsa'
        elif sport == 'trail_running': sport_tradotto = 'Trail Running'
        elif sport == 'virtual_running': sport_tradotto = 'Corsa Virtuale'
        else: sport_tradotto = str(sport).replace('_', ' ').title() if sport else 'N/D'

        distanza_totale_km = distanza_totale_m / 1000 if distanza_totale_m else 0
        andatura_media_totale = (durata_totale_s / 60) / distanza_totale_km if distanza_totale_km > 0 else 0
        
        df_lap = pd.DataFrame(lista_lap)
        if df_lap.empty: return None, None
        
        df_lap['tipo_lap'] = df_lap['intensity'].apply(lambda x: 'Corsa' if x == 'active' else 'Riposo')
        
        riga_storico = {
            'Data': start_time.date() if start_time else None,
            'Nome Attività': dati_attivita.get('activity_name', 'N/D'),
            'Tipo Sport': sport_tradotto,
            'Distanza Totale (km)': round(distanza_totale_km, 2),
            'Durata Totale': formatta_durata(durata_totale_s),
            'Andatura Media Totale': formatta_passo(andatura_media_totale),
            'FC Media (bpm)': int(dati_attivita.get('avg_heart_rate', 0)),
            'File_ID': start_time.strftime('%Y-%m-%d_%H-%M-%S') if start_time else os.path.basename(percorso_file)
        }
        
        df_lap['Data_Attività'] = start_time.date() if start_time else None
        df_lap['Numero_Lap'] = df_lap.index + 1
        df_lap['Distanza_Lap (m)'] = df_lap['total_distance'].round(2)
        df_lap['Tempo_Lap'] = df_lap['total_elapsed_time'].apply(formatta_durata)
        df_lap['Andatura_Lap (min/km)'] = df_lap.apply(lambda row: formatta_passo((row['total_elapsed_time']/60) / (row['total_distance']/1000) if row['total_distance'] > 0 else 0), axis=1)
        df_lap['FC_Media_Lap (bpm)'] = df_lap['avg_heart_rate'].fillna(0).astype(int)
        
        righe_lap = df_lap[['Data_Attività', 'Numero_Lap', 'tipo_lap', 'Distanza_Lap (m)', 'Tempo_Lap', 'Andatura_Lap (min/km)', 'FC_Media_Lap (bpm)']]

        return riga_storico, righe_lap

    except Exception as e:
        print(f"ERRORE CRITICO durante l'analisi di {os.path.basename(percorso_file)}: {e}")
        return None, None