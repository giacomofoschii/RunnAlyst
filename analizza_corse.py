# -*- coding: utf-8 -*-
import os
import pandas as pd
from fitparse import FitFile
from datetime import timedelta

# --- CONFIGURAZIONE ---
# Percorso della cartella contenente le sottocartelle degli anni con i file .fit
CARTELLA_FIT = os.path.dirname(os.path.abspath(__file__))

# Percorso e nome del file Excel di output
FILE_EXCEL_OUTPUT = os.path.join(CARTELLA_FIT, "Report_Corsa.xlsx")
# --------------------

def formatta_passo(passo_decimale):
    """Converte il passo da minuti decimali (es. 5.5) a formato MM:SS (es. 05:30)."""
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
            if record.name == 'event' and record.get_value('event_type') == 'activity':
                 dati_attivita['activity_name'] = record.get_value('data')
            if record.name == 'file_id' and 'product_name' in [f.name for f in record.fields]:
                 dati_attivita['activity_name'] = record.get_value('product_name')

        start_time = dati_attivita.get('start_time')
        distanza_totale_m = dati_attivita.get('total_distance', 0)
        durata_totale_s = dati_attivita.get('total_elapsed_time', 0)
        
        # --- MODIFICA: Logica di classificazione basata su sport e sub_sport ---
        sport = dati_attivita.get('sport')
        sub_sport = dati_attivita.get('sub_sport')
        sport_tradotto = "Sconosciuto"

        if sport == 'running':
            if sub_sport == 'track':
                sport_tradotto = 'Corsa su Pista'
            elif sub_sport == 'indoor running':
                sport_tradotto = 'Corsa Indoor'
            elif sub_sport == 'trail':
                sport_tradotto = 'Trail Running'
            elif sub_sport == 'ultra':
                sport_tradotto = 'Ultra Running'
            elif sub_sport == 'treadmill':
                sport_tradotto = 'Corsa su Tapis Roulant'
        else:
            sport_tradotto = str(sport).replace('_', ' ').title() if sport else 'N/D'
        # --------------------------------------------------------------------

        distanza_totale_km = distanza_totale_m / 1000 if distanza_totale_m else 0
        andatura_media_totale = (durata_totale_s / 60) / distanza_totale_km if distanza_totale_km > 0 else 0
        
        df_lap = pd.DataFrame(lista_lap)
        if df_lap.empty:
            return None, None
        
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

def calcola_riepiloghi(df_storico):
    if df_storico.empty or 'Data' not in df_storico.columns:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df = df_storico.copy()
    df['Data'] = pd.to_datetime(df['Data'])
    
    def passo_a_decimale(passo_str):
        try:
            minuti, secondi = map(int, passo_str.split(':'))
            return minuti + secondi / 60
        except: return None
            
    df['Andatura_Totale_Dec'] = df['Andatura Media Totale'].apply(passo_a_decimale)
    
    def weighted_avg(group, avg_name, weight_name):
        d = group[avg_name]
        w = group[weight_name]
        try: return (d * w).sum() / w.sum()
        except ZeroDivisionError: return 0
    
    df['Settimana'] = df['Data'].dt.strftime('%Y-%U')
    df['Mese'] = df['Data'].dt.strftime('%Y-%m')
    df['Anno'] = df['Data'].dt.year

    df_sett = df.groupby('Settimana').apply(lambda g: pd.Series({'Km Totali': g['Distanza Totale (km)'].sum(), 'N° Allenamenti': g['Nome Attività'].count(),'Andatura Media Ponderata': formatta_passo(weighted_avg(g, 'Andatura_Totale_Dec', 'Distanza Totale (km)'))}), include_groups=False).round(2).reset_index()
    df_mese = df.groupby('Mese').apply(lambda g: pd.Series({'Km Totali': g['Distanza Totale (km)'].sum(), 'N° Allenamenti': g['Nome Attività'].count(),'Andatura Media Ponderata': formatta_passo(weighted_avg(g, 'Andatura_Totale_Dec', 'Distanza Totale (km)'))}), include_groups=False).round(2).reset_index()
    df_anno = df.groupby('Anno').apply(lambda g: pd.Series({'Km Totali': g['Distanza Totale (km)'].sum(), 'N° Allenamenti': g['Nome Attività'].count(),'Andatura Media Ponderata': formatta_passo(weighted_avg(g, 'Andatura_Totale_Dec', 'Distanza Totale (km)'))}), include_groups=False).round(2).reset_index()
    
    return df_sett, df_mese, df_anno

def main():
    print(f"Avvio analisi allenamenti...")
    
    try:
        df_storico_esistente = pd.read_excel(FILE_EXCEL_OUTPUT, sheet_name="Storico Allenamenti")
        df_lap_esistente = pd.read_excel(FILE_EXCEL_OUTPUT, sheet_name="Dettaglio Lap")
        file_gia_processati = set(df_storico_esistente['File_ID'])
        print(f"Trovato file Excel esistente con {len(file_gia_processati)} attività.")
    except FileNotFoundError:
        df_storico_esistente = pd.DataFrame()
        df_lap_esistente = pd.DataFrame()
        file_gia_processati = set()
        print("Nessun file Excel trovato. Verrà creato un nuovo report.")
    
    elenco_file_fit = []
    for root, _, files in os.walk(CARTELLA_FIT):
        for file in files:
            if file.lower().endswith('.fit'):
                elenco_file_fit.append(os.path.join(root, file))
    
    if not elenco_file_fit:
        print("\nATTENZIONE: Nessun file .fit trovato.")
        return
    
    print(f"Trovati {len(elenco_file_fit)} file .fit. Inizio la scansione...")
    
    nuove_righe_storico = []
    lista_df_lap_nuovi = []

    for i, file_path in enumerate(elenco_file_fit):
        try:
            fitfile_temp = FitFile(file_path)
            session = next(fitfile_temp.get_messages('session'))
            start_time = session.get_value('start_time')
            file_id = start_time.strftime('%Y-%m-%d_%H-%M-%S') if start_time else os.path.basename(file_path)
        except Exception:
            file_id = os.path.basename(file_path)

        if file_id not in file_gia_processati:
            print(f"  -> Analizzo nuovo file: {os.path.basename(file_path)}")
            riga_storico, df_lap = analizza_file_fit(file_path)
            if riga_storico:
                nuove_righe_storico.append(riga_storico)
                lista_df_lap_nuovi.append(df_lap)

    if not nuove_righe_storico:
        print("\nNessuna nuova attività da aggiungere. Report già aggiornato.")
        return

    df_storico_nuovo = pd.DataFrame(nuove_righe_storico)
    df_storico_finale = pd.concat([df_storico_esistente, df_storico_nuovo], ignore_index=True)
    
    if 'Tipo di Allenamento' not in df_storico_finale.columns:
        df_storico_finale['Tipo di Allenamento'] = ''
    if 'Descrizione' not in df_storico_finale.columns:
        df_storico_finale['Descrizione'] = ''
        
    colonne_ordinate = [
        'Data', 'Tipo di Allenamento', 'Descrizione', 'Nome Attività', 'Tipo Sport', 
        'Distanza Totale (km)', 'Durata Totale', 'Andatura Media Totale', 
        'FC Media (bpm)', 'File_ID'
    ]
    colonne_finali = [col for col in colonne_ordinate if col in df_storico_finale.columns]
    df_storico_finale = df_storico_finale[colonne_finali]
    
    df_lap_nuovi = pd.concat(lista_df_lap_nuovi, ignore_index=True) if lista_df_lap_nuovi else pd.DataFrame()
    df_lap_finale = pd.concat([df_lap_esistente, df_lap_nuovi], ignore_index=True)
    
    if 'File_ID' in df_storico_finale.columns:
         df_storico_finale = df_storico_finale.drop(columns=['File_ID'])
        
    df_storico_finale['Data'] = pd.to_datetime(df_storico_finale['Data'])
    df_storico_finale = df_storico_finale.sort_values(by='Data', ascending=False)
    
    if not df_lap_finale.empty:
        df_lap_finale['Data_Attività'] = pd.to_datetime(df_lap_finale['Data_Attività'])
        df_lap_finale = df_lap_finale.sort_values(by=['Data_Attività', 'Numero_Lap'], ascending=[False, True])

    df_sett, df_mese, df_anno = calcola_riepiloghi(df_storico_finale)
    
    print("\nScrittura del file Excel in corso...")
    with pd.ExcelWriter(FILE_EXCEL_OUTPUT, engine='openpyxl', date_format='YYYY-MM-DD') as writer:
        df_storico_finale.to_excel(writer, sheet_name='Storico Allenamenti', index=False)
        df_lap_finale.to_excel(writer, sheet_name='Dettaglio Lap', index=False)
        df_sett.to_excel(writer, sheet_name='Riepiloghi', index=False, startrow=1)
        df_mese.to_excel(writer, sheet_name='Riepiloghi', index=False, startrow=len(df_sett)+5)
        df_anno.to_excel(writer, sheet_name='Riepiloghi', index=False, startrow=len(df_sett)+len(df_mese)+9)
        
        workbook = writer.book
        sheet = writer.sheets['Riepiloghi']
        sheet['A1'] = 'Riepilogo Settimanale'
        sheet[f'A{len(df_sett)+5}'] = 'Riepilogo Mensile'
        sheet[f'A{len(df_sett)+len(df_mese)+9}'] = 'Riepilogo Annuale'

    print(f"\nReport creato/aggiornato con successo!")
    print(f"Lo trovi qui: {FILE_EXCEL_OUTPUT}")

if __name__ == "__main__":
    main()
    input("\nPremere Invio per chiudere...")