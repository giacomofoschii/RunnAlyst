# runnalyst/cli.py
import os
import pandas as pd
from .core.parser import analizza_file_fit
from .core.reporter import crea_report_excel

# --- CONFIGURAZIONE ---
CARTELLA_BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
FILE_EXCEL_OUTPUT = os.path.join(CARTELLA_BASE, "Report_Corsa.xlsx")
CARTELLA_FIT = CARTELLA_BASE
# --------------------

def main():
    print(f"Avvio analisi allenamenti in RunnAlyst...")
    
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
        if '.venv' in root or '.git' in root: # Ignora le cartelle di servizio
            continue
        for file in files:
            if file.lower().endswith('.fit'):
                elenco_file_fit.append(os.path.join(root, file))
    
    if not elenco_file_fit:
        print("\nATTENZIONE: Nessun file .fit trovato.")
        return
    
    print(f"Trovati {len(elenco_file_fit)} file .fit. Inizio la scansione...")
    
    nuove_righe_storico = []
    lista_df_lap_nuovi = []

    for file_path in elenco_file_fit:
        try:
            from fitparse import FitFile
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
                if df_lap is not None:
                    lista_df_lap_nuovi.append(df_lap)

    if not nuove_righe_storico:
        print("\nNessuna nuova attività da aggiungere. Report già aggiornato.")
        return

    df_storico_nuovo = pd.DataFrame(nuove_righe_storico)
    df_storico_finale = pd.concat([df_storico_esistente, df_storico_nuovo], ignore_index=True)
    
    if 'Tipo di Allenamento' not in df_storico_finale.columns:
        df_storico_finale.insert(1, 'Tipo di Allenamento', '')
    if 'Descrizione' not in df_storico_finale.columns:
        df_storico_finale.insert(2, 'Descrizione', '')
        
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
    
    crea_report_excel(FILE_EXCEL_OUTPUT, df_storico_finale, df_lap_finale)

if __name__ == "__main__":
    main()