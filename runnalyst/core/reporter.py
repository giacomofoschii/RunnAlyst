import pandas as pd

def calcola_riepiloghi(df_storico):
    # ... (copia la funzione calcola_riepiloghi da analizza_corse.py) ...
    # ...
    return df_sett, df_mese, df_anno

def crea_report_excel(output_path, df_storico, df_lap):
    """Scrive i DataFrame nel file Excel finale."""
    df_sett, df_mese, df_anno = calcola_riepiloghi(df_storico)
    
    print("\nScrittura del file Excel in corso...")
    with pd.ExcelWriter(output_path, engine='openpyxl', date_format='YYYY-MM-DD') as writer:
        df_storico.to_excel(writer, sheet_name='Storico Allenamenti', index=False)
        df_lap.to_excel(writer, sheet_name='Dettaglio Lap', index=False)
        # ... resto della logica di scrittura ...
    print(f"\nReport creato/aggiornato con successo!")