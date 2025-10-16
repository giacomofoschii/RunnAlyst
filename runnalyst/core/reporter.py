# runnalyst/core/reporter.py
import pandas as pd
from .parser import formatta_passo # Importiamo la funzione di formattazione

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

def crea_report_excel(output_path, df_storico, df_lap):
    """Scrive i DataFrame nel file Excel finale."""
    df_sett, df_mese, df_anno = calcola_riepiloghi(df_storico)
    
    print("\nScrittura del file Excel in corso...")
    with pd.ExcelWriter(output_path, engine='openpyxl', date_format='YYYY-MM-DD') as writer:
        df_storico.to_excel(writer, sheet_name='Storico Allenamenti', index=False)
        df_lap.to_excel(writer, sheet_name='Dettaglio Lap', index=False)
        df_sett.to_excel(writer, sheet_name='Riepiloghi', index=False, startrow=1)
        df_mese.to_excel(writer, sheet_name='Riepiloghi', index=False, startrow=len(df_sett)+5)
        df_anno.to_excel(writer, sheet_name='Riepiloghi', index=False, startrow=len(df_sett)+len(df_mese)+9)
        
        workbook = writer.book
        sheet = writer.sheets['Riepiloghi']
        sheet['A1'] = 'Riepilogo Settimanale'
        sheet[f'A{len(df_sett)+5}'] = 'Riepilogo Mensile'
        sheet[f'A{len(df_sett)+len(df_mese)+9}'] = 'Riepilogo Annuale'

    print(f"\nReport creato/aggiornato con successo!")
    print(f"Lo trovi qui: {output_path}")