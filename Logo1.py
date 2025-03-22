import pandas as pd
import numpy as np
from itertools import combinations
import os
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

def analyze_logos(parquet_file):
    """
    Analizează similaritățile între logouri din logos.snappy(2).parquet și generează Excel-uri:
    - Perfect.xlsx: 4+ litere comune
    - Medium.xlsx: 2-3 litere comune
    - Similar.xlsx: 1 literă comună
    """
    try:
        # Verificăm explicit existența fișierului parquet
        if not os.path.exists(parquet_file):
            print(f"Eroare: Nu s-a găsit fișierul {parquet_file}")
            return
        
        # Încărcăm datele din parquet
        df = pd.read_parquet(parquet_file)
        
        # Verificăm dacă avem date
        if df.empty:
            print("Eroare: Fișierul parquet nu conține date")
            return
            
        # Extragem domeniile din prima coloană
        domains = df.iloc[:, 0].dropna().tolist()
        
        if not domains:
            print("Eroare: Nu s-au găsit domenii valide pentru analiză")
            return
        
        # Dicționare pentru diferite nivele de similaritate
        perfect_matches = []    # 4+ litere comune
        medium_matches = []     # 2-3 litere comune
        similar_matches = []    # 1 literă comună
        
        # Analizăm fiecare pereche de domenii
        for domain1, domain2 in combinations(domains, 2):
            common_letters = set(str(domain1).lower()) & set(str(domain2).lower())
            num_common = len(common_letters)
            
            pair_info = {
                'Domeniu 1': domain1,
                'Domeniu 2': domain2,
                'Litere comune': ', '.join(sorted(common_letters)),
                'Număr litere comune': num_common
            }
            
            if num_common >= 4:
                perfect_matches.append(pair_info)
            elif num_common in [2, 3]:
                medium_matches.append(pair_info)
            elif num_common == 1:
                similar_matches.append(pair_info)

        def save_to_excel(data, filename):
            if not data:
                return 0
                
            df = pd.DataFrame(data)
            df = df.sort_values('Număr litere comune', ascending=False)
            
            # Folosim mode='w' pentru a suprascrie fișierul dacă există
            with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
                # Salvăm datele principale
                df.to_excel(writer, sheet_name='Date', index=False)
                
                # Adăugăm statistici
                stats_data = {
                    'Metric': ['Total perechi', 'Medie litere comune'],
                    'Valoare': [len(df), round(df['Număr litere comune'].mean(), 2)]
                }
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name='Statistici', index=False)
            
            return len(df)

        # Salvăm rezultatele în fișiere separate
        results = []
        
        # Perfect.xlsx
        count = save_to_excel(perfect_matches, 'Perfect.xlsx')
        if count > 0:
            results.append(f"Perfect.xlsx: {count} perechi")
            
        # Medium.xlsx
        count = save_to_excel(medium_matches, 'Medium.xlsx')
        if count > 0:
            results.append(f"Medium.xlsx: {count} perechi")
            
        # Similar.xlsx
        count = save_to_excel(similar_matches, 'Similar.xlsx')
        if count > 0:
            results.append(f"Similar.xlsx: {count} perechi")
        
        # Afișăm doar rezultatul final
        if results:
            print("Fișiere create cu succes:")
            for result in results:
                print(f"✓ {result}")
        else:
            print("Nu s-au găsit suficiente perechi pentru analiză")
            
    except Exception as e:
        print(f"A apărut o eroare în timpul procesării: {str(e)}")

if __name__ == "__main__":
    analyze_logos('logos.snappy(2).parquet')
