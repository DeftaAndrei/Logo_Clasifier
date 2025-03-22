import pandas as pd
import numpy as np
from itertools import combinations
import os
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from fuzzywuzzy import fuzz, process
from collections import defaultdict

class LogoAnalyzer:
    def __init__(self, parquet_file='logos.snappy(2).parquet'):
        self.parquet_file = parquet_file
        self.domains = []
        self.company_names = []
        self.perfect_matches = []
        self.medium_matches = []
        self.similar_matches = []
        self.similar_companies = defaultdict(list)
        self.SIMILARITY_THRESHOLD = 85

    def extract_company_name(self, domain):
        """Extrage numele companiei din domeniu."""
        parts = str(domain).split(".")
        if len(parts) > 2:
            return parts[-3]  # Returnează partea principală a domeniului
        return parts[0]

    def load_data(self):
        """Încarcă și validează datele din fișierul parquet."""
        if not os.path.exists(self.parquet_file):
            raise FileNotFoundError(f"Nu s-a găsit fișierul {self.parquet_file}")
        
        df = pd.read_parquet(self.parquet_file)
        if df.empty:
            raise ValueError("Fișierul parquet nu conține date")
            
        self.domains = df.iloc[:, 0].dropna().tolist()
        if not self.domains:
            raise ValueError("Nu s-au găsit domenii valide pentru analiză")
        
        # Extragem numele companiilor
        self.company_names = [self.extract_company_name(domain) for domain in self.domains]
        self.company_names = list(set(self.company_names))  # Eliminăm duplicatele
        
        return True

    def find_similar_pairs(self):
        """Găsește perechi de domenii cu litere comune."""
        for domain1, domain2 in combinations(self.domains, 2):
            common_letters = set(str(domain1).lower()) & set(str(domain2).lower())
            num_common = len(common_letters)
            
            if num_common == 0:
                continue
                
            pair_info = {
                'Domeniu 1': domain1,
                'Domeniu 2': domain2,
                'Litere comune': ', '.join(sorted(common_letters)),
                'Număr litere comune': num_common
            }
            
            if num_common >= 4:
                self.perfect_matches.append(pair_info)
            elif num_common in [2, 3]:
                self.medium_matches.append(pair_info)
            elif num_common == 1:
                self.similar_matches.append(pair_info)

    def find_similar_companies(self):
        """Găsește companii cu nume similare folosind fuzzy matching."""
        for name in self.company_names:
            matches = process.extract(name, self.company_names, scorer=fuzz.token_sort_ratio, limit=10)
            similar_group = [match[0] for match in matches if match[1] >= self.SIMILARITY_THRESHOLD]
            
            if len(similar_group) > 1:
                self.similar_companies[name] = similar_group

    def save_to_excel(self, data, filename, sheet_name='Date'):
        """Salvează rezultatele într-un fișier Excel."""
        if not data:
            return 0
            
        df = pd.DataFrame(data)
        if 'Număr litere comune' in df.columns:
            df = df.sort_values('Număr litere comune', ascending=False)
        
        with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            if 'Număr litere comune' in df.columns:
                stats_data = {
                    'Metric': ['Total perechi', 'Medie litere comune'],
                    'Valoare': [len(df), round(df['Număr litere comune'].mean(), 2)]
                }
                pd.DataFrame(stats_data).to_excel(writer, sheet_name='Statistici', index=False)
        
        return len(df)

    def analyze(self):
        """Rulează analiza completă și salvează rezultatele."""
        try:
            self.load_data()
            
            # Analiza literelor comune
            self.find_similar_pairs()
            
            # Analiza numelor similare
            self.find_similar_companies()
            
            results = []
            
            # Salvăm rezultatele analizei literelor comune
            for matches, filename in [
                (self.perfect_matches, 'Perfect.xlsx'),
                (self.medium_matches, 'Medium.xlsx'),
                (self.similar_matches, 'Similar.xlsx')
            ]:
                count = self.save_to_excel(matches, filename)
                if count > 0:
                    results.append(f"{filename}: {count} perechi")
            
            # Salvăm rezultatele analizei numelor similare
            similar_companies_df = pd.DataFrame(list(self.similar_companies.items()),
                                             columns=["Companie", "Companii Similare"])
            count = self.save_to_excel(similar_companies_df, 'CompaniiSimilare.xlsx', 'Nume Similare')
            if count > 0:
                results.append(f"CompaniiSimilare.xlsx: {count} grupuri")
            
            if results:
                print("Fișiere create cu succes:")
                for result in results:
                    print(f"✓ {result}")
            else:
                print("Nu s-au găsit suficiente perechi pentru analiză")
                
        except Exception as e:
            print(f"Eroare: {str(e)}")

if __name__ == "__main__":
    analyzer = LogoAnalyzer()
    analyzer.analyze() 