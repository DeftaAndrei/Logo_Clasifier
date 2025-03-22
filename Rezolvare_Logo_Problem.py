import pandas as pd
import numpy as np
from itertools import combinations
import os
from fuzzywuzzy import fuzz, process
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class LogoAnalyzer:
    def __init__(self):
        self.parquet_file = 'logos.snappy(2).parquet'  # Specificăm exact fișierul cu care lucrăm
        self.df = None
        self.domains = []
        self.company_names = []
        self.analysis_results = {
            'letter_similarity': defaultdict(list),
            'name_similarity': defaultdict(list),
            'domain_patterns': defaultdict(list),
            'statistics': {},
            'parquet_info': {}
        }

    def load_and_clean_data(self):
        """Încarcă și curăță datele din fișierul logos.snappy(2).parquet."""
        print(f"\nVerificăm fișierul {self.parquet_file}...")
        
        if not os.path.exists(self.parquet_file):
            raise FileNotFoundError(f"EROARE: Nu s-a găsit fișierul {self.parquet_file}")
        
        print("✓ Fișierul parquet există")
        print("\nÎncărcăm datele din parquet...")
        
        # Încărcăm datele
        self.df = pd.read_parquet(self.parquet_file)
        
        # Verificăm structura datelor
        self.analysis_results['parquet_info'] = {
            'Număr total înregistrări': len(self.df),
            'Coloane disponibile': list(self.df.columns),
            'Dimensiune fișier (bytes)': os.path.getsize(self.parquet_file)
        }
        
        if self.df.empty:
            raise ValueError("EROARE: Fișierul parquet nu conține date")
        
        print(f"✓ Date încărcate cu succes: {len(self.df)} înregistrări")
        
        # Extragem și curățăm domeniile
        print("\nProcesăm domeniile...")
        self.domains = self.df.iloc[:, 0].dropna().astype(str).tolist()
        print(f"✓ {len(self.domains)} domenii valide găsite")
        
        # Extragem numele companiilor
        print("\nExtragem numele companiilor...")
        self.company_names = [self.extract_company_name(d) for d in self.domains]
        self.company_names = list(set(filter(None, self.company_names)))
        print(f"✓ {len(self.company_names)} nume unice de companii extrase")

    def extract_company_name(self, domain):
        """Extrage și curăță numele companiei din domeniu."""
        try:
            parts = domain.lower().split('.')
            if len(parts) > 2:
                return parts[-3]  # Luăm partea principală a domeniului
            return parts[0]
        except:
            return None

    def analyze_letter_similarity(self):
        """Analizează similaritatea bazată pe litere comune între domenii."""
        print("\nAnalizăm similaritatea literelor între domenii...")
        total_combinations = sum(1 for _ in combinations(self.domains, 2))
        
        for domain1, domain2 in combinations(self.domains, 2):
            common_letters = set(domain1.lower()) & set(domain2.lower())
            num_common = len(common_letters)
            
            if num_common == 0:
                continue
                
            similarity_info = {
                'Domeniu 1': domain1,
                'Domeniu 2': domain2,
                'Litere comune': ', '.join(sorted(common_letters)),
                'Număr litere comune': num_common,
                'Procent similaritate': round(num_common / max(len(domain1), len(domain2)) * 100, 2)
            }
            
            if num_common >= 4:
                self.analysis_results['letter_similarity']['perfect'].append(similarity_info)
            elif num_common in [2, 3]:
                self.analysis_results['letter_similarity']['medium'].append(similarity_info)
            elif num_common == 1:
                self.analysis_results['letter_similarity']['basic'].append(similarity_info)
        
        print("✓ Analiză similaritate litere completă")

    def analyze_name_similarity(self):
        """Analizează similaritatea între numele companiilor."""
        print("\nAnalizăm similaritatea între numele companiilor...")
        SIMILARITY_THRESHOLD = 85
        
        for name in self.company_names:
            matches = process.extract(name, self.company_names, scorer=fuzz.token_sort_ratio, limit=10)
            similar_names = [(match[0], match[1]) for match in matches if match[1] >= SIMILARITY_THRESHOLD and match[0] != name]
            
            if similar_names:
                self.analysis_results['name_similarity']['groups'].append({
                    'Nume companie': name,
                    'Nume similare': [n[0] for n in similar_names],
                    'Scoruri similaritate': [n[1] for n in similar_names]
                })
        
        print("✓ Analiză similaritate nume completă")

    def analyze_domain_patterns(self):
        """Analizează tipare în structura domeniilor."""
        print("\nAnalizăm structura domeniilor...")
        
        for domain in self.domains:
            parts = domain.split('.')
            tld = parts[-1] if len(parts) > 0 else ''
            subdomain_count = len(parts) - 2 if len(parts) > 2 else 0
            
            self.analysis_results['domain_patterns']['tlds'].append(tld)
            self.analysis_results['domain_patterns']['structure'].append({
                'Domain': domain,
                'TLD': tld,
                'Număr subdomenii': subdomain_count,
                'Lungime': len(domain)
            })
        
        print("✓ Analiză structură domenii completă")

    def calculate_statistics(self):
        """Calculează statistici generale despre datele analizate."""
        print("\nCalculăm statisticile generale...")
        
        # Statistici despre fișierul parquet
        parquet_stats = self.analysis_results['parquet_info']
        
        # Statistici despre analiză
        analysis_stats = {
            'Total domenii': len(self.domains),
            'Total companii unice': len(self.company_names),
            'Perechi perfecte (4+ litere)': len(self.analysis_results['letter_similarity']['perfect']),
            'Perechi medii (2-3 litere)': len(self.analysis_results['letter_similarity']['medium']),
            'Perechi basic (1 literă)': len(self.analysis_results['letter_similarity']['basic']),
            'Grupuri nume similare': len(self.analysis_results['name_similarity']['groups']),
            'TLD-uri unice': len(set(self.analysis_results['domain_patterns']['tlds']))
        }
        
        self.analysis_results['statistics'] = {**parquet_stats, **analysis_stats}
        print("✓ Statistici calculate")

    def save_results(self):
        """Salvează rezultatele analizei în fișiere Excel."""
        print("\nSalvăm rezultatele analizei...")
        
        # Creăm un director pentru rezultate
        output_dir = "Rezultate_Analiza_Logo"
        os.makedirs(output_dir, exist_ok=True)
        
        # 1. Salvăm rezultatele similarității literelor
        for level, data in self.analysis_results['letter_similarity'].items():
            if data:
                filename = os.path.join(output_dir, f'Similaritate_{level.capitalize()}.xlsx')
                df = pd.DataFrame(data)
                df.to_excel(filename, index=False)
                print(f"✓ Salvat {filename}")

        # 2. Salvăm rezultatele similarității numelor
        if self.analysis_results['name_similarity']['groups']:
            filename = os.path.join(output_dir, 'Similaritate_Nume.xlsx')
            df_names = pd.DataFrame(self.analysis_results['name_similarity']['groups'])
            df_names.to_excel(filename, index=False)
            print(f"✓ Salvat {filename}")

        # 3. Salvăm analiza pattern-urilor
        filename = os.path.join(output_dir, 'Analiza_Domenii.xlsx')
        df_patterns = pd.DataFrame(self.analysis_results['domain_patterns']['structure'])
        df_patterns.to_excel(filename, index=False)
        print(f"✓ Salvat {filename}")

        # 4. Salvăm statisticile generale
        filename = os.path.join(output_dir, 'Statistici_Generale.xlsx')
        df_stats = pd.DataFrame(list(self.analysis_results['statistics'].items()),
                              columns=['Metric', 'Valoare'])
        df_stats.to_excel(filename, index=False)
        print(f"✓ Salvat {filename}")

    def run_analysis(self):
        """Rulează întreaga analiză."""
        try:
            print("=== Începem analiza logo-urilor din logos.snappy(2).parquet ===")
            
            # Încărcăm și curățăm datele
            self.load_and_clean_data()
            
            # Rulăm toate analizele
            self.analyze_letter_similarity()
            self.analyze_name_similarity()
            self.analyze_domain_patterns()
            self.calculate_statistics()
            
            # Salvăm rezultatele
            self.save_results()
            
            # Afișăm statisticile finale
            print("\n=== Statistici finale ===")
            for metric, value in self.analysis_results['statistics'].items():
                print(f"• {metric}: {value}")
            
            print("\n=== Analiză completă! ===")
            print("Toate rezultatele au fost salvate în directorul 'Rezultate_Analiza_Logo'")
                
        except Exception as e:
            print(f"\nEROARE în timpul analizei: {str(e)}")

if __name__ == "__main__":
    analyzer = LogoAnalyzer()
    analyzer.run_analysis() 