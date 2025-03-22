import pandas as pd
import numpy as np
from collections import defaultdict
from itertools import combinations
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import os

class LogoSimilarityAnalyzer:
    def __init__(self, parquet_file):
        if not os.path.exists(parquet_file):
            raise FileNotFoundError(f"Fișierul {parquet_file} nu a fost găsit!")
            
        print(f"Se încarcă datele din {parquet_file}...")
        self.df = pd.read_parquet(parquet_file)
        print(f"Date încărcate cu succes: {len(self.df)} înregistrări")
        
        # Afișăm informații despre structura datelor
        print("\nStructura datelor:")
        print(f"Coloane disponibile: {', '.join(self.df.columns)}")
        
    def get_common_letters(self, word1, word2):
        """Calculează numărul de litere comune între două cuvinte"""
        return set(word1.lower()) & set(word2.lower())
    
    def extract_domains(self):
        """Extrage domeniile din dataset"""
        potential_columns = [col for col in self.df.columns if isinstance(self.df[col].iloc[0], str)]
        
        if not potential_columns:
            raise ValueError("Nu s-au găsit coloane cu text în fișierul parquet!")
            
        self.domain_column = potential_columns[0]
        domains = self.df[self.domain_column].dropna().tolist()
        
        print(f"\nAm extras {len(domains)} domenii din coloana '{self.domain_column}'")
        print(f"Exemplu de domenii: {', '.join(domains[:5])}...")
        
        return domains

    def analyze_similarity_levels(self):
        """Analizează și grupează domeniile pe nivele de similaritate"""
        print("\nÎncepe analiza similarității...")
        domains = self.extract_domains()
        
        if not domains:
            raise ValueError("Nu s-au găsit domenii pentru analiză!")
        
        # Dicționare pentru fiecare nivel de similaritate
        max_similarity = []     # 4+ litere comune
        medium_similarity = []  # 2-3 litere comune
        basic_similarity = []   # 1 literă comună
        
        total_combinations = sum(1 for _ in combinations(domains, 2))
        print(f"\nAnalizăm {total_combinations} combinații posibile de domenii...")
        
        # Analizăm toate perechile posibile
        for domain1, domain2 in combinations(domains, 2):
            common_letters = self.get_common_letters(domain1, domain2)
            num_common = len(common_letters)
            
            similarity_data = {
                'Domeniu 1': domain1,
                'Domeniu 2': domain2,
                'Litere comune': ', '.join(sorted(common_letters)),
                'Număr litere comune': num_common
            }
            
            if num_common >= 4:
                max_similarity.append(similarity_data)
            elif num_common in [2, 3]:
                medium_similarity.append(similarity_data)
            elif num_common == 1:
                basic_similarity.append(similarity_data)
        
        print("\nRezultate preliminare:")
        print(f"- Similaritate maximă (4+ litere): {len(max_similarity)} perechi")
        print(f"- Similaritate medie (2-3 litere): {len(medium_similarity)} perechi")
        print(f"- Similaritate minimă (1 literă): {len(basic_similarity)} perechi")
        
        return max_similarity, medium_similarity, basic_similarity

    def export_similarity_analysis(self):
        """Exportă analizele în trei fișiere Excel separate"""
        try:
            max_pairs, medium_pairs, basic_pairs = self.analyze_similarity_levels()
            
            print("\nExportăm rezultatele în fișiere Excel...")
            
            # 1. Max Similarity (4+ litere comune)
            if max_pairs:
                df_max = pd.DataFrame(max_pairs)
                df_max = df_max.sort_values('Număr litere comune', ascending=False)
                
                with pd.ExcelWriter('Max_SimilarityLogos.xlsx', engine='openpyxl') as writer:
                    df_max.to_excel(writer, sheet_name='Similaritate Maximă', index=False)
                    
                    stats = pd.DataFrame({
                        'Metric': ['Total perechi', 'Medie litere comune', 'Maxim litere comune'],
                        'Valoare': [
                            len(df_max),
                            round(df_max['Număr litere comune'].mean(), 2),
                            df_max['Număr litere comune'].max()
                        ]
                    })
                    stats.to_excel(writer, sheet_name='Statistici', index=False)
                print(f"✓ Max_SimilarityLogos.xlsx creat cu {len(df_max)} perechi")
            
            # 2. Medium Similarity (2-3 litere comune)
            if medium_pairs:
                df_medium = pd.DataFrame(medium_pairs)
                df_medium = df_medium.sort_values('Număr litere comune', ascending=False)
                
                with pd.ExcelWriter('Medium_SimilarityLogos.xlsx', engine='openpyxl') as writer:
                    df_medium.to_excel(writer, sheet_name='Similaritate Medie', index=False)
                    
                    stats = pd.DataFrame({
                        'Metric': ['Total perechi', 'Medie litere comune'],
                        'Valoare': [
                            len(df_medium),
                            round(df_medium['Număr litere comune'].mean(), 2)
                        ]
                    })
                    stats.to_excel(writer, sheet_name='Statistici', index=False)
                print(f"✓ Medium_SimilarityLogos.xlsx creat cu {len(df_medium)} perechi")
            
            # 3. Basic Similarity (1 literă comună)
            if basic_pairs:
                df_basic = pd.DataFrame(basic_pairs)
                
                with pd.ExcelWriter('Basic_SimilarityLogos.xlsx', engine='openpyxl') as writer:
                    df_basic.to_excel(writer, sheet_name='Similaritate Minimă', index=False)
                    
                    stats = pd.DataFrame({
                        'Metric': ['Total perechi'],
                        'Valoare': [len(df_basic)]
                    })
                    stats.to_excel(writer, sheet_name='Statistici', index=False)
                print(f"✓ Basic_SimilarityLogos.xlsx creat cu {len(df_basic)} perechi")
            
            print("\nAnaliza completă! Fișierele au fost create cu succes.")
            
        except Exception as e:
            print(f"\nEroare în timpul analizei: {str(e)}")
            raise

# Exemplu de utilizare
if __name__ == "__main__":
    try:
        parquet_file = 'logos.snappy(2).parquet'
        print(f"\nÎncepe analiza fișierului: {parquet_file}")
        analyzer = LogoSimilarityAnalyzer(parquet_file)
        analyzer.export_similarity_analysis()
    except Exception as e:
        print(f"\nEroare: {str(e)}")
        print("Programul s-a oprit din cauza unei erori.")
