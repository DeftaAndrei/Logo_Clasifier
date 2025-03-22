import pandas as pd
import numpy as np

# Citim fișierul parquet
df = pd.read_parquet('logos.snappy(2).parquet')

# Afișăm primele câteva rânduri pentru a vedea structura datelor
print("Primele 5 rânduri din fișierul parquet:")
print(df.head())

# Salvăm datele în Excel
df.to_excel('output.xlsx', index=False)
print("\nDatele au fost salvate în fișierul 'output.xlsx'")
