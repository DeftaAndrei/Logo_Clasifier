import pandas as pd
from fuzzywuzzy import fuzz, process
from collections import defaultdict

# Încarcă fișierul Excel (schimbă "output.xlsx" cu calea fișierului tău)
file_path = "output.xlsx"
df = pd.read_excel(file_path)

# Funcție pentru a extrage numele companiei din domeniu
def extract_company_name(domain):
    parts = domain.split(".")
    if len(parts) > 2:
        return parts[-3]  # Returnează partea principală a domeniului
    return parts[0]

# Aplică funcția pe coloana domeniilor
df["company_name"] = df["domain"].astype(str).apply(extract_company_name)

# Setăm un prag pentru similaritate (85% este un punct de referință bun)
SIMILARITY_THRESHOLD = 85

# Listă de nume unice de companii
company_names = df["company_name"].unique()

# Dicționar pentru gruparea companiilor similare
similar_companies = defaultdict(list)

# Comparăm fiecare nume de companie cu celelalte
for name in company_names:
    matches = process.extract(name, company_names, scorer=fuzz.token_sort_ratio, limit=10)
    similar_group = [match[0] for match in matches if match[1] >= SIMILARITY_THRESHOLD]
    
    # Salvăm rezultatul în dicționar
    if len(similar_group) > 1:
        similar_companies[name] = similar_group

# Convertim rezultatul într-un DataFrame
similar_companies_df = pd.DataFrame(list(similar_companies.items()), columns=["Company", "Similar Companies"])

# Salvăm rezultatul într-un fișier Excel
similar_companies_df.to_excel("similar_companies.xlsx", index=False)

# Afișăm primele 10 rezultate
print(similar_companies_df.head(10))
