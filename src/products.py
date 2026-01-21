import csv
from pathlib import Path

# products.py
import csv
from pathlib import Path

def load_products(csv_file: Path):
    products = []
    with open(csv_file, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            products.append({
                "codigo": row["codigo"],
                "descripcion": row["Descripcion"],
                "cantidad": int(row["Cantidad"]),
                "precio_unitario": float(row["PrecioUnitario"]),
                "precio_total": float(row["PrecioTotal"])
            })
    return products
