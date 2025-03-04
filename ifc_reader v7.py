import ifcopenshell
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import filedialog
import os

# 📌 Mevcut script'in olduğu dizini al
script_dir = os.path.dirname(os.path.abspath(__file__))

# 📌 Kullanıcıdan IFC dosyasını al
root = tk.Tk()
root.withdraw()
ifc_file_path = filedialog.askopenfilename(title="Select IFC File", filetypes=[("IFC files", "*.ifc")])

if not ifc_file_path:
    print("❌ No file selected. Exiting...")
    exit()

# 📌 IFC Dosyasını Aç
ifc_file = ifcopenshell.open(ifc_file_path)

# 📌 OmniClass Mapping TXT Dosyasını Oku
omniclass_mapping_path = os.path.join(script_dir, "omniclass_mapping.txt")
omniclass_mapping = {}

with open(omniclass_mapping_path, "r", encoding="utf-8") as file:
    for line in file:
        if ":" in line:
            ifc_type, omniclass_desc = line.strip().split(":", 1)
            omniclass_mapping[ifc_type.strip()] = omniclass_desc.strip()

# 📌 Firma ve Ürün Bilgilerini CSV'den Oku
company_data_path = os.path.join(script_dir, "company_data.csv")
company_df = pd.read_csv(company_data_path, delimiter=";", encoding="utf-8")

# 📌 Beklenen CSV Sütunlarını Kontrol Et
expected_columns = ["OmniClass", "Company", "Product", "Product Code", "Price", "CO2_Emissions", "Lead Time"]
missing_columns = [col for col in expected_columns if col not in company_df.columns]

if missing_columns:
    print(f"❌ ERROR: Missing columns in CSV: {missing_columns}")
    print("⚠️ Please check your CSV file and make sure it has the correct column names!")
    exit()

# 📌 IFC Objelerini Listele
all_types = sorted(set(entity.is_a() for entity in ifc_file.by_type("IfcProduct")))

# ✅ OmniClass Eşleşmesi Olanlar Filtrelendi
filtered_types = [t for t in all_types if t in omniclass_mapping]

print("\n✅ IFC Objects with OmniClass Mapping:")
for i, obj_type in enumerate(filtered_types, start=1):
    print(f"{i} - {omniclass_mapping[obj_type]}")  # Sadece OmniClass Açıklaması Göster

# 📌 Kullanıcıdan sıralama kriteri al (Kısayollar: P, C, T)
sort_shortcut = input("\nSort by (P = Price, C = CO2_Emissions, T = Lead Time)? ").strip().upper()

# 📌 Kullanıcının girişini uygun sütun ismine çevir
sort_mapping = {
    "P": ("Price", "price"),
    "C": ("CO2_Emissions", "co2"),
    "T": ("Lead Time", "leadtime"),
}

if sort_shortcut not in sort_mapping:
    print("⚠️ Invalid input! Defaulting to 'P' (Price).")
    sort_shortcut = "P"

sort_criteria, file_suffix = sort_mapping[sort_shortcut]

# 📌 Objeleri Çek ve OmniClass ile Eşleştir
data = []

# Paralel işlem ile veriyi hızlı çek
def process_element(element, obj_type):
    omniclass_desc = omniclass_mapping.get(obj_type, "Unknown")  # OmniClass Açıklamasını al
    name = getattr(element, "Name", "Unknown")  # 🔥 Name eksik olmasın
    
    # **ObjectType Verisini Al**
    object_type = getattr(element, "ObjectType", "Unknown")  # 🔥 ObjectType düzeltildi

    # OmniClass'a uygun firmaları filtrele
    matched_firms = company_df[company_df["OmniClass"] == omniclass_desc]

    if matched_firms.empty:
        firm = {"Company": "No Supplier", "Product": "N/A", "Product Code": "N/A", "Price": "N/A", "CO2_Emissions": "N/A", "Lead Time": "N/A"}
    else:
        # Seçilen kritere göre firmaları sırala ve en iyisini seç
        firm = matched_firms.sort_values(by=sort_criteria, ascending=True).iloc[0]

    return {
        "OmniClass": omniclass_desc,
        "Object Type": object_type,  # 🔥 Artık gerçek ObjectType geliyor!
        "Name": name,  # 🔥 Name eksik olmasın
        "Company": firm["Company"],
        "Product": firm["Product"],
        "Product Code": firm["Product Code"],
        "Price": firm["Price"],
        "CO2_Emissions": firm["CO2_Emissions"],
        "Lead Time": firm["Lead Time"]
    }

# 🚀 Çoklu işlem (Multithreading) kullanarak hızlandır
with ThreadPoolExecutor() as executor:
    for obj_type in filtered_types:
        elements = ifc_file.by_type(obj_type)
        results = executor.map(lambda element: process_element(element, obj_type), elements)
        data.extend(results)

# 📌 Sadece **ObjectType** bazında benzersiz objeleri ve sayıları al
df = pd.DataFrame(data)

# 📌 **COUNT Sütunu ObjectType Bazında Olsun**
df["Count"] = df.groupby(["Object Type"])["Object Type"].transform("count")  

# 📌 **Benzersiz ObjectType Olanları Al**
df = df.drop_duplicates(subset=["Object Type"])  

# 📌 **COUNT Sütununu Name'den Sonra Koy**
column_order = ["OmniClass", "Object Type", "Name", "Count", "Company", "Product", "Product Code", "Price", "CO2_Emissions", "Lead Time"]
df = df[column_order]

# 📌 Kullanıcının seçtiği kritere göre dosya ismini belirle
output_file_name = f"ifc_unique_{file_suffix}.csv"
output_file_path = os.path.join(script_dir, output_file_name)

# 📌 CSV Kaydet
df.to_csv(output_file_path, index=False, sep=";", encoding="utf-8")

print(f"✅ Process complete! Unique objects CSV created:\n🔹 {output_file_path}")
