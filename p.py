import pandas as pd


#Cek sheet apa aja yang ada
xl = pd.ExcelFile('Sample - Superstore (1) (2).xls', engine='xlrd')
print("--- Daftar Sheet ---")
print(xl.sheet_names) 
print("\n")

#Baca data, lihat kolom apa aja
df = pd.read_excel('Sample - Superstore (1) (2).xls', engine='xlrd')

print("--- Kolom yang Tersedia ---")
print(df.columns)
print(f"Ukuran Data: {df.shape[0]} baris, {df.shape[1]} kolom")
print("\n")


#Agregasi untuk dapat insight
print("--- Hasil Analisis per Kategori ---")
insight_df = df.groupby('Category').agg(
    Sales=('Sales', 'sum'),
    Profit=('Profit', 'sum'),
    Avg_Discount=('Discount', 'mean')
)
print(insight_df)

#Tambah kolom Profit Margin
insight_df['Profit_Margin'] = insight_df['Profit'] / insight_df['Sales']
insight_df = insight_df.reset_index()  # biar 'Category' jadi kolom biasa, bukan index

#Bikin tabel tambahan (sub-category)
subcat_df = df.groupby(['Category', 'Sub-Category']).agg(
    Sales=('Sales', 'sum'),
    Profit=('Profit', 'sum'),
    Avg_Discount=('Discount', 'mean')
).reset_index()
subcat_df['Profit_Margin'] = subcat_df['Profit'] / subcat_df['Sales']

#Export ke Excel
with pd.ExcelWriter('output_for_dashboard.xlsx') as writer:
    insight_df.to_excel(writer, sheet_name='By Category', index=False)
    subcat_df.to_excel(writer, sheet_name='By Sub-Category', index=False)