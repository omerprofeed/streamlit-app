import pandas as pd
import sqlite3

# Master data yükleme
master_data_path = 'C:\\Users\\Perakende\\Desktop\\data pivot 2.0\\master_data_eski.xlsx'
master_df = pd.read_excel(master_data_path, dtype={'Barcode': str})
master_df = master_df.rename(columns={'A': 'Barcode', 'B': 'Category'})  # Sütun isimlerini kontrol edin ve gerekirse düzeltin

# SQLite veritabanı bağlantısı
conn = sqlite3.connect('master_data.db')
cursor = conn.cursor()

# Master data tablo oluşturma
master_df.to_sql('master_data', conn, if_exists='replace', index=False)

conn.close()
print("Master data veritabanına yüklendi.")
