import streamlit as st
import pandas as pd
import plotly.express as px
import sqlite3
from datetime import datetime
import os

@st.cache
def load_master_data(db_path):
    conn = sqlite3.connect(db_path)
    master_df = pd.read_sql('SELECT * FROM master_data', conn)
    conn.close()
    return master_df

@st.cache
def load_sales_data(file):
    sales_df = pd.read_excel(
        file, 
        usecols=['MarketPlace', 'Order Date', 'Status', 'Barcode', 'Product', 'Quantity', 'Amount', 'Discount', 'Price', 'Vat Amount'],
        dtype={'Barcode': str, 'Quantity': float, 'Amount': float, 'Discount': float, 'Price': float, 'Vat Amount': float}
    )
    return sales_df

# SQLite veritabanı bağlantısı
db_path = os.path.join(os.path.dirname(__file__), 'master_data.db')
master_df = load_master_data(db_path)

# Streamlit başlığı
st.title('Sales Report Analysis')

# Excel dosyasını yükleme
sales_report_file = st.file_uploader("Upload Sales Report Excel File", type="xlsx")

if sales_report_file:
    sales_df = load_sales_data(sales_report_file)
    st.write("Sales Data Sample", sales_df.head())  # Verilerin doğru yüklendiğini doğrulamak için

    # Veritabanı ile birleştirme
    sales_df = pd.merge(sales_df, master_df[['Barcode', 'Category']], on='Barcode', how='left')
    st.write("Merged Data Sample", sales_df.head())  # Birleştirilmiş verilerin doğru olduğunu doğrulamak için

    # Tarihleri datetime formatına dönüştürme
    sales_df['Order Date'] = pd.to_datetime(sales_df['Order Date'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
    st.write("Data Types After Date Conversion", sales_df.dtypes)  # Veri tiplerini kontrol etmek için

    # Kullanıcıdan tarih aralığı ve pazaryeri seçimini alma
    st.write("## Select Date Range")
    start_date = st.date_input('Start date', datetime(2023, 1, 1))
    end_date = st.date_input('End date', datetime(2023, 12, 31))

    # Tarih aralığı filtrelemesi
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    filtered_df = sales_df[(sales_df['Order Date'] >= start_date) & (sales_df['Order Date'] <= end_date)]
    st.write("Filtered Data by Date", filtered_df.head())  # Tarih filtrelemesinin doğru olduğunu kontrol etmek için

    # Pazaryeri seçimi
    st.write("## Select Marketplaces")
    marketplaces = st.multiselect('Marketplaces', sales_df['MarketPlace'].unique())
    if marketplaces:
        filtered_df = filtered_df[filtered_df['MarketPlace'].isin(marketplaces)]
        st.write("Filtered Data by Marketplaces", filtered_df.head())  # Pazaryeri filtrelemesinin doğru olduğunu kontrol etmek için

    # 'CANCELLED' statüsündeki satırları filtrele
    filtered_df = filtered_df[filtered_df['Status'] != 'CANCELLED']
    st.write("Filtered Data After Removing Cancelled", filtered_df.head())  # 'CANCELLED' filtrelemesinin doğru olduğunu kontrol etmek için

    # Toplam gelir sütunu ekle
    filtered_df['Total Amount'] = filtered_df['Quantity'] * filtered_df['Price']
    st.write("Data with Total Amount", filtered_df.head())  # Toplam gelir hesaplamasının doğru olduğunu kontrol etmek için

    # Pivot tablo oluştur
    pivot_table = filtered_df.pivot_table(index=['MarketPlace', 'Barcode', 'Product', 'Price', 'Category'],
                                          values=['Quantity', 'Total Amount'],
                                          aggfunc={'Quantity': 'sum', 'Total Amount': 'sum'}, 
                                          margins=True, 
                                          margins_name='Total').reset_index()

    st.write("## Pivot Table")
    st.dataframe(pivot_table)

    # MarketPlace kıyaslama tablosunu oluştur ve göster
    marketplace_comparison = filtered_df.groupby('MarketPlace').agg({'Quantity': 'sum', 'Total Amount': 'sum'}).reset_index()
    st.write("## Marketplace Comparison")
    st.dataframe(marketplace_comparison)

    # En çok satan ilk 10 ürünü bul ve göster
    top_10_products = filtered_df.groupby(['Barcode', 'Product', 'Category']).agg({'Quantity': 'sum'}).reset_index().sort_values(by='Quantity', ascending=False).head(10)
    st.write("## Top 10 Products")
    st.dataframe(top_10_products)

    # Kategori bazlı satış adet ve ciro tablosu oluştur ve göster
    category_comparison = filtered_df.groupby('Category').agg({'Quantity': 'sum', 'Total Amount': 'sum'}).reset_index()
    st.write("## Category Comparison")
    st.dataframe(category_comparison)

    # Grafikleri oluşturma ve gösterme
    st.subheader('Top 10 Best Selling Products')
    fig1 = px.bar(top_10_products, x='Product', y='Quantity', color='Category', title='Top 10 Best Selling Products')
    st.plotly_chart(fig1)

    st.subheader('Total Quantity Sold by Marketplace')
    fig2 = px.bar(marketplace_comparison, x='MarketPlace', y='Quantity', title='Total Quantity Sold by Marketplace')
    st.plotly_chart(fig2)

    st.subheader('Total Revenue by Marketplace')
    fig3 = px.bar(marketplace_comparison, x='MarketPlace', y='Total Amount', title='Total Revenue by Marketplace')
    st.plotly_chart(fig3)

    st.subheader('Daily Sales Trend')
    daily_sales = filtered_df.set_index('Order Date').resample('D').sum().reset_index()
    fig4 = px.line(daily_sales, x='Order Date', y='Quantity', title='Daily Sales Trend')
    st.plotly_chart(fig4)

    # Excel dosyasına pivot tabloyu yazma
    output_file_path = st.text_input('Output File Path', 'pivot_table.xlsx')
    if st.button('Generate Excel Report'):
        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            pivot_table.to_excel(writer, sheet_name='Pivot Table', index=False)
            marketplace_comparison.to_excel(writer, sheet_name='Marketplace Comparison', index=False)
            top_10_products.to_excel(writer, sheet_name='Top 10 Products', index=False)
            category_comparison.to_excel(writer, sheet_name='Category Comparison', index=False)
        st.success(f'Report saved to {output_file_path}')
