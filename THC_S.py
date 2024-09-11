import streamlit as st
import pandas as pd
import numpy as np
import io


st.title('Aplikasi Pengolahan THC Simpanan')
st.markdown("""
## Catatan:
1. Buat file baru dengan nama THC.xlsx lalu isi dengan data yang ada di "Format Data THC Gabungan.xlsb" di sheet atau lembar "Hasil Pivot 1"
2. Untuk data yang diambil hanya dari ID s.d Cr Total (simpanan)
3. Nama Lembar atau Sheet ganti jadi "Lembar1"
4. Untuk kolom CENTER dan KEL (text-to-coloumn) 
-delimited, tab, general. Yang tadinya Center 001 Kel 01 menjadi Center 1 Kel 1.
5. Untuk menyamakan Header excel gunakan seperti format dibawah ini (Koma nya jangan diikuti) 
ID, Dummy, NAMA, CENTER, KEL, HARI, JAM, STAF, TRANS. DATE, Db Qurban, Cr Qurban, Db Khusus, Cr Khusus, Db Sihara, Cr Sihara, Db Pensiun, Cr Pensiun, Db Pokok, Cr Pokok, Db SIPADAN, Cr SIPADAN, Db Sukarela, Cr Sukarela, Db Wajib, Cr Wajib, Db Total, Cr Total.
6. Untuk DbSimpanan, tidak udah di edit lagi bagian isi nya karena otomatis memotong header nya, jadi hanya di ubah namanya menjadi DbSimpanan.xlsx 
""")

uploaded_files = st.file_uploader("Unggah file Excel", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
    dfs = {}
    for file in uploaded_files:
        df = pd.read_excel(file, engine='openpyxl')
        dfs[file.name] = df

# Nama Dataframe
df_db = pd.read_excel('DbSimpanan.xlsx')
df_thc = pd.read_excel('THC.xlsx')

# Db Simpanan
df_simpanan = df_db[(df_db['Sts. Anggota'] == 'AKTIF') &
                    (df_db['Sts. Simpanan'] == 'AKTIF')]

# Filter sihara
df_sihara = df_simpanan[(df_simpanan['Product Name'] == 'Simpanan Hari Raya')]
st.write("Sihara:")
st.write(df_sihara)
# Filter sukarela
df_sukarela = df_simpanan[(df_simpanan['Product Name'] == 'Simpanan Sukarela')]
st.write("Sukarela:")
st.write(df_sukarela)
# Filter hariraya
df_pensiun = df_simpanan[(df_simpanan['Product Name'] == 'Simpanan Pensiun')]
st.write("Pensiun:")
st.write(df_pensiun)

# Pivot table simpanan
pivot_table_simpanan = pd.pivot_table(df_thc,
                            index=['ID', 'NAMA', 'CENTER', 'KEL'],
                            values=['Db Sihara', 'Cr Sihara', 'Db Pensiun', 'Cr Pensiun', 'Db Sukarela', 'Cr Sukarela', 'Db Wajib', 'Cr Wajib', 'Db Total', 'Cr Total'],
                            aggfunc='sum')
pivot_table_simpanan.to_excel('THC S.xlsx')

# Membaca df1 sebagai thc simpanan
df1 = pd.read_excel('THC S.xlsx')

selected_columns = ['ID', 'NAMA', 'CENTER', 'KEL', 'Db Sihara', 'Cr Sihara']
df1_selected_1 = df1[selected_columns]
df['Modus_Sihara'] = df.groupby(['ID', 'NAMA'])['Db Sihara'].transform(lambda x: x.mode()[0])
df1_selected = df.loc[:, ['ID', 'NAMA', 'Modus_Sihara']]
df1_selected.drop_duplicates(subset=['ID', 'NAMA'], keep='first', inplace=True)

# Handle duplicate IDs
df1_selected['Nilai_Modus'] = df1_selected.groupby('ID')['Modus_Sihara'].transform('first')
    
df1_selected_1['Sisa'] = df1_selected_1['Db Sihara'] - df1_selected_1['Cr Sihara']
df.rename(columns=lambda x: x.strip(), inplace=True)
df.rename(columns={'TRANS. DATE': 'TRANS_DATE'}, inplace=True)
df_baru_2 = df[['ID', 'TRANS_DATE']].groupby('ID').nunique().reset_index().rename(columns={'TRANS_DATE': 'Total Transaksi'})
df_baru_3 = pd.merge(df[['ID', 'NAMA', 'CENTER', 'KEL']], df_baru_2, on='ID')
df_baru_3.drop_duplicates(subset=['ID', 'NAMA'], keep='first', inplace=True)
df_temp = pd.merge(df1_selected_1, df1_selected, on=['ID', 'NAMA'], how='left')
df2 = pd.merge(df_temp, df_baru_3, on=['ID', 'NAMA'], how='left')
df_sample = df[(df['Db Sihara'] == df['Modus_Sihara'])].groupby('ID').size().reset_index()
df_sample.rename(columns={0: 'Transaksi Sesuai'}, inplace=True)
df_final = pd.merge(df2, df_sample, on='ID', how='left')
df_sample_2 = df[(df['Db Sihara'] == 0)].groupby('ID').size().reset_index()
df_sample_2.rename(columns={0: 'Transaksi Nol'}, inplace=True)
df_final_2 = pd.merge(df_final, df_sample_2, on='ID', how='left')
df_final_2 = df_final_2.drop(columns=['CENTER_y', 'KEL_y'])
df_final_2['Transaksi Nol'] = df_final_2['Transaksi Nol'].fillna(0).astype(int)
df_final_5 = df[(df['Db Sihara'] != 0) & (df['Db Sihara'] != df['Modus_Sihara'])].groupby('ID').size().reset_index()
df_final_5.rename(columns={0: 'Transaksi Tidak Sesuai'}, inplace=True)
df_final_5 = pd.merge(df_final_2, df_final_5, on='ID', how='left')
df_final_5['Transaksi Tidak Sesuai'].fillna(0, inplace=True)
df_final_5['Transaksi Tidak Sesuai'] = df_final_5['Transaksi Tidak Sesuai'].astype(int)
df_final_5['Sisa'].fillna(0, inplace=True)
df_final_5['Sisa'] = df_final_5['Sisa'].astype(int)
df_final_5 = df_final_5.rename(columns={
        'ID': 'ID Anggota',
        'NAMA': 'Nama',
        'CENTER_x': 'Center',
        'KEL_x': 'Kelompok',
        'Db_Sihara': 'Db Sihara',
        'Cr_Sihara': 'Cr Sihara',
        'Nilai_Modus': 'Nilai Modus',
        'Modus_Sihara': 'Modus Sihara'
    })
ordered_columns = [
        'ID Anggota', 'Nama', 'Center', 'Kelompok', 'Modus Sihara',
        'Nilai Modus', 'Sisa', 'Total Transaksi', 'Transaksi Sesuai',
        'Transaksi Nol', 'Transaksi Tidak Sesuai'
    ]

df_final_5 = df_final_5.reindex(columns=ordered_columns)

st.write("Sihara:")
st.write(df_final_5)
