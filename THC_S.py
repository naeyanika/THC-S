import streamlit as st
import pandas as pd
import os

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
6. 
""")

uploaded_files = st.file_uploader("Unggah file Excel", accept_multiple_files=True, type=["xlsx"])
dfs = {}  # Inisialisasi dfs di luar conditional
if uploaded_files:
    for file in uploaded_files:
        df = pd.read_excel(file, engine='openpyxl')
        dfs[file.name] = df

# Nama Dataframe
db_simpanan_path = 'DbSimpanan.xlsx'
thc_path = 'THC.xlsx'

if db_simpanan_path in dfs and thc_path in dfs:
    df_db = dfs[db_simpanan_path]
    df_thc = dfs[thc_path]
    # Lanjutkan dengan pengolahan data
else:
    st.error("Harap unggah file 'DbSimpanan.xlsx' dan 'THC.xlsx'")

if 'DbSimpanan.xlsx' in dfs and 'THC.xlsx' in dfs:
    df_db = dfs['DbSimpanan.xlsx']
    df_thc = dfs['THC.xlsx']

    st.write("THC:")
    st.write(df_thc)
    
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
    df_thc = df_thc.rename(columns=lambda x: x.strip())
    pivot_table_simpanan = pd.pivot_table(df_thc,
                                index=['ID', 'NAMA', 'CENTER', 'KEL'],
                                values=['Db Sihara', 'Cr Sihara', 'Db Pensiun', 'Cr Pensiun', 'Db Sukarela', 'Cr Sukarela', 'Db Wajib', 'Cr Wajib', 'Db Total', 'Cr Total'],
                                aggfunc='sum')
    
    desired_order = [
            'ID', 'NAMA', 'CENTER', 'KEL', 'Db Sihara','Cr Sihara','Db Pensiun','Cr Pensiun','Db Sukarela','Cr Sukarela', 'Db Total','Cr Total'
            ]
    desired_order = [col for col in desired_order if col in pivot_table_simpanan.columns]
    pivot_table_simpanan = pivot_table_simpanan[desired_order]
    
    st.write("THC Simpanan")
    st.write(pivot_table_simpanan)

    pivot_table_simpanan.to_excel('THC S.xlsx')

    # Membaca df1 sebagai thc simpanan
    df1 = pd.read_excel('THC S.xlsx')

    selected_columns = ['ID', 'NAMA', 'CENTER', 'KEL', 'Db Sihara', 'Cr Sihara']
    df1_selected_1 = df1[selected_columns]
    df1_selected_1['Modus_Sihara'] = df1_selected_1.groupby(['ID', 'NAMA'])['Db Sihara'].transform(lambda x: x.mode()[0])
    df1_selected = df1_selected_1.loc[:, ['ID', 'NAMA', 'Modus_Sihara']]
    df1_selected.drop_duplicates(subset=['ID', 'NAMA'], keep='first', inplace=True)

    # Handle duplicate IDs
    df1_selected['Nilai_Modus'] = df1_selected.groupby('ID')['Modus_Sihara'].transform('first')
    
    df1_selected_1['Sisa'] = df1_selected_1['Db Sihara'] - df1_selected_1['Cr Sihara']

    
    df1_selected_1.rename(columns=lambda x: x.strip(), inplace=True)
    df1_selected_1.rename(columns={'TRANS. DATE': 'TRANS_DATE'}, inplace=True)
    
    df_baru_2 = df1_selected_1[['ID', 'TRANS_DATE']].groupby('ID').nunique().reset_index().rename(columns={'TRANS_DATE': 'Total Transaksi'})
    
    df_baru_3 = pd.merge(df1_selected_1[['ID', 'NAMA', 'CENTER', 'KEL']], df_baru_2, on='ID')
    df_baru_3.drop_duplicates(subset=['ID', 'NAMA'], keep='first', inplace=True)
    df_temp = pd.merge(df1_selected_1, df1_selected, on=['ID', 'NAMA'], how='left')
    df2 = pd.merge(df_temp, df_baru_3, on=['ID', 'NAMA'], how='left')
    df_sample = df1_selected_1[(df1_selected_1['Db Sihara'] == df1_selected_1['Modus_Sihara'])].groupby('ID').size().reset_index()
    df_sample.rename(columns={0: 'Transaksi Sesuai'}, inplace=True)
    df_final = pd.merge(df2, df_sample, on='ID', how='left')
    df_sample_2 = df1_selected_1[(df1_selected_1['Db Sihara'] == 0)].groupby('ID').size().reset_index()
    df_sample_2.rename(columns={0: 'Transaksi Nol'}, inplace=True)
    df_final_2 = pd.merge(df_final, df_sample_2, on='ID', how='left')
    df_final_2 = df_final_2.drop(columns=['CENTER_y', 'KEL_y'])
    df_final_2['Transaksi Nol'] = df_final_2['Transaksi Nol'].fillna(0).astype(int)
    df_final_5 = df1_selected_1[(df1_selected_1['Db Sihara'] != 0) & (df1_selected_1['Db Sihara'] != df1_selected_1['Modus_Sihara'])].groupby('ID').size().reset_index()
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
            'Db Sihara': 'Db Sihara',
            'Cr Sihara': 'Cr Sihara',
            'Nilai Modus': 'Nilai Modus',
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

    # Pensiun
    df_pensiun = pd.read_excel('THC S.xlsx')

    selected_columns = ['ID', 'NAMA', 'CENTER', 'KEL', 'Db Pensiun', 'Cr Pensiun']
    df1_pensiun = df_pensiun[selected_columns]
    df1_pensiun['Sisa'] = df1_pensiun['Db Pensiun'] - df1_pensiun['Cr Pensiun']

    st.write("Pensiun:")
    st.write(df1_pensiun)   

    # Sukarela
    df_sukarela = pd.read_excel('THC S.xlsx')

    selected_columns = ['ID', 'NAMA', 'CENTER', 'KEL', 'Db Sukarela', 'Cr Sukarela']
    df1_sukarela = df_sukarela[selected_columns]
    df['Modus Sukarela'] = df.groupby(['ID', 'NAMA'])['Db Sukarela'].transform(lambda x: x.mode()[0])
    df_selected = df.loc[:, ['ID', 'NAMA', 'Modus Sukarela']]
    df_selected.drop_duplicates(subset=['ID', 'NAMA'], keep='first', inplace=True)
    
    # Handle duplicate IDs
    df1_sukarela['Nilai Modus'] = df1_sukarela['ID'].map(df_selected.set_index('ID')['Modus Sukarela'].to_dict())
    
    df2 = df1_sukarela.merge(df_baru_3, on=['ID', 'NAMA'], how='left')
    df2.drop_duplicates(subset=['ID', 'NAMA', 'Db Sukarela', 'Cr Sukarela', 'Nilai Modus'], keep='first', inplace=True)
    df2_cleaned = df2.drop(['CENTER_y', 'KEL_y'], axis=1)
    df2_cleaned = df2_cleaned.rename(columns={'CENTER_x': 'CENTER', 'KEL_x': 'KEL'})
    df_sample_2 = df[(df['Db Sukarela'] != 0) | (df['Db Sukarela'] == df['Modus Sukarela'])].groupby('ID').size().reset_index()
    df_sample_2.rename(columns={0: 'Total Menabung > 0'}, inplace=True)
    df_final_3 = pd.merge(df2_cleaned, df_sample_2, on='ID', how='left')
    df_sample = df[(df['Db Sukarela'] != 0) & (df['Db Sukarela'] != df['Modus Sukarela'])].groupby('ID').size().reset_index()
    df_sample.rename(columns={0: 'Transaksi > 0 & ≠ Modus Sukarela'}, inplace=True)
    df_final = pd.merge(df_final_3, df_sample, on='ID', how='left')
    df_final['Transaksi > 0 & ≠ Modus Sukarela'] = df_final['Transaksi > 0 & ≠ Modus Sukarela'].fillna(0).astype(int)
    
    st.write("Sukarela:")
    st.write(df_final)

    # Download links for all
    for name, df in {
        'Sihara.xlsx': df_final_5,
        'Sukarela.xlsx': df_final,
        'Pensiun.xlsx': df1_pensiun
    }.items():
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        buffer.seek(0)
        st.download_button(
            label=f"Unduh {name}",
            data=buffer.getvalue(),
            file_name=name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

else:
    st.error("File DbSimpanan.xlsx atau THC.xlsx tidak ditemukan. Pastikan file ada di lokasi yang benar.")
