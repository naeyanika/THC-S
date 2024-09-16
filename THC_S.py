import streamlit as st
import pandas as pd
import os
import io


st.title('Aplikasi Pengolahan THC Simpanan')
st.markdown("""
## File yang dibutuhkan
1. **THC.xlsx**
   - Buat file baru dengan nama **THC.xlsx** lalu isi dengan data yang ada di "Format Data THC Gabungan Pivot.xlsb" di sheet atau lembar "Pivot".
   - Data yang diambil hanya dari **ID s.d Cr Total**.
   - Untuk menyamakan header Excel, gunakan format di bawah ini (tanda koma tidak perlu diikuti):
   [ ID | Dummy | NAMA | CENTER | KEL | HARI | JAM | STAF | TRANS. DATE | Db Qurban | Cr Qurban | Db Khusus | Cr Khusus | Db Sihara | Cr Sihara | Db Pensiun | Cr Pensiun | Db Pokok | Cr Pokok | Db SIPADAN | Cr SIPADAN | Db Sukarela | Cr Sukarela | Db Wajib | Cr Wajib | Db Total | Cr Total ].
   - File ini digunakan untuk melihat total transaksi, total nominal menabung, jumlah rata-rata anggota menabung, dan kebiasaan menabung anggota.

2. **DbSimpanan.xlsx**
   - Nama lembar atau sheet diganti menjadi **"Lembar1"**.
   - Kolom **CENTER** dan **KEL** (gunakan **text-to-column**) -delimited → tab → general. Yang tadinya **Center 001 Kelompok 01** menjadi **Center 1 Kelompok 1**.
   - Nama sheet pada **DbSimpanan.xlsx** tetap **"IA_SimpananDB"**, jangan diubah lagi.
   - File ini digunakan untuk melihat total saldo simpanan anggota.

3. **TAK.xlsx**
   - Data yang diambil dari **TAK.xlsx** yang sudah diolah di poin **"2.TAK"**.
   - File ini digunakan untuk melihat apakah anggota tersebut masih aktif atau sudah keluar.

4. **SimpananSiharaRpt.xlsx**
   - Data ini diambil dari modul detail simpanan pada **Operasional** → **Laporan Operasional Cabang** → **Detail Simpanan**.
   - Pilih tanggal sesuai periode audit.
   - Hapus bagian header sebelum di upload.
   - Ganti Nama Header **Member's Status** menjadi **Member Status**

    Jadi kita membutuhkan 4 file yaitu : :blue-background[THC.xlsx], :blue-background[DbSimpanan.xlsx], :blue-background[TAK.xlsx], dan :blue-background[SimpananSiharaRpt.xlsx]. Pastikan nama file sudah sesuai dengan instruksi.              
""")

uploaded_files = st.file_uploader("Unggah file Excel", accept_multiple_files=True, type=["xlsx"])
dfs = {}

required_files = ['DbSimpanan.xlsx', 'THC.xlsx', 'TAK.xlsx', 'SimpananSiharaRpt.xlsx']

if uploaded_files:
    for file in uploaded_files:
        if file.name == 'DbSimpanan.xlsx':
            df = pd.read_excel(file, engine='openpyxl', skiprows=1)
        else:
            df = pd.read_excel(file, engine='openpyxl')
        df.columns = df.columns.str.strip()
        dfs[file.name] = df

    if all(file in dfs for file in required_files):
        df_db = dfs['DbSimpanan.xlsx']
        df = dfs['THC.xlsx']
        df_tak = dfs['TAK.xlsx']
        df_shr = dfs['SimpananSiharaRpt.xlsx']

#-----------------------------Sesi Filter
    #Filter Db Simpanan
    df_simpanan = df_db[(df_db['Sts. Anggota'] == 'AKTIF') &
                        (df_db['Sts. Simpanan'] == 'AKTIF')]
    # Filter sihara
    df_sihara = df_simpanan[(df_simpanan['Product Name'] == 'Simpanan Hari Raya')]

    # Filter sukarela
    df_sukarela_2 = df_simpanan[(df_simpanan['Product Name'] == 'Simpanan Sukarela')]

    # Filter pensiun
    df_pensiun_2 = df_simpanan[(df_simpanan['Product Name'] == 'Simpanan Pensiun')]

#----------------------------Sesi Pivot
    # Pivot table simpanan
    df = df.rename(columns=lambda x: x.strip())
    pivot_table_simpanan = pd.pivot_table(df,
                                index=['ID', 'NAMA', 'CENTER', 'KEL'],
                                values=['Db Sihara', 'Cr Sihara', 'Db Pensiun', 'Cr Pensiun', 'Db Sukarela', 'Cr Sukarela', 'Db Wajib', 'Cr Wajib', 'Db Total', 'Cr Total'],
                                aggfunc='sum')
    
    desired_order = [
            'ID', 'NAMA', 'CENTER', 'KEL', 'Db Sihara','Cr Sihara','Db Pensiun','Cr Pensiun','Db Sukarela','Cr Sukarela', 'Db Wajib', 'Cr Wajib', 'Db Total','Cr Total'
            ]
    desired_order = [col for col in desired_order if col in pivot_table_simpanan.columns]
    pivot_table_simpanan = pivot_table_simpanan[desired_order]

    pivot_table_simpanan.to_excel('THC S.xlsx')

    # Membaca df1 sebagai thc simpanan
    df1 = pd.read_excel('THC S.xlsx')

#-------------Arsip Sesi Sihara 
    selected_columns = ['ID', 'NAMA', 'CENTER', 'KEL', 'Db Sihara', 'Cr Sihara']
    df1_selected_1 = df1[selected_columns]
    
    df['Modus_Sihara'] = df.groupby(['ID', 'NAMA'])['Db Sihara'].transform(lambda x: x.mode()[0])
    
    df1_selected = df.loc[:, ['ID', 'NAMA', 'Modus_Sihara']]
    df1_selected.drop_duplicates(subset=['ID', 'NAMA'], keep='first', inplace=True)
    df1_selected['Nilai_Modus'] = df1_selected['ID'].map(df1_selected.set_index('ID')['Modus_Sihara'])
    df1_selected_1['Sisa'] = df1_selected_1['Db Sihara'] - df1_selected_1['Cr Sihara']

    df.rename(columns=lambda x: x.strip(), inplace=True)
    df.rename(columns={'TRANS. DATE': 'TRANS_DATE'}, inplace=True)

    df_baru_2 = df[['ID', 'TRANS_DATE']].groupby('ID').nunique().reset_index().rename(columns={'TRANS_DATE':'Total Transaksi'})
    df_baru_3 = pd.merge(df[['ID', 'NAMA', 'CENTER', 'KEL']], df_baru_2, on='ID')
    df_baru_3.drop_duplicates(subset=['ID', 'NAMA'], keep='first', inplace=True)

    ################################
    df_temp = pd.merge(df1_selected_1, df1_selected, on=['ID', 'NAMA'], how='left')
    df2 = pd.merge(df_temp, df_baru_3, on=['ID', 'NAMA'], how='left')

    df_sample = df[(df['Db Sihara'] == df['Modus_Sihara'])].groupby('ID').size().reset_index()
    df_sample.rename(columns={0: 'Transaksi Sesuai'}, inplace=True)
    df_final = pd.merge(df2, df_sample, on='ID', how='left')
    ####################################
    df_sample_2 = df[(df['Db Sihara'] == 0)].groupby('ID').size().reset_index()
    df_sample_2.rename(columns={0: 'Transaksi Nol'}, inplace=True)
    df_final_2 = pd.merge(df_final, df_sample_2, on='ID', how='left')
    df_final_2 = df_final_2.drop(columns=['CENTER_y', 'KEL_y'])
    df_final_2['Transaksi Nol'] = df_final_2['Transaksi Nol'].fillna(0)
    df_final_2['Transaksi Nol'] = df_final_2['Transaksi Nol'].astype(int)
    ####################################
    df_final_5 = df[(df['Db Sihara'] != 0) & (df['Db Sihara'] != df['Modus_Sihara'])].groupby('ID').size().reset_index()
    df_final_5.rename(columns={0: 'Transaksi Tidak Sesuai'}, inplace=True)
    df_final_5 = pd.merge(df_final_2, df_final_5, on='ID', how='left')
    df_final_5['Transaksi Tidak Sesuai'].fillna(0, inplace=True)
    df_final_5['Transaksi Tidak Sesuai'] = df_final_5['Transaksi Tidak Sesuai'].astype(int)
    df_final_5['Sisa'].fillna(0, inplace=True)
    df_final_5['Sisa'] = df_final_5['Sisa'].astype(int)
    

    df_final_5= df_final_5.rename(columns={
    'ID':'ID Anggota',
    'NAMA':'Nama',
    'CENTER_x':'Center',
    'KEL_x':'Kelompok',
    'Db_Sihara': 'Db Sihara',
    'Cr_Sihara': 'Cr Sihara',
    'Nilai_Modus': 'Nilai Modus',
    'Modus_Sihara':'Modus Sihara'})

    ordered_columns = [
    'ID Anggota', 'Nama', 'Center', 'Kelompok', 'Modus Sihara',
    'Nilai Modus', 'Sisa', 'Total Transaksi', 'Transaksi Sesuai',
    'Transaksi Nol', 'Transaksi Tidak Sesuai'
]
    df_final_5 = df_final_5.reindex(columns=ordered_columns)

    merged_df = df_final_5.merge(df_sihara[['Client ID', 'Saldo']], left_on='ID Anggota', right_on='Client ID', how='left')
    merged_df.rename(columns={'Saldo': 'Saldo Sebelumnya'}, inplace=True)
    merged_df.drop(columns=['Client ID'], inplace=True)
    merged_df['Saldo Sebelumnya'].fillna(0, inplace=True)
# Tambah selisih saldo di sihara
    merged_df2 = merged_df.merge(df1[['ID', 'Db Sihara', 'Cr Sihara']], left_on='ID Anggota', right_on='ID', how='left')
    merged_df2['Saldo Akhir'] = merged_df2['Saldo Sebelumnya'] + merged_df2['Db Sihara'] - merged_df2['Cr Sihara']
    merged_df2.drop(columns=['ID', 'Db Sihara', 'Cr Sihara'], inplace=True)
    merged_df2.rename(columns={'Sisa': 'Selisih Transaksi'}, inplace=True)

    desired_order = [
        'ID Anggota','Nama','Center','Kelompok','Saldo Sebelumnya','Modus Sihara','Nilai Modus','Selisih Transaksi','Saldo Akhir','Total Transaksi','Transaksi Sesuai','Transaksi Nol','Transaksi Tidak Sesuai'
    ]
    for col in desired_order:
        if col not in merged_df2.columns:
            merged_df2[col] = 0

    final_sihara = merged_df2[desired_order]

#-------------Sihara Session
#Pilih data yang diperlukan untuk kolom THC S
    selected_columns = ['ID', 'NAMA', 'CENTER', 'KEL', 'Db Sihara', 'Cr Sihara']
    df1_sihara = df1[selected_columns]

#Pilih data yang di perlukan untuk kolom SiharaRpt
    selected_columns = ['Center', 'Group', 'Client ID', 'Name', 'Deposit Standard', 'Member Status']
    df1_shr = df_shr[selected_columns]
    
    #Ubah Uama Kolom
    rename_dict = {
    'Client ID': 'ID',
    'Name': 'NAMA',
    'Center': 'CENTER',
    'Group': 'KEL',
    'Deposit Standard': 'PAKET',
    'Member Status': 'STATUS'
    }
    df1_shr =df1_shr.rename(columns=rename_dict)

    # Ubah Urutan Kolom
    desired_order = ['ID', 'NAMA', 'CENTER', 'KEL', 'PAKET', 'STATUS']
    df1_shr = df1_shr[desired_order]

    #VLOOKUP df1_sihara dan df1_shr
    merge_column = 'ID'
    df_sihara_merge = pd.merge(df1_sihara, df1_shr, on=merge_column, suffixes=('_df1_sihara','_df1_shr'))

    # Ubah urutan kolom
    desired_order_merge = [
        'ID','NAMA_df1_sihara','CENTER_df1_sihara','KEL_df1_sihara','PAKET','Db Sihara','Cr Sihara','STATUS'
    ]
    df_sihara_merge = df_sihara_merge[desired_order_merge]

    rename_dict = {
        'NAMA_df1_sihara': 'NAMA',
        'CENTER_df1_sihara': 'CENTER',
        'KEL_df1_sihara': 'KEL'
    }
    df_sihara_merge = df_sihara_merge.merge(column=rename_dict)


    st.write("Sihara:")
    st.write(df_sihara_merge)

#-------------Pensiun Session
    # Baca data pensiun dan hanya memilih beberapa kolom
    df_pensiun = pd.read_excel('THC S.xlsx')
    selected_columns = ['ID', 'NAMA', 'CENTER', 'KEL', 'Db Pensiun', 'Cr Pensiun']
    df1_pensiun = df_pensiun[selected_columns]

    #Konversi tipe data ke string
    df1_pensiun['ID'] = df1_pensiun['ID'].astype(str)
    df1_pensiun['NAMA'] = df1_pensiun['NAMA'].astype(str)
    df1_pensiun['CENTER'] = df1_pensiun['CENTER'].astype(str)
    df1_pensiun['KEL'] = df1_pensiun['KEL'].astype(str)

    merged_df5 = df1_pensiun.merge(df_pensiun_2[['Client ID', 'Saldo']], left_on='ID', right_on='Client ID', how='left')
    
    # Ganti nama kolom
    merged_df5.rename(columns={
        'Saldo': 'Saldo Sebelumnya',
        'NAMA': 'Nama',
        'CENTER': 'Center',
        'KEL': 'Kelompok'
    }, inplace=True)
    
    # Hapus kolom ID di df_s dan df_db
    merged_df5.drop(columns=['Client ID'], inplace=True)
    
    # Jika ada data #N/A maka di replace dengan nol
    merged_df5['Saldo Sebelumnya'].fillna(0, inplace=True)
    
    # Selisih sisa saldo diambil dari Saldo Sebelumnya + Db Pensiun - Cr Pensiun
    merged_df5['Sisa'] = merged_df5['Saldo Sebelumnya'] + merged_df5['Db Pensiun'] - merged_df5['Cr Pensiun']
    
    # Cek data tersebut masih aktif atau sudah keluar
    merged_df5['Status'] = merged_df5['ID'].apply(lambda x: 'KELUAR' if x in df_tak['ID ANGGOTA'].values else 'AKTIF')

    # Anomali
    merged_df5['Anomali'] = merged_df5.apply(
    lambda row: 0 if row['Status'] == 'KELUAR' else (1 if row['Sisa'] < row['Saldo Sebelumnya'] else 0), axis=1
    )

    # Susun ulang dataframe
    desired_order = [
        'ID', 'Nama', 'Center', 'Kelompok', 'Status', 'Saldo Sebelumnya', 'Db Pensiun', 'Cr Pensiun', 'Sisa', 'Anomali'
    ]
    for col in desired_order:
        if col not in merged_df5.columns:
            merged_df5[col] = 0

    final_pensiun = merged_df5[desired_order]

    st.write("THC Pensiun:")
    st.write(final_pensiun)
#-------------Sukarela
    df_sukarela = pd.read_excel('THC S.xlsx')
    selected_columns = ['ID', 'NAMA', 'CENTER', 'KEL', 'Db Sukarela', 'Cr Sukarela']
    df1_sukarela = df_sukarela[selected_columns]
    
    df['Modus Sukarela'] = df.groupby(['ID', 'NAMA'])['Db Sukarela'].transform(lambda x: x.mode()[0])
    df_selected = df.loc[:, ['ID', 'NAMA', 'Modus Sukarela']]
    df_selected.drop_duplicates(subset=['ID', 'NAMA'], keep='first', inplace=True)
    df1_sukarela['Nilai Modus'] = df1_sukarela['ID'].map(df_selected.set_index('ID')['Modus Sukarela'])
    
    df_baru_2.rename(columns=lambda x: x.strip(), inplace=True)
    df_baru_2.rename(columns={'TRANS. DATE': 'TRANS_DATE'}, inplace=True)

    df_baru_2 = df[['ID', 'TRANS_DATE']].groupby('ID').nunique().reset_index().rename(columns={'TRANS_DATE':'Total Transaksi'})
    df_baru_2.head()


    df_baru_3 = pd.merge(df[['ID', 'NAMA', 'CENTER', 'KEL']], df_baru_2, on='ID')
    df_baru_3.drop_duplicates(subset=['ID', 'NAMA'], keep='first', inplace=True)

    df2 = df1_sukarela.merge(df_baru_3, on=['ID', 'NAMA'], how='left')
    df2.drop_duplicates(subset=['ID', 'NAMA', 'Db Sukarela', 'Cr Sukarela', 'Nilai Modus'], keep='first', inplace=True)
    df2_cleaned = df2.drop(['CENTER_y', 'KEL_y'], axis=1)
    df2_cleaned = df2_cleaned.rename(columns={'CENTER_x': 'CENTER', 'KEL_x': 'KEL',})
    
    df_sample_2 = df[(df['Db Sukarela'] != 0) + (df['Db Sukarela'] == df['Modus Sukarela'])].groupby('ID').size().reset_index()
    df_sample_2.rename(columns={0: 'Total Menabung > 0'}, inplace=True)
    df_final_3 = pd.merge(df2_cleaned, df_sample_2, on='ID', how='left')
    
    df_sample = df[(df['Db Sukarela'] != 0) & (df['Db Sukarela'] != df['Modus Sukarela'])].groupby('ID').size().reset_index()
    df_sample.rename(columns={0: 'Transaksi > 0 & ≠ Modus Sukarela'}, inplace=True)
    df_sample.head()

    df_final = pd.merge(df_final_3, df_sample, on='ID', how='left')

    df_final['Transaksi > 0 & ≠ Modus Sukarela'] = df_final['Transaksi > 0 & ≠ Modus Sukarela'].fillna(0)
    df_final['Transaksi > 0 & ≠ Modus Sukarela'] = df_final['Transaksi > 0 & ≠ Modus Sukarela'].astype(int)
    
    merged_df_final = df_final.merge(df_sukarela_2[['Client ID', 'Saldo']], left_on='ID', right_on='Client ID', how='left')
    merged_df_final.rename(columns={'Saldo': 'Saldo Sebelumnya'}, inplace=True)
    merged_df_final.drop(columns=['Client ID'], inplace=True)
    merged_df_final['Saldo Sebelumnya'].fillna(0, inplace=True)
    merged_df_final['Sisa Saldo'] = merged_df_final['Saldo Sebelumnya'] + merged_df_final['Db Sukarela'] - merged_df_final['Cr Sukarela']

    desired_order = ['ID','NAMA','CENTER','KEL','Saldo Sebelumnya','Db Sukarela','Cr Sukarela','Sisa Saldo','Nilai Modus','Total Transaksi','Total Menabung > 0','Transaksi > 0 & ≠ Modus Sukarela'
                    ]
    for col in desired_order:
        if col not in merged_df_final.columns:
            merged_df_final[col] = 0

    final_sukarela = merged_df_final[desired_order]


    st.write("THC Sukarela:")
    st.write(final_sukarela)


    def download_multiple_sheets():
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Menulis tiap dataframe ke sheet yang berbeda
            final_sihara.to_excel(writer, index=False, sheet_name='Sihara')
            final_pensiun.to_excel(writer, index=False, sheet_name='Pensiun')
            final_sukarela.to_excel(writer, index=False, sheet_name='Sukarela')
        buffer.seek(0)
        return buffer


    for name, df in {
        'Sihara.xlsx': final_sihara,
        'Pensiun.xlsx': final_pensiun,
        'Sukarela.xlsx': final_sukarela
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


        buffer_all = download_multiple_sheets()
        st.download_button(
        label="Unduh Semua Anomali Simpanan.xlsx",
        data=buffer_all.getvalue(),
        file_name="Anomali Simpanan.xlsx",
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )




    else:
        missing_files = [file for file in required_files if file not in dfs]
        st.error(f"File berikut tidak ditemukan: {', '.join(missing_files)}. Pastikan semua file yang diperlukan diunggah.")
else:
    st.warning("Silakan unggah file Excel yang diperlukan.")
