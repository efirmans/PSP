import streamlit as st
import pandas as pd
from openpyxl import Workbook
import re
from datetime import datetime
import altair as alt

st.set_page_config(
    page_title='Master Tagihan',
    layout="centered")

# st.write("### Data Preview")

st.markdown(
        """
        ### :blue[Upload file xls Master Tagihan dari Aplikasi PSP] 
          
    """
    ,)

uploaded_file = st.file_uploader("Upload an Excel file", type=["xls"])
if uploaded_file is not None:
    # Baca file excel
    try:
        df = pd.read_excel(uploaded_file)
        df['NIS'] = df['NIS'].astype(str)
        pattern = r"(\d{2}-\d{2}-\d{4})"
        df['Tanggal Pembayaran'] = df['Tanggal Pembayaran'].str.extract(pattern)
        
        df['Tanggal Pembayaran'] = pd.to_datetime(df['Tanggal Pembayaran'],format='%d-%m-%Y',errors='coerce')
        df['Tanggal Tagihan'] = pd.to_datetime(df['Tanggal Tagihan'],errors='coerce')
        df['Tanggal Jatuh Tempo'] = pd.to_datetime(df['Tanggal Jatuh Tempo'],errors='coerce')
        
        # bikin urut bulan
        pola_bulan = r"(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)" 
        df['Urutan'] = df['Tagihan'].str.extract(pola_bulan)
        month_map = {
            "Januari": "07", "Februari": "08", "Maret": "09", "April": "10",
            "Mei": "11", "Juni": "12", "Juli": "01", "Agustus": "02",
            "September": "03", "Oktober": "04", "November": "05", "Desember": "06"
        }
        month_map = {key.lower(): value for key, value in month_map.items()}
        df['Urutan'] = df['Urutan'].str.lower().map(month_map) 

        #bikin kategori
        kategori = {
    "BIAYA PANGKAL" : '01',
    "SPP" : '02',
    "BIAYA KEGIATAN SATU TAHUN" :"03",
    "JEMPUTAN" : '04'  ,
    "INFAQ": '05',
    "TUNGGAKAN ALUMNI" : '06',
    "PEMBELIAN MINIMART": '07',
    "BIAYA PENDIDIKAN":'08'
}
        
        pola_kategori = r"(SPP|BIAYA KEGIATAN SATU TAHUN|INFAQ|BIAYA PANGKAL|JEMPUTAN|PEMBELIAN MINIMART|TUNGGAKAN ALUMNI |BIAYA PENDIDIKAN)"
        df['Kategori'] = df['Tagihan'].str.extract(pola_kategori)
        df['Kategori'] = df['Kategori'].map(kategori) 
        reverse_kategori = {v: k for k, v in kategori.items()}

        #bikin unit
        pola_unit = r"(?i)(SDIT|TKAE|TKAS|PKBM)"
        df['unit'] = df['Tags'].str.extract(pola_unit)
        df['unit'] = df['unit'].str.upper()
        # Tahun Ajaran
        def convert_date_range(date_value):
            month = date_value.month
            year = date_value.year
            
            if 1 <= month <= 6:
                return f"{year-1}-{year}"
            elif 7 <= month <= 12:
                return f"{year}-{year+1}"
            
            return None  # Handle unexpected cases

        df['tahun'] = df["Tanggal Tagihan"].apply(convert_date_range)
        df['kelas'] = df['Tags'].str.extract(r'(\d[a-zA-Z])')
        df['kelas'] = df['kelas'].str.upper()        
        # multi select
        unitPendidikan = df['unit'].drop_duplicates()
        pilihUnit = st.sidebar.multiselect ("Unit",unitPendidikan,key='selected_options',placeholder='Pilih unit',default=unitPendidikan)
        
       

        
        
        #penjabaran
        nilaiJabarkan = ['Summary','Detail Kategori Tagihan','Tunggakan','Rekap Pembayaran Siswa']
        jabarkan = st.sidebar.radio(label='Menu',options=nilaiJabarkan,key='jabar')
        # jabarkan =  st.sidebar.checkbox("Jabarkan kategori tagihan")
        if jabarkan == 'Summary':
            tagihan = df[df['unit'].isin(pilihUnit)]['Tagihan'].count()
            terbayar = df[df['unit'].isin(pilihUnit)]['Terbayarkan'].sum()
            kekurangan = df[df['unit'].isin(pilihUnit)]['Kekurangan'].sum()
            
            filtered_df = df[df['unit'].isin(pilihUnit)]
            if filtered_df.empty or filtered_df['Tanggal Jatuh Tempo'].isna().all():
                awal = "Data tidak tersedia"
                akhir = "Data tidak tersedia"
                last_payment = "Data tidak tersedia"
            else:
                awal = filtered_df['Tanggal Jatuh Tempo'].min().strftime('%d-%m-%Y')
                akhir = filtered_df['Tanggal Jatuh Tempo'].max().strftime('%d-%m-%Y')
                last_payment = filtered_df['Tanggal Pembayaran'].max().strftime('%d-%m-%Y') 
            
            st.subheader(':grey[RINGKASAN TAGIHAN]',divider=True)
            st.text(
                f'''

    Jumlah Tagihan: {tagihan}  
    Terbayar: {terbayar:,d}
    Belum terbayar: {kekurangan:,d}
    Tgl awal Tagihan : {awal}
    Tgl akhir Jatuh tempo : {akhir}
    Transaksi pembayaran terakhir: {last_payment} 

    '''
                    )


            
            
            filtered_df['Tagihan']= filtered_df['Tagihan'].str.replace(r'(SPP BULAN|JEMPUTAN)','',regex=True)

            filtered_df["Kategori"] = filtered_df["Kategori"].map(reverse_kategori).str.capitalize()

            
            tagihan = filtered_df['Kategori'].drop_duplicates().to_list()
            
            pilihTagihan = st.radio(label='Kategori',options=tagihan,horizontal=True,)
            urutan = filtered_df[filtered_df['Kategori'] == pilihTagihan].sort_values(by='Urutan')
            
            if pilihTagihan == 'SPP' or pilihTagihan == 'Jemputan':

                df = urutan.groupby(['Tagihan', 'Urutan']).agg(
                    Terbayarkan=('Terbayarkan', 'sum'),
                    Kekurangan=('Kekurangan', 'sum')
                ).reset_index().sort_values(by='Urutan')  # Reset index agar 'Urutan' jadi kolom biasa
                
                
            else:
                df = urutan.groupby('Tagihan').agg(
                    Terbayarkan=('Terbayarkan', 'sum'),
                    Kekurangan=('Kekurangan', 'sum')
                ).reset_index()  # Reset index agar 'Urutan' jadi kolom biasa
                
            df_melted = df.melt(id_vars="Tagihan", value_vars=["Terbayarkan", "Kekurangan"],
                                var_name="Category", value_name="Value")
            # Create the line chart
            chart = alt.Chart(df_melted).mark_bar(opacity=0.7).encode(
                x=alt.X("Tagihan:O", title="Periode",sort=list(df["Tagihan"]),axis=alt.Axis(labelAngle=45)),
                y=alt.Y("Value:Q", title="Nominal" ),
                color="Category:N" ,
                tooltip=[
                    alt.Tooltip("Tagihan:O", title="Periode",),
                    alt.Tooltip("Category:N", title="Kategori"),
                    alt.Tooltip("Value:Q", title="Nominal", format=",.0f")] , # Add thousand separator
            ).properties(
                title="Pembayaran", 
                width=700,  # Lebar chart
                height=500)  # Tinggi chart).interactive()

            st.altair_chart(chart)

           

            
        elif jabarkan == 'Detail Kategori Tagihan':
            # df["Kategori"] = df["Kategori"].map(reverse_kategori)
            st.subheader(':green[KATEGORI TAGIHAN]',divider=True)
            #Agregat Tagihan
            df["Kategori"] = df["Kategori"].map(reverse_kategori)
            hasilFilter = df[df['unit'].isin(pilihUnit)][['Kategori','Terbayarkan','Kekurangan']]
            agg_data_namaTagihan = hasilFilter.groupby('Kategori').agg(
                    Terbayarkan=('Terbayarkan', 'sum'),
                    Kekurangan=('Kekurangan', 'sum')
        )
            st.dataframe(agg_data_namaTagihan)
            kategoriTagihan = st.sidebar.selectbox(label='Kategori tagihan',options=df['Kategori'].unique())
            
            if kategoriTagihan:
               
                if kategoriTagihan == 'SPP' or  kategoriTagihan == 'JEMPUTAN':
                    df = df.sort_values(by='Urutan')
                    filterKategori = df[(df['unit'].isin(pilihUnit)) & (df['Kategori'] == kategoriTagihan)][['Tagihan', 'Terbayarkan', 'Kekurangan', 'Urutan']]
                    agg_kategoriTagihan = filterKategori.groupby(['Tagihan', 'Urutan'],as_index=False).agg(
                        Terbayarkan=('Terbayarkan', 'sum'),
                        Kekurangan=('Kekurangan', 'sum')
                    )
                    # Urutkan berdasarkan 'Urutan'
                    agg_kategoriTagihan = agg_kategoriTagihan.sort_values(by='Urutan')

                    # Hapus kolom 'Urutan' sebelum menampilkan dataframe
                    agg_kategoriTagihan = agg_kategoriTagihan.drop(columns=['Urutan'])
                else:
                    
                    filterKategori = df[(df['unit'].isin(pilihUnit)) & (df['Kategori'] == kategoriTagihan)][['Tagihan', 'Terbayarkan', 'Kekurangan']]
                    agg_kategoriTagihan = filterKategori.groupby(['Tagihan'], as_index=False).agg(
                        Terbayarkan=('Terbayarkan', 'sum'),
                        Kekurangan=('Kekurangan', 'sum')
                    
                    )

                st.subheader(f':orange[RINCIAN TAGIHAN {kategoriTagihan}]',anchor='Kategori',divider='red')
                if agg_kategoriTagihan.empty:
                    pass
                else:
                    st.dataframe(agg_kategoriTagihan,hide_index=True)
                
                jabarRincian =  st.sidebar.checkbox("Jabarkan rincian tagihan")
                if jabarRincian:
                    rincianKategori = df[(df['unit'].isin(pilihUnit)) & (df['Kategori'] == kategoriTagihan)]['Tagihan'].drop_duplicates().to_list() 
                    pilihRincian =    st.sidebar.selectbox(label='rincian',options=rincianKategori)
                    if pilihRincian:
                        status = ['Lunas','Belum']
                        hasilStatus = st.sidebar.radio('status pembayaran',options=status,horizontal=True)

                        hasilRincian = df[(df['Tagihan'] == pilihRincian) & ((df['Lunas'] == hasilStatus) | (df['Belum'] == hasilStatus))]       [['Nama','NIS','Terbayarkan','Kekurangan']]
                        hasilRincian.reset_index(drop=True,inplace=True)
                        hasilRincian.index +=1
                        hasilRincian.index.name = 'No'
                        
                        
                        st.subheader(f':red[STATUS TAGIHAN {pilihRincian}]',anchor='rincian',divider=True)
                        st.text(f'status pembayaran {hasilStatus}')
                        st.dataframe(hasilRincian)
            
        #Tunggakan
        elif jabarkan == 'Tunggakan':
            st.subheader(':violet[TUNGGAKAN SISWA]',divider=True)
            df = df[(df['Belum'] == 'Belum') & (df['unit'].isin(pilihUnit))].sort_values(by=['Nama','Urutan'])
            agg_tunggakan =df.groupby(['Nama', 'NIS'], as_index=False).agg(
                Kekurangan=('Kekurangan', 'sum')
            )
            
            st.dataframe(agg_tunggakan.sort_values(by='Kekurangan', ascending=False),hide_index=True)
            nunggak = st.toggle('rincian tunggakan ')
            if nunggak:
                NIS = st.text_input(label='NIS', placeholder='input Nomor Induk Siswa (NIS)',key='nisTunggakan')
                if NIS:
                    nama = df[df['NIS'] == NIS]['Nama'].drop_duplicates().values[0]
                    df['kelas'] = df['kelas'].fillna(' ')
                    kelas = df[df['NIS'] == NIS]['kelas'].drop_duplicates().values[0]
                    tunggakan = df[df['NIS'] == NIS]['Kekurangan'].sum()
                    st.text(f'''
                            
Nama: {nama}
Kelas:{kelas}
jumlah tunggakan: {tunggakan:,d}      

'''


                    )
                    df = df[(df['Belum'] == 'Belum') & (df['unit'].isin(pilihUnit)) & (df['NIS'] == NIS)    ]
                    
                    df.reset_index(drop=True,inplace=True)
                    df.index +=1
                    df.index.name = 'No'
                    st.dataframe(df.iloc[:,[1,8,11,12]])            



        else:
            st.subheader(':blue[REKAP]',divider=True)
            geserNama = st.toggle('cari Nama/NIS siswa')
            
            if geserNama:
                cariNamaNis = st.text_input(label='Nama / NIS', placeholder='input Nama/NIS',key='carinamanis')
                    
                if not cariNamaNis:
                    pass
                else:
                    df = df[df['Nama'].str.contains(cariNamaNis, regex=False, case=False) | 
                    df['NIS'].str.contains(cariNamaNis, regex=False, case=False)]
                    hasil = df.drop_duplicates(subset=['Nama', 'NIS'])
                    st.dataframe(hasil[['Nama','NIS']],hide_index=True)

                
                
              
                   
            NIS = st.text_input(label='Rekap Pembayaran per siswa', placeholder='input Nomor Induk Siswa (NIS)',key='nisBayar')
            if not NIS:
                pass
            else:
                namaSiswa = df[df['NIS'] == NIS]['Nama'].drop_duplicates().values[0]
                st.write(f'Nama siswa: {namaSiswa}')
                df = df[df['NIS'] == NIS].sort_values(by=['Kategori','Urutan'])
                
                df.reset_index(drop=True,inplace=True)
                df.index +=1
                df.index.name = 'No'
                
                st.dataframe(df[['Tagihan','Total','Terbayarkan','Kekurangan']])   
                
    
    except Exception as e:
        st.error(f"Error reading the file: {e}")