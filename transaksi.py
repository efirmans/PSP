import streamlit as st
import pandas as pd
from openpyxl import Workbook
import re
from datetime import datetime
import altair as alt
from io import BytesIO
import numpy as np
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase.pdfmetrics import stringWidth
import matplotlib.pyplot as plt


uploaded_file = st.file_uploader("Upload file Excel dari Aplikasi PSP - Mutasi Saldo", type=["xlsx"])
if uploaded_file is not None:
    # Read the Excel file
    try:
        df = pd.read_excel(uploaded_file,thousands=",",decimal=',')
        df['Tanggal'] = pd.to_datetime (df['Tanggal'], format='%d-%m-%Y %H:%M')
        df['Tanggal'] = pd.to_datetime(df['Tanggal']).dt.date

        df['Nama COA'] = df['Kode COA'].str.extract(r",\s\n(.+)")
        df['Kode COA'] = df['Kode COA'].str.extract(r"(.+)\,")
        df['Nama User'] = df['Nama Akun'].str.extract(r"(.+)\,")
        df['Nama Akun'] = df['Nama Akun'].str.extract(r",\s\n(.+)")
        df['Deskripsi'] = df['Deskripsi'].str.extract(r"(.+)\,")
        kategori = r"(INFAQ|DOMPET PENDIDIKAN|SPP|BKS|SERAGAM|JEMPUTAN|BIAYA MASUK|TOP UP|SALDO|POOLING|PENAMPUNGAN|PEMBELIAN)"
        df['Kategori'] = df['Nama COA'].str.extract(kategori)
        df ['Debit'] = df ['Debet'] 
        pola_unit = r"(?i)(SDIT ANAK SHALIH|TKIT AISYAH|TKIT ANAK SHALIH|PKBM ANAK SHALIH)"
        df['Unit'] = df['Nama COA'].str.extract(pola_unit)
        inisial = {'SDIT ANAK SHALIH':'SDIT','TKIT AISYAH':'TKAE','TKIT ANAK SHALIH':'TKAS','PKBM ANAK SHALIH':'PKBM'}
        df['Unit']= df['Unit'].map(inisial)
        df.loc[df['Deskripsi'].str.contains('OVERFLOW', case=False, na=False), 'Nama Akun'] = 'YPIIAH'
        df =df [['Tanggal','Kode COA','Nama COA','Kategori','Unit','Nama User','Nama Akun',
                'Deskripsi','Debit','Kredit']]
        
        def summary():
            st.write('Summary')
                    # Summary
            awalTransaksi = df['Tanggal'].min().strftime('%d-%m-%Y')
            akhirTransaksi = df['Tanggal'].max().strftime('%d-%m-%Y')
            st.write(
                f'Transaksi dari tanggal :blue[{awalTransaksi}] s.d. :blue[{akhirTransaksi}]'
            )
            TransaksiUnit = df.groupby('Unit').agg(Jumlah = ('Kredit','sum')).reset_index()
            TransaksiUnit['Persentase'] = (TransaksiUnit['Jumlah'] / TransaksiUnit['Jumlah'].sum())*100
            # TransaksiUnit['Persentase'] = TransaksiUnit['Persentase'].apply(lambda x: f"{x:.2%}")
            transaksiUnit_diplay = TransaksiUnit[['Unit','Jumlah']]
            columns_to_format = ['Jumlah']  
            TransaksiUnit_Formatted = transaksiUnit_diplay.style.format({col: "{:,}" for col in columns_to_format})
            st.subheader(f':orange[Pendapatan per Unit]')
            st.dataframe(TransaksiUnit_Formatted,hide_index=True)
            st.subheader(f':green[Porsi Perolehan]')

            labels = TransaksiUnit['Unit']
            sizes = TransaksiUnit['Persentase']
            explode = (0, 0.1, 0, 0)  # only "explode" the 2nd slice

            labels = [
                f"{unit} - {persen:.1f}%" 
                for unit, persen in zip(TransaksiUnit['Unit'], TransaksiUnit['Persentase'])
            ]
            sizes = TransaksiUnit['Persentase']
            explode = (0, 0.2, 0, 0)  # Bisa disesuaikan

            fig1, ax1 = plt.subplots()

            # Pie chart tanpa label & autopct
            wedges, texts = ax1.pie(
                sizes,
                explode=explode,
                labels=None,
                autopct=None,
                shadow=False,
                startangle=90
            )

            # Tambahkan legend dengan label + persentase
            ax1.legend(
                wedges,
                labels,
                title="Unit",
                loc="center left",
                bbox_to_anchor=(1, 0, 0.5, 1),
                fontsize=10
            )

            ax1.axis('equal')  # Lingkaran sempurna
            st.pyplot(fig1)

       
            # st.bar_chart(TransaksiUnit,x= 'Unit',y= 'Persentase',color='Unit',horizontal=True,height=400,stack=True,)
                                    
        def AkunKumulatif():          
            # group berdasar kode dan nama akun
            st.subheader('Kumulatif per Akun')
            unit = ["SDIT", "TKAE", "TKAS", "PKBM"]
            pilihUnit = st.segmented_control(label="Unit", options=unit, selection_mode="multi")
            df['Tanggal'] = pd.to_datetime(df['Tanggal'])
            awalTransaksi = df['Tanggal'].min()
            akhirTransaksi = df['Tanggal'].max()
            Col1,Col2 = st.columns(2)
            with Col1:    
                awal = st.date_input("tanggal awal",key='group_awal',format="DD/MM/YYYY",value=awalTransaksi)
            with Col2:
                akhir = st.date_input("tanggal akhir",key='group_akhir',format="DD/MM/YYYY",value=akhirTransaksi)
            group_awal = pd.to_datetime(awal)
            group_akhir = pd.to_datetime(akhir)
            group_filter = df[(df['Unit'].isin(pilihUnit)) & (df['Tanggal'] >=group_awal) & (df['Tanggal'] <=group_akhir) ]         
            groupAkun = group_filter.groupby(['Kode COA','Nama COA','Unit']).agg(Kredit= ('Kredit',sum)
            ).reset_index().sort_values('Unit')
            if groupAkun.empty:
                pass
            else:
                df_groupAkun = groupAkun.style.format({col: "{:,}" for col in ['Kredit']})
                st.dataframe(df_groupAkun,hide_index=True)
                rincikan = st.toggle('rincian akun', key='rincikan')
                if rincikan:
                    NamaCOA = groupAkun['Nama COA'].drop_duplicates().to_list()
                    pilihCOA = st.multiselect('Nama Akun',options=NamaCOA)
                    filterCOA = group_filter[group_filter['Nama COA'].isin(pilihCOA)]
                    rincianAkun = filterCOA.groupby(['Deskripsi','Unit']).agg(Kredit = ('Kredit','sum')).reset_index().sort_values('Unit') 
                    rincianAkun = rincianAkun[rincianAkun['Kredit']>0] 
                    if rincianAkun.empty:
                        pass
                    else:
                        rincianAkun_formatted = rincianAkun.style.format({col: "{:,}" for col in ['Kredit']})    
                        st.dataframe(rincianAkun_formatted,hide_index=True)
                        rincianTransaksi = st.toggle('rincian transaksi')
                        if rincianTransaksi:
                            rincian_filterred = filterCOA[filterCOA['Deskripsi'].isin(rincianAkun['Deskripsi'])][['Tanggal','Nama Akun','Unit','Deskripsi','Kredit']]
                            if not rincian_filterred.empty:
                                rincian_filterred['Tanggal'] = rincian_filterred['Tanggal'].astype(str)
                                rincian_filterred['Tanggal'] = rincian_filterred['Tanggal'].apply(lambda x: pd.to_datetime(x).strftime('%d-%m-%Y'))
                                
                                #tambah total row
                                total_row = pd.DataFrame([{
                                'Tanggal': 'Total',
                                'Nama Akun': '',
                                'Unit': '',
                                'Deskripsi': '',
                                'Kredit': rincian_filterred['Kredit'].sum()
                            }])
                                rincian_filterred = pd.concat([rincian_filterred, total_row], ignore_index=True)
                                
                                rincian_akhir = rincian_filterred.style.format({col: "{:,}" for col in ['Kredit']})                               
                                st.dataframe(rincian_akhir,hide_index=True  )
                               
                                df_for_pdf = rincian_filterred.fillna("")

                                # Buat buffer untuk file PDF
                                buffer = BytesIO()
                                pdf = SimpleDocTemplate(buffer, pagesize=letter)

                                # thousand separator columns 
                                kolom_format = ["Debit","Kredit"]
                                table_data = [list(rincian_filterred.columns)] 
                                
                                # Buat data tabel dari dataframe
                                for _, row in rincian_filterred.iterrows():
                                    row_list = []
                                    for col in rincian_filterred.columns:
                                        if col in kolom_format:
                                            try:
                                                formatted = f"{row[col]:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
                                            except:
                                                formatted = row[col]
                                            row_list.append(formatted)
                                        else:
                                            row_list.append(row[col])
                                    table_data.append(row_list)

                                def get_column_widths(data, font_name="Helvetica", font_size=12, padding=12):
                                    col_widths = []
                                    num_cols = len(data[0])
                                    for col_idx in range(num_cols):
                                        max_width = 0
                                        for row in data:
                                            try:
                                                cell_text = str(row[col_idx])
                                                cell_width = stringWidth(cell_text, font_name, font_size)
                                                if cell_width > max_width:
                                                    max_width = cell_width
                                            except:
                                                pass
                                        col_widths.append(max_width + padding)  
                                    return col_widths

                                col_widths = get_column_widths(table_data, font_name="Helvetica", font_size=11)
                                
                                # Buat tabel dan style-nya
                                table = Table(table_data, colWidths=col_widths,repeatRows=1)

                                table_style = TableStyle([
                                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                    ('FONTSIZE', (0, 0), (-1, 0), 14),
                                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                    ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                                    ('FONTSIZE', (0, 1), (-1, -1), 11),
                                    ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
                                    ('GRID', (0, 0), (-1, -1), 0.1, colors.black),
                                ])

                                # rata kolom
                                kolom_kiri = ["Nama Akun","Deskripsi"]
                                kolom_kanan = ["Kredit"]
                                # Cari index kolom yang perlu left align
                                left_align_indexes = [i for i, col in enumerate(rincian_filterred.columns) if col in kolom_kiri]
                                right_align_indexes = [i for i, col in enumerate(rincian_filterred.columns) if col in kolom_kanan]

                                # Tambahkan style untuk kolom rata kiri
                                for col_idx in left_align_indexes:
                                    table_style.add('ALIGN', (col_idx, 0), (col_idx, -1), 'LEFT')  # dari baris 0 sampai terakhir

                                # Tambahkan style untuk kolom rata kanan
                                for col_idx in right_align_indexes:
                                    table_style.add('ALIGN', (col_idx, 0), (col_idx, -1), 'RIGHT')  # dari baris 0 sampai terakhir

                                awal_str = awal.strftime("%d %B %Y")
                                akhir_str = akhir.strftime("%d %B %Y")

                                def add_page_number_and_header(canvas, doc):
                                    # Header
                                    canvas.setFont('Helvetica-Bold', 14)
                                    canvas.drawString(2.5 * cm, 27.5 * cm, f"Pendapatan Periode {awal_str} - {akhir_str}")

                                    # Optional: garis bawah header
                                    canvas.line(2.5 * cm, 27.3 * cm, 19.5 * cm, 27.3 * cm)

                                    # Page number
                                    page_num = canvas.getPageNumber()
                                    canvas.setFont('Helvetica', 9)
                                    canvas.drawRightString(20.5 * cm, 1.5 * cm, f"{page_num}")

                                table.setStyle(table_style)

                                # Tambahkan ke elemen PDF dan build
                                elements = [table]
                                pdf.build(elements, onFirstPage=add_page_number_and_header, onLaterPages=add_page_number_and_header)

                                # Pindah ke awal stream untuk dibaca oleh Streamlit
                                buffer.seek(0)

                                # Tombol download di Streamlit
                                st.download_button(
                                    label="📄 Download PDF",
                                    data=buffer,
                                    file_name=f"Pendapatan {awal_str}-{akhir_str}.pdf",
                                    mime="application/pdf",
                                    key='pendapatan'
                                )

                else:
                    pass
            
        def transaksiHarian():
            
            # Saldo minimart
            st.subheader('Transaksi Minimart')
            df['Tanggal'] = pd.to_datetime(df['Tanggal'])
            awalTransaksi = df['Tanggal'].min()
            akhirTransaksi = df['Tanggal'].max()
            Col1,Col2 = st.columns(2)
            with Col1:    
                awal = st.date_input("tanggal awal",key='awal',format="DD/MM/YYYY",value=awalTransaksi)
            with Col2:
                akhir = st.date_input("tanggal akhir",key='akhir',format="DD/MM/YYYY",value=akhirTransaksi)

            df['Tanggal'] = pd.to_datetime(df['Tanggal']).dt.date
            awal = pd.to_datetime(awal)
            akhir = pd.to_datetime(akhir)

            transMinimart = df[(df['Nama User']=='MINIMART ANAK SHALIH') & (df['Tanggal'] >=awal) & (df['Tanggal'] <=akhir) ]
            transMinimart = transMinimart[['Tanggal','Nama Akun','Debit','Kredit']] 
                
            # Menambahkan baris Total
            total_row = pd.DataFrame([{
                'Tanggal': 'Total',
                'Nama Akun': '',
                'Debit': transMinimart['Debit'].sum(),
                'Kredit': transMinimart['Kredit'].sum()
            }])
            transMinimart = pd.concat([transMinimart, total_row], ignore_index=True)

            if not transMinimart.empty:
                transMinimart['Tanggal'] = transMinimart['Tanggal'].astype(str)
                transMinimart['Tanggal'] = transMinimart['Tanggal'].apply(
                    lambda x: pd.to_datetime(x).strftime('%d-%m-%Y') if x != '' and x != 'nan' and x != 'NaT' and x != 'Total' else x
                )
                transMinimart_styled = transMinimart.style.format({col: "{:,.0f}" for col in ['Debit','Kredit']})
                
                st.dataframe(transMinimart_styled,hide_index=True)
                
                # Buat buffer untuk file PDF
                buffer = BytesIO()
                pdf = SimpleDocTemplate(buffer, pagesize=letter)

                # Kolom yang ingin diberi pemisah ribuan
                kolom_format = ["Debit","Kredit"]
                table_data = [list(transMinimart.columns)] 
                
                # Buat data tabel dari dataframe
                for _, row in transMinimart.iterrows():
                    row_list = []
                    for col in transMinimart.columns:
                        if col in kolom_format:
                            try:
                                formatted = f"{row[col]:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            except:
                                formatted = row[col]
                            row_list.append(formatted)
                        else:
                            row_list.append(row[col])
                    table_data.append(row_list)

                # Buat tabel dan style-nya
                table = Table(table_data,repeatRows=1)

                table_style = TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 14),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 12),
                    ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 0.1, colors.black),
                ])

                # Misalnya kolom yang ingin dirata kiri:
                kolom_kiri = ["Nama Akun"]
                kolom_kanan = ["Debit","Kredit"]
                # Cari index kolom yang perlu left align
                left_align_indexes = [i for i, col in enumerate(transMinimart.columns) if col in kolom_kiri]
                right_align_indexes = [i for i, col in enumerate(transMinimart.columns) if col in kolom_kanan]

                # Tambahkan style untuk kolom rata kiri
                for col_idx in left_align_indexes:
                    table_style.add('ALIGN', (col_idx, 0), (col_idx, -1), 'LEFT')  # dari baris 0 sampai terakhir

                # Tambahkan style untuk kolom rata kanan
                for col_idx in right_align_indexes:
                    table_style.add('ALIGN', (col_idx, 0), (col_idx, -1), 'RIGHT')  # dari baris 0 sampai terakhir

                awal_str = awal.strftime("%d %B %Y")
                akhir_str = akhir.strftime("%d %B %Y")

                def add_page_number_and_header(canvas, doc):
                    # Header
                    canvas.setFont('Helvetica-Bold', 14)
                    canvas.drawString(2.5 * cm, 27.5 * cm, f"Transaksi Minimart Periode {awal_str} - {akhir_str}")

                    # Optional: garis bawah header
                    canvas.line(2.5 * cm, 27.3 * cm, 19.5 * cm, 27.3 * cm)

                    # Page number
                    page_num = canvas.getPageNumber()
                    canvas.setFont('Helvetica', 9)
                    canvas.drawRightString(20.5 * cm, 1.5 * cm, f"{page_num}")

                table.setStyle(table_style)

                # Tambahkan ke elemen PDF dan build
                elements = [table]
                pdf.build(elements, onFirstPage=add_page_number_and_header, onLaterPages=add_page_number_and_header)

                # Pindah ke awal stream untuk dibaca oleh Streamlit
                buffer.seek(0)

                # Tombol download di Streamlit
                st.download_button(
                    label="📄 Download PDF",
                    data=buffer,
                    file_name=f"Transaksi minimart {awal_str}-{akhir_str}.pdf",
                    mime="application/pdf",
                    key='minimart'
                )

        pg = st.navigation(pages=[
            st.Page(summary, title="Summary"),
            st.Page(AkunKumulatif, title="Akun"),
            st.Page(transaksiHarian, title="Transaksi minimart")
        ],)
        pg.run()
              
    except Exception as e:
        st.error(f"Tidak bisa membaca file: {e}")