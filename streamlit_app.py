import streamlit as st
from streamlit_option_menu import option_menu
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import scipy.cluster.hierarchy as sch
from io import BytesIO
from sklearn.cluster import AgglomerativeClustering
from sklearn.metrics import davies_bouldin_score
from num2words import num2words
from functools import reduce

st.set_page_config(page_title="Beasiswa SMANSASERA")

class MainClass():

    def __init__(self):
        #inisiasi objek
        self.data = Data()
        self.preprocessing = Preprocessing()
        self.dbi = Dbi()
        self.clustering = Clustering()

    # Fungsi judul halaman
    def judul_halaman(self, header, subheader):
        nama_app = "Aplikasi Rekomendasi Calon Penerima Beasiswa"
        st.title(nama_app)
        st.header(header)
        st.subheader(subheader)
    
    # Fungsi menu sidebar
    def sidebar_menu(self):
        with st.sidebar:
            selected = option_menu('Menu',['Data','Pre Processing dan Transformation','DBI','Clustering'],default_index=0)
            
        if (selected == 'Data'):
            self.data.menu_data()

        if (selected == 'Pre Processing dan Transformation'):
            self.preprocessing.menu_preprocessing()

        if (selected == 'DBI'):
            self.dbi.menu_dbi()

        if (selected == 'Clustering'):
            self.clustering.menu_clustering()

class Data(MainClass):

    def __init__(self):
        # Membuat state untuk menampung dataframe
        self.state = st.session_state.setdefault('state', {})
        if 'datanilai' not in self.state:
            self.state['datanilai'] = pd.DataFrame()
        if 'dataekstrakurikuler' not in self.state:
            self.state['dataekstrakurikuler'] = pd.DataFrame()
        if 'dataekonomi' not in self.state:
            self.state['dataekonomi'] = pd.DataFrame()
        if 'dataprestasi' not in self.state:
            self.state['dataprestasi'] = pd.DataFrame()

    def upload_datanilai(self):
        try:
            uploaded_file1 = st.file_uploader("Upload Data Nilai", type=["xlsx"], key="nilai")
            if uploaded_file1 is not None:
                self.state['datanilai'] = pd.DataFrame()
                fnilai = pd.ExcelFile(uploaded_file1)

                # Membaca file excel dari banyak sheet
                list_of_dfs_nilai = []
                for sheet in fnilai.sheet_names:
            
                    # Parse data from each worksheet as a Pandas DataFrame
                    dfnilai = fnilai.parse(sheet, header=[0,1,2])

                    # And append it to the list
                    list_of_dfs_nilai.append(dfnilai)

                # Combine all DataFrames into one
                datanilai = pd.concat(list_of_dfs_nilai, ignore_index=True)
                datanilai = datanilai.rename(columns=lambda x: x if not 'Unnamed' in str(x) else '')
                datanilai.columns = datanilai.columns.map(' '.join)

                # Mengambil atribut
                datanilai = datanilai.iloc[:, [1, 2, 3, 18, 14, 26]]

                # Rename atribut dan validasi atribut
                datanilai.rename(columns = {'NIS  ':'NIS', 'Nama  ':'Nama',
                                'L/P  ':'L/P', 'Matematika (W) Pengetahuan Nilai':'Nilai Pengetahuan Matematika (W)',
                                'Bahasa Indonesia Pengetahuan Nilai':'Nilai Pengetahuan Bahasa Indonesia',
                                'Bahasa Inggris Pengetahuan Nilai':'Nilai Pengetahuan Bahasa Inggris'}, inplace = True)

                self.state['datanilai'] = datanilai
        except(TypeError, IndexError, KeyError):
            st.error("Data yang diupload tidak sesuai")

    def upload_dataekstrakurikuler(self):
        try:
            uploaded_file2 = st.file_uploader("Upload Data Ekstrakurikuler", type=["xlsx"], key="ekstrakurikuler")
            if uploaded_file2 is not None:
                self.state['dataekstrakurikuler'] = pd.DataFrame()
                fekstrakurikuler = pd.ExcelFile(uploaded_file2)

                # Membaca file excel dari banyak sheet
                list_of_dfs_ekstrakurikuler = []
                for sheet in fekstrakurikuler.sheet_names:
                    
                    # Parse data from each worksheet as a Pandas DataFrame
                    dfekstrakurikuler = fekstrakurikuler.parse(sheet)

                    # And append it to the list
                    list_of_dfs_ekstrakurikuler.append(dfekstrakurikuler)

                # Combine all DataFrames into one
                dataekstrakurikuler = pd.concat(list_of_dfs_ekstrakurikuler, ignore_index=True)

                # Mengambil atribut
                dataekstrakurikuler = dataekstrakurikuler.iloc[:, [1, 4]]
                dataekstrakurikuler = dataekstrakurikuler.fillna(method='ffill', limit = 4)

                # Validasi atribut
                dataekstrakurikuler.rename(columns = {'EKSTRAKURIKULER':'EKSTRAKURIKULER'}, errors="raise", inplace = True)

                self.state['dataekstrakurikuler'] = dataekstrakurikuler
        except(KeyError, IndexError):
            st.error("Data yang diupload tidak sesuai")

    def upload_dataekonomi(self):
        try:
            uploaded_file3 = st.file_uploader("Upload Data Ekonomi", type=["xlsx"], key="ekonomi")
            if uploaded_file3 is not None:
                self.state['dataekonomi'] = pd.DataFrame()
                fekonomi = pd.ExcelFile(uploaded_file3)

                # Membaca file excel dari banyak sheet
                list_of_dfs_ekonomi = []
                for sheet in fekonomi.sheet_names:
                    
                    # Parse data from each worksheet as a Pandas DataFrame
                    dfekonomi = fekonomi.parse(sheet)

                    # And append it to the list
                    list_of_dfs_ekonomi.append(dfekonomi)

                # Combine all DataFrames into one
                dataekonomi = pd.concat(list_of_dfs_ekonomi, ignore_index=True)

                # Mengambil atribut
                dataekonomi = dataekonomi.iloc[:, [0, 3, 4, 5, 6, 7, 8, 9, 10]]

                # Validasi atribut
                dataekonomi.rename(columns = {'Pekerjaan Ayah':'Pekerjaan Ayah'}, errors="raise", inplace = True)

                self.state['dataekonomi'] = dataekonomi
        except(KeyError, IndexError):
            st.error("Data yang diupload tidak sesuai")

    def upload_dataprestasi(self):
        try:
            uploaded_file4 = st.file_uploader("Upload Data Rekap Jumlah Prestasi", type=["xlsx"], key="prestasi")
            if uploaded_file4 is not None:
                self.state['dataprestasi'] = pd.DataFrame()
                dataprestasi = pd.read_excel(uploaded_file4)

                # Mengambil atribut
                dataprestasi = dataprestasi.iloc[:, [1, 9]]
                dataprestasi = dataprestasi.dropna()
                dataprestasi = dataprestasi.reset_index(drop=True)

                # Validasi atribut
                dataprestasi.rename(columns = {'Jumlah':'Jumlah'}, errors="raise", inplace = True)

                self.state['dataprestasi'] = dataprestasi
        except(KeyError, IndexError):
            st.error("Data yang diupload tidak sesuai")

    def tampil_datanilai(self):
        if not self.state['datanilai'].empty:
            st.dataframe(self.state['datanilai'])

    def tampil_dataekstrakurikuler(self):
        if not self.state['dataekstrakurikuler'].empty:
            st.dataframe(self.state['dataekstrakurikuler'])

    def tampil_dataekonomi(self):
        if not self.state['dataekonomi'].empty:
            st.dataframe(self.state['dataekonomi'])

    def tampil_dataprestasi(self):
        if not self.state['dataprestasi'].empty:
            st.dataframe(self.state['dataprestasi'])

    def menu_data(self):
        self.judul_halaman('Data','Import Dataset')
        self.upload_datanilai()
        self.tampil_datanilai()
        self.upload_dataekstrakurikuler()
        self.tampil_dataekstrakurikuler()
        self.upload_dataekonomi()
        self.tampil_dataekonomi()
        self.upload_dataprestasi()
        self.tampil_dataprestasi()

class Preprocessing(Data):

    def __init__(self):
        super().__init__()
        if 'dataset' not in self.state:
            self.state['dataset'] = pd.DataFrame()
        if 'tombol' not in self.state:
            self.state['tombol'] = 0

    def show_null_datanilai(self):
        st.subheader('Data Nilai')
        st.write("Jumlah nilai null pada data nilai")
        st.table(self.state['datanilai'].isnull().sum())

    def show_null_dataekstrakurikuler(self):
        st.subheader('Data Ekstrakurikuler')
        st.write("Jumlah nilai null pada data ekstrakurikuler:")

        self.state['dataekstrakurikulernull'] = self.state['dataekstrakurikuler'].isnull().sum()/5
        self.state['dataekstrakurikulernullstr'] = self.state['dataekstrakurikulernull'].astype(str).str.replace('\.0', '')

        st.table(self.state['dataekstrakurikulernullstr'])

    def show_null_dataekonomi(self):
        st.subheader('Data Ekonomi')
        st.write("Jumlah nilai null pada data ekonomi")
        st.table(self.state['dataekonomi'].isnull().sum())

    def show_null_dataprestasi(self):
        st.subheader('Data Prestasi')
        st.write("Jumlah nilai null pada data prestasi")
        st.table(self.state['dataprestasi'].isnull().sum())

    def iqr_datanilai(self):
        # IQR Matematika
        q11, q21, q31 = np.percentile(self.state['datanilai'].iloc[:, [3]], [25 , 50, 75])
        iqr1 = q31 - q11
        bbawah1 = q11 - 1.5 * iqr1
        batas1 = q31 + 1.5 * iqr1

        self.state['outliermtk'] = self.state['datanilai'].loc[(self.state['datanilai']['Nilai Pengetahuan Matematika (W)'] > batas1) | (self.state['datanilai']['Nilai Pengetahuan Matematika (W)'] < bbawah1)]
        self.state['sizeoutliermtk'] = self.state['outliermtk'].index.size

        if self.state['sizeoutliermtk'] > 0:
            st.write('Jumlah nilai outlier pada atribut Nilai Pengetahuan Matematika (W) =', str(self.state['sizeoutliermtk']))
            st.write('Data yang memiliki nilai outlier akan dihaluskan menggunakan metode binning')
            st.dataframe(self.state['outliermtk'])
        else:
            st.write('Tidak ditemukan nilai outlier pada atribut Nilai Pengetahuan Matematika (W)')

        # IQR Bahasa Indonesia
        q12, q22, q32 = np.percentile(self.state['datanilai'].iloc[:, [4]], [25 , 50, 75])
        iqr2 = q32 - q12
        bbawah2 = q12 - 1.5 * iqr2
        batas2 = q32 + 1.5 * iqr2

        self.state['outlierind'] = self.state['datanilai'].loc[(self.state['datanilai']['Nilai Pengetahuan Bahasa Indonesia'] > batas2) | (self.state['datanilai']['Nilai Pengetahuan Bahasa Indonesia'] < bbawah2)]
        self.state['sizeoutlierind'] = self.state['outlierind'].index.size

        if self.state['sizeoutlierind'] > 0:
            st.write('Jumlah nilai outlier pada atribut Nilai Pengetahuan Bahasa Indonesia =', str(self.state['sizeoutlierind']))
            st.write('Data yang memiliki nilai outlier akan dihaluskan menggunakan metode binning')
            st.dataframe(self.state['outlierind'])
        else:
            st.write('Tidak ditemukan nilai outlier pada atribut Nilai Pengetahuan Bahasa Indonesia')

        # IQR Bahasa Inggris
        q13, q23, q33 = np.percentile(self.state['datanilai'].iloc[:, [5]], [25 , 50, 75])
        iqr3 = q33 - q13
        bbawah3 = q13 - 1.5 * iqr3
        batas3 = q33 + 1.5 * iqr3

        self.state['outliering'] = self.state['datanilai'].loc[(self.state['datanilai']['Nilai Pengetahuan Bahasa Inggris'] > batas3) | (self.state['datanilai']['Nilai Pengetahuan Bahasa Inggris'] < bbawah3)]
        self.state['sizeoutliering'] = self.state['outliering'].index.size

        if self.state['sizeoutliering'] > 0:
            st.write('Jumlah nilai outlier pada atribut Nilai Pengetahuan Bahasa Inggris =', str(self.state['sizeoutliering']))
            st.write('Data yang memiliki nilai outlier akan dihaluskan menggunakan metode binning')
            st.dataframe(self.state['outliering'])
        else:
            st.write('Tidak ditemukan nilai outlier pada atribut Nilai Pengetahuan Bahasa Inggris')

    def iqr_dataprestasi(self):
        q14, q24, q34 = np.percentile(self.state['dataprestasi'].iloc[:, [1]], [25 , 50, 75])
        iqr4 = q34 - q14
        bbawah4 = q14 - 1.5 * iqr4
        batas4 = q34 + 1.5 * iqr4

        self.state['outlierprs'] = self.state['dataprestasi'].loc[(self.state['dataprestasi']['Jumlah'] > batas4) | (self.state['dataprestasi']['Jumlah'] < bbawah4)]
        self.state['sizeoutlierprs'] = self.state['outlierprs'].index.size

        if self.state['sizeoutlierprs'] > 0:
            st.write('Jumlah nilai outlier pada atribut Jumlah =', str(self.state['sizeoutlierprs']))
            st.write('Data yang memiliki nilai outlier akan dihaluskan menggunakan metode binning')
            st.dataframe(self.state['outlierprs'])
        else:
            st.write('Tidak ditemukan nilai outlier pada atribut Jumlah')

    def pre_processing(self):
        self.state['tombol'] = 1
        self.state['datasetasli'] = {}

        # Preprocessing data nilai
        self.state['datanilai'] = self.state['datanilai'].dropna()
        self.state['datanilai'][['NIS']] = self.state['datanilai'][['NIS']].astype(str)
        self.state['datanilai']['NIS'] = self.state['datanilai']['NIS'].str[:10]
        self.state['datanilai'] = self.state['datanilai'].sort_values(by=['NIS'],ascending=[True])
        self.state['datanilai'] = self.state['datanilai'].reset_index(drop=True)

        # Preprocessing data ekstrakurikuler
        self.state['dataekstrakurikuler'] = self.state['dataekstrakurikuler'].dropna()
        self.state['dataekstrakurikuler'][['NIS']] = self.state['dataekstrakurikuler'][['NIS']].astype(str)
        self.state['dataekstrakurikuler']['NIS'] = self.state['dataekstrakurikuler']['NIS'].str[:10]
        self.state['dataekstrakurikuler'] = self.state['dataekstrakurikuler'].sort_values(by=['NIS','EKSTRAKURIKULER'],ascending=[True, False])
        self.state['dataekstrakurikuler'].drop(self.state['dataekstrakurikuler'][self.state['dataekstrakurikuler']['EKSTRAKURIKULER'] == '-'].index, inplace=True)
        self.state['dataekstrakurikuler'] = self.state['dataekstrakurikuler'].reset_index(drop=True)

        self.state['dataekstrakurikuler']['Jumlah Ekstrakurikuler'] = self.state['dataekstrakurikuler'].groupby('NIS')['NIS'].transform('count')
        self.state['dataekstrakurikuler'] = self.state['dataekstrakurikuler'].iloc[:, [0, 2]]

        self.state['dataekstrakurikuler'] = self.state['dataekstrakurikuler'].drop_duplicates(subset=['NIS'])
        self.state['dataekstrakurikuler'] = self.state['dataekstrakurikuler'].reset_index(drop=True)

        # Preprocessing data ekonomi
        self.state['dataekonomi'] = self.state['dataekonomi'].dropna()
        self.state['dataekonomi'][['NIS']] = self.state['dataekonomi'][['NIS']].astype(str)
        self.state['dataekonomi']['NIS'] = self.state['dataekonomi']['NIS'].str[:10]
        self.state['dataekonomi'] = self.state['dataekonomi'].sort_values(by=['NIS'],ascending=[True])
        self.state['dataekonomi'] = self.state['dataekonomi'].reset_index(drop=True)

        # Preprocessing data prestasi            
        self.state['dataprestasi'][['NIS']] = self.state['dataprestasi'][['NIS']].astype(str)
        self.state['dataprestasi']['NIS'] = self.state['dataprestasi']['NIS'].str[:10]

        # Penggabungan data
        data = [self.state['datanilai'], self.state['dataekstrakurikuler'], self.state['dataekonomi'], self.state['dataprestasi']]
        self.state['dataset'] = reduce(lambda left, right: pd.merge(left,right,on=['NIS'], how='outer'), data)
        
        self.state['dataset'] = self.state['dataset'].iloc[:, [0, 1, 2, 7, 3, 4, 5, 8, 9, 10, 11, 12, 13, 14, 15, 6]]

        self.state['dataset']['Jumlah'] = self.state['dataset']['Jumlah'].fillna(0)
        self.state['dataset'] = self.state['dataset'].rename(columns={'Jumlah':'Jumlah Prestasi'})

        self.state['dataset'] = self.state['dataset'].dropna()
        self.state['dataset'] = self.state['dataset'].reset_index(drop=True)

        # Inisiasi datasetasli untuk menyimpan dataset bersih yang tidak dibinning
        self.state['datasetasli'] = self.state['dataset'].copy()
        self.state['datasetasli'] = self.state['datasetasli'].sort_values(by=['NIS'])

        # Binning data outlier
        binmtk = self.state['dataset'][['NIS','Nilai Pengetahuan Matematika (W)']].copy()
        binind = self.state['dataset'][['NIS','Nilai Pengetahuan Bahasa Indonesia']].copy()
        bining = self.state['dataset'][['NIS','Nilai Pengetahuan Bahasa Inggris']].copy()
        binprs = self.state['dataset'][['NIS','Jumlah Prestasi']].copy()

        binprs.drop(binprs[binprs['Jumlah Prestasi'] == 0].index, inplace = True)

        self.state['dataset'].drop(['Nilai Pengetahuan Matematika (W)','Nilai Pengetahuan Bahasa Indonesia',
                                    'Nilai Pengetahuan Bahasa Inggris', 'Jumlah Prestasi'], axis=1, inplace=True)
        
        if self.state['sizeoutliermtk'] > 0:
            # Menghitung total baris data
            row1 = binmtk.index.size

            # Mencari akar jumlah baris data
            squareroot1 = np.sqrt(row1)

            # Membulatkan hasil akar
            binsum1 = np.round(squareroot1)

            # Membuat label
            qlabels1 = []
            for i in range(1,int(binsum1)+1):
                t1 = i
                qlabels1.append(t1)

            # Mengurutkan data dari yang terkecil
            binmtk = binmtk.sort_values(by=['Nilai Pengetahuan Matematika (W)'])
            binmtk = binmtk.reset_index(drop=True)

            # Membuat bin dan memasukannya ke atribut Bin
            binsmtk = pd.qcut(binmtk.index, q=int(binsum1), labels=qlabels1)
            binmtk['Bin'] = binsmtk

            # Mengganti nilai berdasarkan nilai rata-rata bin dan menghapus atribut Bin
            binmtk['Nilai Pengetahuan Matematika (W)'] = binmtk.groupby('Bin')['Nilai Pengetahuan Matematika (W)'].transform('mean')
            binmtk.drop(['Bin'], axis=1, inplace=True)

        if self.state['sizeoutlierind'] > 0:
            # Menghitung total baris data
            row1 = binind.index.size

            # Mencari akar jumlah baris data
            squareroot1 = np.sqrt(row1)
            
            # Membulatkan hasil akar
            binsum1 = np.round(squareroot1)
            
            # Membuat label
            qlabels1 = []
            for i in range(1,int(binsum1)+1):
                t1 = i
                qlabels1.append(t1)

            # Mengurutkan data dari yang terkecil
            binind = binind.sort_values(by=['Nilai Pengetahuan Bahasa Indonesia'])
            binind = binind.reset_index(drop=True)

            # Membuat bin dan memasukannya ke atribut Bin
            binsind = pd.qcut(binind.index, q=int(binsum1), labels=qlabels1)
            binind['Bin'] = binsind

            # Mengganti nilai berdasarkan nilai rata-rata bin dan menghapus atribut Bin
            binind['Nilai Pengetahuan Bahasa Indonesia'] = binind.groupby('Bin')['Nilai Pengetahuan Bahasa Indonesia'].transform('mean')
            binind.drop(['Bin'], axis=1, inplace=True)

        if self.state['sizeoutliering'] > 0:
            # Menghitung total baris data
            row1 = bining.index.size

            # Mencari akar jumlah baris data
            squareroot1 = np.sqrt(row1)

            # Membulatkan hasil akar
            binsum1 = np.round(squareroot1)

            # Membuat label
            qlabels1 = []
            for i in range(1,int(binsum1)+1):
                t1 = i
                qlabels1.append(t1)

            # Mengurutkan data dari yang terkecil
            bining = bining.sort_values(by=['Nilai Pengetahuan Bahasa Inggris'])
            bining = bining.reset_index(drop=True)

            # Membuat bin dan memasukannya ke atribut Bin
            binsing = pd.qcut(bining.index, q=int(binsum1), labels=qlabels1)
            bining['Bin'] = binsing

            # Mengganti nilai berdasarkan nilai rata-rata bin dan menghapus atribut Bin
            bining['Nilai Pengetahuan Bahasa Inggris'] = bining.groupby('Bin')['Nilai Pengetahuan Bahasa Inggris'].transform('mean')
            bining.drop(['Bin'], axis=1, inplace=True)

        if self.state['sizeoutlierprs'] > 0:
            # Menghitung total baris data
            row2 = binprs.index.size

            # Mencari akar jumlah baris data
            squareroot2 = np.sqrt(row2)

            # Membulatkan hasil akar
            binsum2 = np.round(squareroot2)

            # Membuat label
            qlabels2 = []
            for i in range(1,int(binsum2)+1):
                t2 = i
                qlabels2.append(t2)

            # Mengurutkan data dari yang terkecil
            binprs = binprs.sort_values(by=['Jumlah Prestasi'])
            binprs = binprs.reset_index(drop=True)

            # Membuat bin dan memasukannya ke atribut Bin
            binsprs = pd.qcut(binprs.index, q=int(binsum2), labels=qlabels2)
            binprs['Bin'] = binsprs

            # Mengganti nilai berdasarkan nilai rata-rata bin dan menghapus atribut Bin
            binprs['Jumlah Prestasi'] = binprs.groupby('Bin')['Jumlah Prestasi'].transform('mean')
            binprs.drop(['Bin'], axis=1, inplace=True)

        # Dataset baru
        datadataset = [self.state['dataset'], binmtk, binind, bining, binprs]
        self.state['dataset'] = reduce(lambda left,right: pd.merge(left,right,on=['NIS'], how='outer'), datadataset)
        self.state['dataset']['Jumlah Prestasi'] = self.state['dataset']['Jumlah Prestasi'].fillna(0)
        self.state['dataset'] = self.state['dataset'].iloc[:, [0, 1, 2, 3, 12, 13, 14, 4, 5, 6, 7, 8, 9, 10, 15, 11]]
        self.state['dataset'][['Nilai Pengetahuan Matematika (W)', 'Nilai Pengetahuan Bahasa Indonesia',
                    'Nilai Pengetahuan Bahasa Inggris']] = self.state['dataset'][['Nilai Pengetahuan Matematika (W)',
                                                                                'Nilai Pengetahuan Bahasa Indonesia',
                                                                                'Nilai Pengetahuan Bahasa Inggris']].astype(float)

        # Membuat dataset untuk mining
        self.state['datasetAHC'] = self.state['dataset'].copy()

        self.state['datasetAHC']['Pekerjaan Ayah'] = self.state['datasetAHC']['Pekerjaan Ayah'].replace(['PNS/TNI/POLRI/BUMN/ASN/Guru', 'Pensiunan', 'Pegawai Swasta', 'Wiraswasta', 'Freelancer', 'Sopir/Driver', 'Security', 'Asisten Rumah Tangga/Cleaning Service', 'Petani/Nelayan', 'Tukang/Pekerjaan Tidak Tetap', 'Tidak Bekerja', 'Telah Meninggal Dunia'], [1,2,3,4,5,6,7,8,9,10,11,12])
        self.state['datasetAHC']['Pekerjaan Ibu'] = self.state['datasetAHC']['Pekerjaan Ibu'].replace(['PNS/TNI/POLRI/BUMN/ASN/Guru', 'Pensiunan', 'Pegawai Swasta', 'Wiraswasta', 'Freelancer', 'Sopir/Driver', 'Security', 'Asisten Rumah Tangga/Cleaning Service', 'Petani/Nelayan', 'Tukang/Pekerjaan Tidak Tetap', 'Tidak Bekerja', 'Telah Meninggal Dunia'], [1,2,3,4,5,6,7,8,9,10,11,12])
        self.state['datasetAHC']['Penghasilan Ayah'] = self.state['datasetAHC']['Penghasilan Ayah'].replace(['>7 juta', '6 - 7 juta', '5 - 5,9 juta', '4 - 4,9 juta', '3 - 3,9 juta', '2 - 2,9 juta', '1 - 1,9 juta', '500 - 900 ribu', '<500 ribu', 'Tidak Berpenghasilan'], [1,2,3,4,5,6,7,8,9,10])
        self.state['datasetAHC']['Penghasilan Ibu'] = self.state['datasetAHC']['Penghasilan Ibu'].replace(['>7 juta', '6 - 7 juta', '5 - 5,9 juta', '4 - 4,9 juta', '3 - 3,9 juta', '2 - 2,9 juta', '1 - 1,9 juta', '500 - 900 ribu', '<500 ribu', 'Tidak Berpenghasilan'], [1,2,3,4,5,6,7,8,9,10])
        self.state['datasetAHC']['Transportasi'] = self.state['datasetAHC']['Transportasi'].replace(['Sepeda Motor', 'Antar Jemput menggunakan Kendaraan Pribadi', 'Menumpang Teman', 'Ojek/Ojek Online', 'Sepeda', 'Transportasi Umum', 'Jalan Kaki'], [1,2,3,4,5,6,7])
        self.state['datasetAHC']['Memiliki KIP'] = self.state['datasetAHC']['Memiliki KIP'].replace(['Tidak', 'Ya'], [0,1])
        self.state['datasetAHC']['Jumlah Saudara Kandung'] = self.state['datasetAHC']['Jumlah Saudara Kandung'].replace(['Tidak Memiliki Saudara Kandung'], [0])

        # Pemilihan atribut yang akan dihitung untuk mining
        self.state['datasetAHC'] = self.state['datasetAHC'].iloc[:, 4:15]

    def tampil_dataset(self):
        if not self.state['dataset'].empty:
            st.subheader('Data Hasil Preprocessing dan Transformation')
            st.dataframe(self.state['dataset'])

    def menu_preprocessing(self):
        try:
            self.judul_halaman('Pre Processing dan Transformation','')
            if (not self.state['datanilai'].empty and not self.state['dataekstrakurikuler'].empty and not self.state['dataekonomi'].empty and not self.state['dataprestasi'].empty):
                self.show_null_datanilai()
                self.iqr_datanilai()
                self.show_null_dataekstrakurikuler()
                self.show_null_dataekonomi()
                self.show_null_dataprestasi()
                self.iqr_dataprestasi()
                if self.state['tombol'] == 0:
                    if st.button("Mulai Pre Processing dan Transformation"):
                        self.pre_processing()
                self.tampil_dataset()
            else:
                st.warning("Tidak ada data yang diupload atau data belum diupload sepenuhnya")

        except (IndexError):
            st.write('')

class Dbi(Data):

    def __init__(self):
        super().__init__()
        self.state['dbi'] = pd.DataFrame()

    # Fungsi perhitungan DBI
    def dbi(self, input1, input2):
        
        try:
            self.state['results'] = {}
            for i in range(input1,input2+1):
                hc = AgglomerativeClustering(n_clusters = i, affinity = 'euclidean', linkage = 'ward')
                y_hc = hc.fit_predict(self.state['datasetAHC'])
                db_index = davies_bouldin_score(self.state['datasetAHC'], y_hc)
                self.state['results'].update({i: db_index})
        except (ValueError):
            st.write('')

    # Fungsi menampilkan hasil evaluasi DBI
    def show_dbi(self):
        try:
            self.state['dbi'] = pd.DataFrame(self.state['results'].values(), self.state['results'].keys())
            if not self.state['dbi'].empty:
                st.table(self.state['results'])
                self.state['dbi'] = self.state['dbi'].round(4)
                st.write("Nilai terkecil adalah ", self.state['dbi'].min().min(), " dengan cluster sebanyak ", self.state['dbi'].idxmin().min())
            else:
                st.error("Nilai rentang cluster tidak valid")
        except(KeyError):
            st.write('')

    def menu_dbi(self):
        self.judul_halaman('DBI','')
        if (not self.state['datanilai'].empty and not self.state['dataekstrakurikuler'].empty and not self.state['dataekonomi'].empty and not self.state['dataprestasi'].empty):
            if not self.state['dataset'].empty:
                st.write('Tentukan Rentang Jumlah Cluster')
                col1, col2 = st.columns([1,1])
                with col1:
                    input1 = st.number_input('Dari', value=0, key=1)
                with col2:
                    input2 = st.number_input('Sampai', value=0, key=2)
                
                if st.button('Mulai'):
                    self.dbi(input1, input2)

                self.show_dbi()
            else:
                st.warning("Data belum dilakukan proses Pre Processing dan Transformation")
        else:
            st.warning("Tidak ada data yang diupload atau data belum diupload sepenuhnya")


class Clustering(Data):

    def __init__(self):
        super().__init__()
        if 'input_c' not in self.state:
            self.state['input_c'] = None
        if 'dfi' not in self.state:
            self.state['dfi'] = {}

    def clustering(self, input_c):
        try:
            self.state['nrs'] = {}
            self.state['nrs_pna'] = {}
            self.state['nrs_pni'] = {}
            self.state['tr'] = {}

            self.state['nrs_pna_rek'] = {}
            self.state['nrs_pni_rek'] = {}
            self.state['tr_rek'] = {}

            self.state['dfi'] = {}
            self.state['rekomendasi'] = {}
            self.state['datarekomendasi'] = {}

            self.state['clustering'] = self.state['datasetAHC'].copy()
            self.state['datahasil'] = self.state['dataset'].copy()

            # Proses AHC
            hc = AgglomerativeClustering(n_clusters = input_c, affinity = 'euclidean', linkage = 'ward')
            self.state['y_hc'] = hc.fit_predict(self.state['datasetAHC'])

            # Memberikan label cluster pada baris data
            self.state['datasetasli']['cluster'] = pd.DataFrame(self.state['y_hc'])
            self.state['datasetasli'] = self.state['datasetasli'].sort_values(by='cluster')
            self.state['datasetasli'] = self.state['datasetasli'].reset_index(drop=True)
            self.state['datasetasli']['cluster'] = self.state['datasetasli']['cluster']+1
            
        except(ValueError, IndexError):
            st.error("Nilai jumlah cluster tidak valid")

    def show_cluster(self, input_c):
        try:
            for i in range(1,input_c+1):
                # Memisahkan data ke masing-masing dataframe percluster
                self.state['dfi']["clustering{0}".format(i)] = self.state['datasetasli'].loc[self.state['datasetasli']['cluster'] == i+1-1]
                self.state['nrs']["clustering{0}".format(i)] = self.state['datasetasli'].loc[self.state['datasetasli']['cluster'] == i+1-1]

                # Analisis karakteristik penghasilan ayah
                self.state['nrs_pna']["clustering{0}".format(i)] = self.state['nrs']["clustering"+str(i+1-1)]['Penghasilan Ayah'].value_counts()
                self.state['nrs_pna']["clustering"+str(i+1-1)] = pd.DataFrame(self.state['nrs_pna']["clustering"+str(i+1-1)])
                self.state['nrs_pna']["clustering"+str(i+1-1)]['value'] = self.state['nrs_pna']["clustering"+str(i+1-1)].index
                self.state['nrs_pna']["clustering"+str(i+1-1)] = self.state['nrs_pna']["clustering"+str(i+1-1)].sort_values(by = ['Penghasilan Ayah', 'value'], ascending = [False, False])
                self.state['nrs_pna_rek']["clustering{0}".format(i)] = self.state['nrs_pna']["clustering"+str(i+1-1)].copy()
                self.state['nrs_pna']["clustering"+str(i+1-1)]['value'] = self.state['nrs_pna']["clustering"+str(i+1-1)]['value'].replace([1,2,3,4,5,6,7,8,9,10],['berpenghasilan lebih dari 7 juta rupiah', 'berpenghasilan 6 sampai 7 juta rupiah', 'berpenghasilan 5 sampai 5,9 juta rupiah', 'berpenghasilan 4 sampai 4,9 juta rupiah', 'berpenghasilan 3 sampai 3,9 juta rupiah', 'berpenghasilan 2 sampai 2,9 juta rupiah', 'berpenghasilan 1 sampai 1,9 juta rupiah', 'berpenghasilan 500 sampai 900 ribu rupiah', 'berpenghasilan kurang dari 500 ribu rupiah', 'tidak berpenghasilan'])

                # Analisis karakteristik penghasilan ibu
                self.state['nrs_pni']["clustering{0}".format(i)] = self.state['nrs']["clustering"+str(i+1-1)]['Penghasilan Ibu'].value_counts()
                self.state['nrs_pni']["clustering"+str(i+1-1)] = pd.DataFrame(self.state['nrs_pni']["clustering"+str(i+1-1)])
                self.state['nrs_pni']["clustering"+str(i+1-1)]['value'] = self.state['nrs_pni']["clustering"+str(i+1-1)].index
                self.state['nrs_pni']["clustering"+str(i+1-1)] = self.state['nrs_pni']["clustering"+str(i+1-1)].sort_values(by = ['Penghasilan Ibu', 'value'], ascending = [False, False])
                self.state['nrs_pni_rek']["clustering{0}".format(i)] = self.state['nrs_pni']["clustering"+str(i+1-1)].copy()
                self.state['nrs_pni']["clustering"+str(i+1-1)]['value'] = self.state['nrs_pni']["clustering"+str(i+1-1)]['value'].replace([1,2,3,4,5,6,7,8,9,10],['berpenghasilan lebih dari 7 juta rupiah', 'berpenghasilan 6 sampai 7 juta rupiah', 'berpenghasilan 5 sampai 5,9 juta rupiah', 'berpenghasilan 4 sampai 4,9 juta rupiah', 'berpenghasilan 3 sampai 3,9 juta rupiah', 'berpenghasilan 2 sampai 2,9 juta rupiah', 'berpenghasilan 1 sampai 1,9 juta rupiah', 'berpenghasilan 500 sampai 900 ribu rupiah', 'berpenghasilan kurang dari 500 ribu rupiah', 'tidak berpenghasilan'])

                # Analisis karakteristik transportasi
                self.state['tr']["clustering{0}".format(i)] = self.state['nrs']["clustering"+str(i+1-1)]['Transportasi'].value_counts()
                self.state['tr']["clustering"+str(i+1-1)] = pd.DataFrame(self.state['tr']["clustering"+str(i+1-1)])
                self.state['tr']["clustering"+str(i+1-1)]['value'] = self.state['tr']["clustering"+str(i+1-1)].index
                self.state['tr']["clustering"+str(i+1-1)] = self.state['tr']["clustering"+str(i+1-1)].sort_values(by = ['Transportasi', 'value'], ascending = [False, False])
                self.state['tr_rek']["clustering{0}".format(i)] = self.state['tr']["clustering"+str(i+1-1)].copy() 
                self.state['tr']["clustering"+str(i+1-1)]['value'] = self.state['tr']["clustering"+str(i+1-1)]['value'].replace([1,2,3,4,5,6,7],['menggunakan kendaraan sepeda motor', 'dengan diantar jemput menggunakan kendaraan pribadi', 'dengan menumpang teman', 'menggunakan ojek atau ojek online', 'menggunakan sepeda', 'menggunakan transportasi umum', 'dengan berjalan kaki'])            

            rekomendasi = []

            for i in range(1,input_c+1):
                # Inisiasi karakteristik penghasilan ayah, penghasilan ibu, transportasi, dan rata2 ke3 mapel per cluster
                pna = str(self.state['nrs_pna']["clustering"+str(i+1-1)]._get_value(0,1,takeable = True))
                pni = str(self.state['nrs_pni']["clustering"+str(i+1-1)]._get_value(0,1,takeable = True))
                tr = str(self.state['tr']["clustering"+str(i+1-1)]._get_value(0,1,takeable = True))
                mtk = str(round(self.state['nrs']["clustering"+str(i+1-1)]['Nilai Pengetahuan Matematika (W)'].mean(),4))
                bind = str(round(self.state['nrs']["clustering"+str(i+1-1)]['Nilai Pengetahuan Bahasa Indonesia'].mean(),4))
                bing = str(round(self.state['nrs']["clustering"+str(i+1-1)]['Nilai Pengetahuan Bahasa Inggris'].mean(),4))

                # Mengambil karakteristik penghasilan ayah, penghasilan ibu, dan transportasi masing2 cluster
                self.state['pnarek'] = str(self.state['nrs_pna_rek']["clustering"+str(i+1-1)]._get_value(0,1,takeable = True))
                self.state['pnirek'] = str(self.state['nrs_pni_rek']["clustering"+str(i+1-1)]._get_value(0,1,takeable = True))
                self.state['trrek'] = str(self.state['tr_rek']["clustering"+str(i+1-1)]._get_value(0,1,takeable = True))

                # Reset index per dataframe berbeda
                self.state['dfi']["clustering"+str(i+1-1)] = self.state['dfi']["clustering"+str(i+1-1)].reset_index(drop=True)
                self.state['dfi']["clustering"+str(i+1-1)].index += 1

                # Menampilkan tabel cluster dan karakteristiknya
                terbilang_angka = num2words(i, lang='id', to='ordinal')
                st.write('**Cluster** ' + terbilang_angka)
                st.dataframe(self.state['dfi']["clustering"+str(i+1-1)])

                st.write('Terlihat bahwa anggota yang tergabung ke dalam cluster ' + str(i),
                            'merupakan siswa yang memiliki nilai rata-rata mata pelajaran Matematika bernilai ' +
                            mtk + ', Bahasa Indonesia bernilai ' + bind + ', dan Bahasa Inggris bernilai ' + bing,
                            '. Kemudian, siswa yang tergabung ke dalam kelompok ini rata-rata memiliki ayah yang ' +
                            pna + ' dan memiliki ibu yang ' + pni +
                            '. Selain itu, siswa yang tergabung ke dalam kelompok ini rata-rata berangkat ke sekolah ' +
                            tr)
                st.write(''); st.write('')

                # Menyatukan karakteristik menjadi satu dataframe
                rowrekomendasi = [self.state['pnarek'], self.state['pnirek'], self.state['trrek'], mtk, bind, bing, terbilang_angka]
                rekomendasi.append(rowrekomendasi)
                self.state['rekomendasi'] = pd.DataFrame(rekomendasi)
                
            # Ubah value ke numeric 
            self.state['rekomendasi'].columns = ['a','b','c','d','e','f','g']
            self.state['rekomendasi']['a'] = self.state['rekomendasi']['a'].replace(['>7 juta', '6 - 7 juta', '5 - 5,9 juta', '4 - 4,9 juta', '3 - 3,9 juta', '2 - 2,9 juta', '1 - 1,9 juta', '500 - 900 ribu', '<500 ribu', 'Tidak Berpenghasilan'], [1,2,3,4,5,6,7,8,9,10])
            self.state['rekomendasi']['b'] = self.state['rekomendasi']['b'].replace(['>7 juta', '6 - 7 juta', '5 - 5,9 juta', '4 - 4,9 juta', '3 - 3,9 juta', '2 - 2,9 juta', '1 - 1,9 juta', '500 - 900 ribu', '<500 ribu', 'Tidak Berpenghasilan'], [1,2,3,4,5,6,7,8,9,10])
            self.state['rekomendasi']['c'] = self.state['rekomendasi']['c'].replace(['Sepeda Motor', 'Antar Jemput menggunakan Kendaraan Pribadi', 'Menumpang Teman', 'Ojek/Ojek Online', 'Sepeda', 'Transportasi Umum', 'Jalan Kaki'], [1,2,3,4,5,6,7])

            # Membuat DataFrame baru dari dataframe rekomendasi untuk menampilkan deskripsi hasil rekomendasi
            self.state['datarekomendasi'] = pd.DataFrame(self.state['rekomendasi'], columns=['a','b','c','d','e','f','g'])
            self.state['datarekomendasi'] = self.state['datarekomendasi'].astype({'a':'int','b':'int','c':'int'})
            self.state['datarekomendasi'] = self.state['datarekomendasi'].sort_values(by = ['a','d','b','e','c','f'], ascending = [False,False,False,False,False,False])
            self.state['datarekomendasi']['a'] = self.state['datarekomendasi']['a'].replace([1,2,3,4,5,6,7,8,9,10],['berpenghasilan lebih dari 7 juta rupiah', 'berpenghasilan 6 sampai 7 juta rupiah', 'berpenghasilan 5 sampai 5,9 juta rupiah', 'berpenghasilan 4 sampai 4,9 juta rupiah', 'berpenghasilan 3 sampai 3,9 juta rupiah', 'berpenghasilan 2 sampai 2,9 juta rupiah', 'berpenghasilan 1 sampai 1,9 juta rupiah', 'berpenghasilan 500 sampai 900 ribu rupiah', 'berpenghasilan kurang dari 500 ribu rupiah', 'tidak berpenghasilan'])
            self.state['datarekomendasi']['b'] = self.state['datarekomendasi']['b'].replace([1,2,3,4,5,6,7,8,9,10],['berpenghasilan lebih dari 7 juta rupiah', 'berpenghasilan 6 sampai 7 juta rupiah', 'berpenghasilan 5 sampai 5,9 juta rupiah', 'berpenghasilan 4 sampai 4,9 juta rupiah', 'berpenghasilan 3 sampai 3,9 juta rupiah', 'berpenghasilan 2 sampai 2,9 juta rupiah', 'berpenghasilan 1 sampai 1,9 juta rupiah', 'berpenghasilan 500 sampai 900 ribu rupiah', 'berpenghasilan kurang dari 500 ribu rupiah', 'tidak berpenghasilan'])
            self.state['datarekomendasi']['c'] = self.state['datarekomendasi']['c'].replace([1,2,3,4,5,6,7],['menggunakan kendaraan sepeda motor', 'dengan diantar jemput menggunakan kendaraan pribadi', 'dengan menumpang teman', 'menggunakan ojek atau ojek online', 'menggunakan sepeda', 'menggunakan transportasi umum', 'dengan berjalan kaki'])            

            # Mengambil nilai karakteristik teratas untuk menghasilkan cluster yang direkomendasikan
            arek = str(self.state['datarekomendasi']['a'].iloc[0])
            brek = str(self.state['datarekomendasi']['b'].iloc[0])
            crek = str(self.state['datarekomendasi']['c'].iloc[0])
            drek = str(self.state['datarekomendasi']['d'].iloc[0])
            erek = str(self.state['datarekomendasi']['e'].iloc[0])
            frek = str(self.state['datarekomendasi']['f'].iloc[0])
            grek = str(self.state['datarekomendasi']['g'].iloc[0])

            # Menuliskan hasil rekomendasi
            st.write('**Kesimpulan Rekomendasi**')
            st.write('Kelompok yang direkomendasikan untuk mendapatkan beasiswa adalah cluster ' + grek+
                    ', di mana siswa dalam kelompok ini memiliki nilai rata-rata mata pelajaran Matematika bernilai ' +
                    drek + ', Bahasa Indonesia bernilai ' + erek + ', dan Bahasa Inggris bernilai ' + frek,
                    '. Kemudian, siswa yang tergabung ke dalam kelompok ini rata-rata memiliki ayah yang ' +
                    arek + ' dan memiliki ibu yang ' + brek +
                    '. Selain itu, siswa yang tergabung ke dalam kelompok ini rata-rata berangkat ke sekolah ' +
                    crek)
                
        except(TypeError, KeyError, IndexError, AttributeError):
            st.write('')

    # Fungsi ubah dataframe hasil clustering ke bentuk excel
    def to_excel(self, df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data
    
    # Fungsi mendownload file excel hasil clustering
    def download_clustering(self):
        df_xlsx = self.to_excel(self.state['datasetasli'])
        st.write('')
        st.write('')
        st.download_button(label='Download Hasil Clustering',
                                data=df_xlsx ,
                                file_name= 'hasil_clustering.xlsx')

    def menu_clustering(self):
        self.judul_halaman('Clustering','')
        if (not self.state['datanilai'].empty and not self.state['dataekstrakurikuler'].empty and not self.state['dataekonomi'].empty and not self.state['dataprestasi'].empty):
            if not self.state['dataset'].empty:
                input_c = st.number_input('Tentukan Jumlah Cluster',value=0)
                if st.button('Mulai Clustering'):
                    st.write('')
                    self.clustering(input_c)
                    self.state['input_c'] = input_c
                if not self.state['dataset'].empty:
                    self.show_cluster(self.state['input_c'])
                else:
                    st.warning("Tidak ada data yang diupload atau data kosong")
                if self.state['dfi']:
                    self.download_clustering()
            else:
                st.warning("Data belum dilakukan proses Pre Processing dan Transformation")
        else:
            st.warning("Tidak ada data yang diupload atau data belum diupload sepenuhnya")

if __name__ == "__main__":
    # Create an instance of the main class
    main = MainClass()
    
main.sidebar_menu()
