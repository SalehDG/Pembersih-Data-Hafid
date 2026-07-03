import streamlit as st
import pandas as pd
import re
import io

# Mengatur tampilan halaman web
st.set_page_config(page_title="Perapih Data Excel BSI Hafid"  , layout="wide")
st.title("✨ Perapih Data Excel BSI Hafid")
st.write("Silakan upload file **.xlsx** mentah.")

def ambil_tanggal(date_val):
    if pd.isna(date_val):
        return ""
    # Memotong string berdasarkan spasi dan mengambil bagian pertamanya saja (Tanggal)
    return str(date_val).strip().split(' ')[0]

def pisahkan_deskripsi(desc):
    if pd.isna(desc):
        return "", ""
    desc = str(desc).strip()
    
    # 1. Kondisi: BIFAST
    if "BIFAST" in desc:
        # Mencari awalan kata 'Bank'
        match_bank = re.search(r'Bank\b', desc, flags=re.IGNORECASE)
        if match_bank:
            sisa_kalimat = desc[match_bank.end():].strip()
            words = sisa_kalimat.split()
            
            fase = 1
            nama_words = []
            uraian_words = []
            
            for w in words:
                # Cek apakah kata tersebut kapital semua (menandakan NAMA)
                is_upper = w.isupper() and any(c.isalpha() for c in w)
                
                if fase == 1:
                    if is_upper:
                        fase = 2  # Mulai masuk ke NAMA
                        nama_words.append(w)
                elif fase == 2:
                    if is_upper:
                        nama_words.append(w) # Masih rentetan NAMA
                    else:
                        fase = 3  # Kembali ke huruf normal, masuk ke URAIAN
                        uraian_words.append(w)
                elif fase == 3:
                    uraian_words.append(w) # Sisa kalimat masuk ke URAIAN semua
                    
            nama = " ".join(nama_words)
            uraian = " ".join(uraian_words)
            return uraian, nama
        return desc, ""
        
    # 2. Kondisi: Trf Dari
    elif desc.startswith("Trf Dari"):
        # Membuang pola awalan seperti "Trf Dari - 009 - - "
        prefix_match = re.match(r'Trf Dari\s*-\s*\d+\s*-\s*-\s*', desc)
        if prefix_match:
            sisa_kalimat = desc[prefix_match.end():]
            words = sisa_kalimat.split()
            if len(words) >= 2:
                uraian = " ".join(words[:2])  # 2 kata pertama adalah Uraian
                nama = " ".join(words[2:])    # Sisanya adalah Nama
                return uraian, nama
        return desc, ""
        
    # 3. Kondisi: Selain 2 di atas
    else:
        return desc, ""

# Area Upload File
file_unggahan = st.file_uploader("Upload File Excel Mentah (.xlsx)", type=["xlsx"])

if file_unggahan is not None:
    # Membaca 20 baris pertama untuk mencari di baris mana tabel sebenarnya dimulai
    df_temp = pd.read_excel(file_unggahan, header=None, nrows=20)
    
    try:
        # Mencari baris yang mengandung tulisan "Date" sebagai patokan header utama
        header_idx = df_temp[df_temp.eq('Date').any(axis=1)].index[0]
        
        # Membaca ulang data dengan menjadikan baris tersebut sebagai Header yang sebenarnya
        df = pd.read_excel(file_unggahan, skiprows=header_idx)
        
        st.write("### Data Mentah (Tabel Utama Ditemukan):")
        st.dataframe(df.head())
        st.divider()
        
        # Tombol Eksekusi
        if st.button("Proses & Rapihkan Data"):
            # Membuat keranjang (dataframe) baru kosong
            df_rapih = pd.DataFrame()
            
            # 1. Date -> Tgl / Bln / Thn
            df_rapih['Tgl / Bln / Thn'] = df['Date'].apply(ambil_tanggal)
            
            # 2. Pisahkan Uraian dan Nama dari Description
            hasil_pisah = df['Description'].apply(pisahkan_deskripsi)
            df_rapih['Uraian'] = [res[0] for res in hasil_pisah]
            df_rapih['Nama'] = [res[1] for res in hasil_pisah]
            
            # Siapkan kolom kosong untuk Debit dan Credit
            df_rapih['Debit'] = ""
            df_rapih['Credit'] = ""
            
            # 3. Logika DB -> Amount pindah jalur
            for index, row in df.iterrows():
                amount = row['Amount']
                db_val = str(row['DB']).strip()
                
                # Sesuai aturan: jika DB berisi "DB", amount -> Credit. Jika tidak -> Debit.
                if db_val == "DB":
                    df_rapih.at[index, 'Credit'] = amount
                else:
                    df_rapih.at[index, 'Debit'] = amount
                    
            # 4. Saldo diambil mentah
            df_rapih['Saldo'] = df['Balance']
            
            st.success("✅ Data berhasil dirapihkan!")
            st.write("### Data Hasil (Setelah Diproses):")
            st.dataframe(df_rapih.head())
            
            # Konversi dataframe ke dalam format file .xlsx
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_rapih.to_excel(writer, index=False, sheet_name="Data Rapih")
            excel_data = output.getvalue()
            
            # Tombol Download Excel Bersih
            st.download_button(
                label="⬇️ Download Excel Rapih (.xlsx)",
                data=excel_data,
                file_name='data_sudah_rapih.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
    except IndexError:
        st.error("⚠️ Tidak dapat menemukan kolom 'Date' pada file ini. Pastikan format file tidak berubah dari contoh aslinya.")
    except Exception as e:
        st.error(f"⚠️ Terjadi kesalahan: {e}")