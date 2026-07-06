import streamlit as st
import pandas as pd
import re
import io

# Mengatur tampilan halaman web
st.set_page_config(page_title="Perapih Data Excel", layout="wide")
st.title("✨ Aplikasi Perapih Data Excel")
st.write("Silakan upload file **.xlsx** mentah.")

def ubah_format_tanggal(date_val):
    if pd.isna(date_val):
        return ""
    # Mengambil bagian tanggal saja (contoh: '2026-07-01')
    date_str = str(date_val).strip().split(' ')[0]
    
    # Memisahkan berdasarkan '-' lalu menyusunnya menjadi DD/MM/YYYY
    try:
        parts = date_str.split('-')
        if len(parts) == 3:
            return f"{parts[2]}/{parts[1]}/{parts[0]}"
    except Exception:
        pass
    
    return date_str

def rapikan_nominal(angka):
    if pd.isna(angka) or str(angka).strip() == "":
        return ""
    angka_str = str(angka).strip()
    
    # Menghapus akhiran .00 jika ada
    if angka_str.endswith(".00"):
        angka_str = angka_str[:-3]
        
    # Mengganti koma menjadi titik
    angka_str = angka_str.replace(",", ".")
    return angka_str

def pisahkan_deskripsi(desc):
    if pd.isna(desc):
        return "", ""
    desc = str(desc).strip()
    
    # 1. Kondisi: BIFAST
    if "BIFAST" in desc:
        match_bank = re.search(r'Bank\b', desc, flags=re.IGNORECASE)
        if match_bank:
            sisa_kalimat = desc[match_bank.end():].strip()
            words = sisa_kalimat.split()
            
            fase = 1
            nama_words = []
            uraian_words = []
            
            for w in words:
                is_upper = w.isupper() and any(c.isalpha() for c in w)
                
                if fase == 1:
                    if is_upper:
                        fase = 2
                        nama_words.append(w)
                elif fase == 2:
                    if is_upper:
                        nama_words.append(w)
                    else:
                        fase = 3
                        uraian_words.append(w)
                elif fase == 3:
                    uraian_words.append(w)
                    
            nama = " ".join(nama_words)
            uraian = " ".join(uraian_words)
            return uraian, nama
        return desc, ""
        
    # 2. Kondisi: [Uraian] Trf Dari - [Nama]
    elif not desc.startswith("Trf Dari") and re.search(r'Trf Dari\s*-', desc, flags=re.IGNORECASE):
        # Pisahkan berdasarkan kata "Trf Dari -"
        match = re.search(r'(.*)\s*Trf Dari\s*-\s*(.*)', desc, flags=re.IGNORECASE)
        if match:
            uraian = match.group(1).strip()
            nama = match.group(2).strip()
            return uraian, nama
        return desc, ""

    # 3. Kondisi: Mengandung "Trf Dari" di awal kalimat
    elif desc.startswith("Trf Dari"):
        # Pola Lama: Trf Dari - 009 - - 206056 GXPP6Y0Z INDHIFA MU
        prefix_match = re.match(r'Trf Dari\s*-\s*\d+\s*-\s*-\s*', desc)
        if prefix_match:
            sisa_kalimat = desc[prefix_match.end():]
            words = sisa_kalimat.split()
            if len(words) >= 2:
                uraian = " ".join(words[:2]) 
                nama = " ".join(words[2:])   
                return uraian, nama
            return desc, ""
            
        # Parameter Baru 1: Trf Dari - 008 - 305876 K9XFM8LZ DAPUR PAHLAWAN
        match_baru = re.match(r'^Trf Dari\s*-\s*\d+\s*-\s*(\S+\s+\S+)\s+(.*)$', desc, flags=re.IGNORECASE)
        if match_baru:
            uraian = match_baru.group(1).strip()
            nama = match_baru.group(2).strip()
            return uraian, nama
            
        return desc, ""
        
    # Parameter Baru 2: Trf Bersama to BSI - Bersama\000000024634\252438
    elif desc.startswith("Trf Bersama to BSI"):
        match_bersama = re.match(r'^Trf Bersama to BSI\s*-\s*(.*)$', desc, flags=re.IGNORECASE)
        if match_bersama:
            uraian = match_bersama.group(1).strip()
            return uraian, ""
        return desc, ""
        
    # 5. Kondisi: Selain di atas
    else:
        return desc, ""

# Area Upload File
file_unggahan = st.file_uploader("Upload File Excel Mentah (.xlsx)", type=["xlsx"])

if file_unggahan is not None:
    df_temp = pd.read_excel(file_unggahan, header=None, nrows=20)
    
    try:
        header_idx = df_temp[df_temp.eq('Date').any(axis=1)].index[0]
        df = pd.read_excel(file_unggahan, skiprows=header_idx)
        
        st.write("### Data Mentah (Tabel Utama Ditemukan):")
        # Menampilkan hanya 5 baris pertama untuk data mentah
        st.dataframe(df.head())
        st.divider()
        
        if st.button("Proses & Rapihkan Data"):
            df_rapih = pd.DataFrame()
            
            # 1. Date -> DD/MM/YYYY
            df_rapih['Tgl / Bln / Thn'] = df['Date'].apply(ubah_format_tanggal)
            
            # 2. Pisahkan Uraian dan Nama dari Description
            hasil_pisah = df['Description'].apply(pisahkan_deskripsi)
            df_rapih['Uraian'] = [res[0] for res in hasil_pisah]
            df_rapih['Nama'] = [res[1] for res in hasil_pisah]
            
            # Siapkan kolom kosong untuk Debit dan Credit
            df_rapih['Debit'] = ""
            df_rapih['Credit'] = ""
            
            # 3. Logika DB -> Amount pindah jalur dengan Nominal Baru
            for index, row in df.iterrows():
                # Terapkan perapihan nominal pada amount
                amount_rapih = rapikan_nominal(row['Amount'])
                db_val = str(row['DB']).strip()
                
                if db_val == "DB":
                    df_rapih.at[index, 'Credit'] = amount_rapih
                else:
                    df_rapih.at[index, 'Debit'] = amount_rapih
                    
            # 4. Saldo diambil dan dirapihkan nominalnya juga
            df_rapih['Saldo'] = df['Balance'].apply(rapikan_nominal)
            
            st.success("✅ Data berhasil dirapihkan!")
            st.write("### Data Hasil (Setelah Diproses):")
            
            # PERUBAHAN DI SINI:
            # Menghapus .head() agar Streamlit merender seluruh DataFrame df_rapih
            st.dataframe(df_rapih)
            
            # Konversi dataframe ke excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_rapih.to_excel(writer, index=False, sheet_name="Data Rapih")
            excel_data = output.getvalue()
            
            st.download_button(
                label="⬇️ Download Excel Rapih (.xlsx)",
                data=excel_data,
                file_name='data_sudah_rapih.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
    except IndexError:
        st.error("⚠️ Tidak dapat menemukan kolom 'Date'. Pastikan format file tidak berubah dari contoh aslinya.")
    except Exception as e:
        st.error(f"⚠️ Terjadi kesalahan: {e}")
