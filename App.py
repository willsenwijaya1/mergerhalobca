import streamlit as st
import pandas as pd
import io

# Fungsi untuk ekstrak tabel dari file Excel
def extract_table(file):
    df_raw = pd.read_excel(file, header=None)

    # Mencari baris yang mengandung header (Kolom 1, Kolom 2, dst.)
    start_row = None
    for i, row in df_raw.iterrows():
        if row.astype(str).str.contains("Download", case=False, na=False).any():
            start_row = i
            break

    if start_row is None:
        st.error(f"⚠ Header tidak ditemukan dalam file {file.name}")
        return None

    # Membaca ulang file dari baris header
    df = pd.read_excel(file, skiprows=start_row)

    # Menghapus baris yang tidak diperlukan
    keywords = ["RESPONSE RATE", "PASSED RESPONSE RATE", "FAILED RESPONSE RATE", "ABANDONES RATE"]
    rows_to_drop = df[df.apply(lambda row: row.astype(str).isin(keywords).any(), axis=1)].index
    for index in rows_to_drop:
        df = df.drop(index, errors='ignore')
        df = df.drop(index + 1, errors='ignore')  # Hapus baris di bawahnya jika ada

    return df.reset_index(drop=True)

# Fungsi untuk ekstrak tanggal periode
def extract_periode_date(file):
    df_raw = pd.read_excel(file, header=None)

    # Mencari baris yang mengandung kata "PERIODE"
    for i, row in df_raw.iterrows():
        for j, cell in enumerate(row):
            if isinstance(cell, str) and "PERIODE" in cell.upper():
                # Mengambil tanggal di sebelah kanan jika ada
                if j + 1 < len(row):
                    date_value = row[j + 1]
                    try:
                        return pd.to_datetime(date_value, errors='coerce').strftime("%d-%b-%Y")
                    except:
                        return None
    return None

# --- Streamlit UI ---
st.title("📊 Data Merger - Halo BCA Chat")

st.write("Unggah file Excel untuk digabungkan secara otomatis")

uploaded_files = st.file_uploader("Pilih beberapa file Excel", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    dataframes = []
    
    for file in uploaded_files:
        df = extract_table(file)
        if df is not None:
            date = extract_periode_date(file)
            if date:
                df["Periode"] = date  # Tambahkan kolom periode
            
            dataframes.append(df)

    if dataframes:
        merged_df = pd.concat(dataframes, ignore_index=True)

        # Menghapus kolom 'Unnamed' jika ada
        merged_df = merged_df.loc[:, ~merged_df.columns.str.contains('^Unnamed')]

        # Menampilkan hasil
        st.write("### 📌 Data Gabungan")
        st.dataframe(merged_df)

        # Simpan file hasil gabungan
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False)
        
        output.seek(0)
        
        st.download_button(
            label="⬇ Unduh Data Gabungan",
            data=output,
            file_name="Merged.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("⚠ Tidak ada data yang berhasil diekstrak.")

