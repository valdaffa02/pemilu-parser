import openpyxl
import pandas as pd
import sys
import os
import glob
import re
import streamlit as st
import io

def load_excel_sheet(file_path):
    """
    Responsibility: Opens the Excel workbook and returns the active sheet.
    """
    print(f"Loading workbook '{file_path}'...")
    try:
        # data_only=True ensures we read calculated values, not raw formulas
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        sheet = workbook.active
        print(f"Successfully loaded sheet: '{sheet.title}'")
        return sheet
        
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred while loading: {e}")
        return None


def extract_vote_by_party(sheet):
    """
    Parses the election tally sheet to extract district headers and party vote counts.
    """
    # Initialize our arrays
    kelurahan = []
    suara = []
    
    # 1.2.2c: Track which party is the latest one
    current_party_index = -1 
    
    # 1.1: Loop through the loaded xlsx file
    for row in sheet.iter_rows(values_only=True):
        
        # We loop through by index (i) so we can easily look "to the right" (i + 1)
        for i, cell_value in enumerate(row):
            if cell_value is None:
                continue
            
            # Clean the current cell string for safer pattern matching
            cell_str = str(cell_value).strip().upper().replace('\n', ' ')
            
            # Find kecamatan
            if cell_str == "KECAMATAN/DISTRIK *)":
                for j in range(i + 1, len(row)):
                    right_val = row[j]
                    
                    # 1. Skip if the cell is completely empty or None
                    if right_val is not None and str(right_val).strip() != "":
                        raw_kecamatan = str(right_val).strip().upper()
                        
                        # 2. Clean up the string (remove the leading colon if it exists)
                        # e.g., ": CIBINONG" becomes "CIBINONG"
                        if raw_kecamatan.startswith(":"):
                            raw_kecamatan = raw_kecamatan[1:].strip()
                            
                        # 3. Store the clean name and break the loop so we stop looking right
                        kecamatan = raw_kecamatan
                        break
            
            # ---------------------------------------------------------
            # 1.2.1a: Find district names (Kelurahan/Desa)
            # ---------------------------------------------------------
            elif "DATA PEROLEHAN SUARA PARTAI POLITIK DAN SUARA CALON" in cell_str:
                first_value_found = False
                
                # Extract every value right side of the cell
                for j in range(i + 1, len(row)):
                    right_val = row[j]
                    
                    # Skip empty cells
                    if right_val is None or str(right_val).strip() == "":
                        continue
                        
                    right_str = str(right_val).strip().upper().replace('\n', ' ')
                    
                    # Stop Condition: "JUMLAH AKHIR"
                    if "JUMLAH AKHIR" in right_str:
                        break
                        
                    # Stop Condition: "JUMLAH PINDAHAN" (only if not the first value)
                    if "JUMLAH PINDAHAN" in right_str and first_value_found:
                        break
                        
                    first_value_found = True
                    
                    # 1.2.1b & 1.2.1c: Append to kelurahan array if it's a new unique value
                    clean_kelurahan_name = str(right_val).strip().replace('\n', ' ')
                    if clean_kelurahan_name not in kelurahan and clean_kelurahan_name != "JUMLAH PINDAHAN":
                        kelurahan.append(clean_kelurahan_name)
                
                # Break inner loop to move to the next row
                break 

            # ---------------------------------------------------------
            # 1.2.2a: Find Party Name (A.1)
            # ---------------------------------------------------------
            elif cell_str == "A.1":
                for j in range(i + 1, len(row)):
                    right_val = row[j+1]
                    
                    if right_val is not None and str(right_val).strip() != "":
                        raw_party_name = right_val
                        
                        # Extract without including numeric characters (and strip stray dots/spaces)
                        clean_party_name = re.sub(r'\d+', '', raw_party_name).strip(' .')
                        
                        # 1.2.2b: Only append if it's a new unique value
                        is_unique = True
                        for idx, party_row in enumerate(suara):
                            if party_row[0] == clean_party_name:
                                is_unique = False
                                current_party_index = idx # Update current party tracker
                                break
                        
                        if is_unique:
                            # Store it to the list of list at index 0
                            suara.append([clean_party_name])
                            current_party_index = len(suara) - 1 # Update current party tracker
                            
                        break # Stop looking right once we found the party name
                
                break # Move to the next row
            
            # ---------------------------------------------------------
            # 1.2.2d: Find Total Valid Votes for Party + Candidates
            # ---------------------------------------------------------
            elif "JUMLAH SUARA SAH PARTAI POLITIK DAN CALON" in cell_str and "(A.1+A.2)" in cell_str:
                if current_party_index != -1:
                    temp_votes = []
                    
                    # Extract every value right side of it
                    for j in range(i + 1, len(row)):
                        right_val = row[j]
                        if right_val is not None and str(right_val).strip() != "":
                            # Safely typecast the extracted vote to an integer
                            try:
                                # Convert to float first to handle cases where openpyxl reads "15.0"
                                vote_int = int(float(str(right_val).strip()))
                                temp_votes.append(vote_int)
                            except ValueError:
                                # Fallback if the cell contains non-numeric text (e.g., a dash or error)
                                temp_votes.append(0)
                    
                    if temp_votes:
                        #print("Before cleaning:", temp_votes)
                        # "...except the last one"
                        temp_votes = temp_votes[:-2]
                        
                        
                        # Store it to the suara list of list based on current_party variable
                        if len(suara[current_party_index]) > 1:
                            pass
                            temp_votes = temp_votes[1:]
                        
                        #print("After cleaning:", temp_votes)
                        suara[current_party_index].extend(temp_votes)
                        
                break # Move to the next row

    return kecamatan, kelurahan, suara


def format_to_dataframe(kelurahan, suara, kecamatan_name="UNKNOWN"):
    """
    Transforms the extracted arrays into a wide pandas DataFrame.
    If there are more votes than kelurahan names, it pads kelurahan with "-".
    """
    print("kecamatan: ", kecamatan_name)
    print("suara:\n ", suara)
    # 1. Find the absolute maximum length needed
    max_len = len(kelurahan)
    for party_data in suara:
        party_votes = party_data[1:-1]
        if len(party_votes) > max_len:
            max_len = len(party_votes)
            
    # 2. Pad the geographical columns to match max_len
    # If max_len is greater than the original kelurahan list, add "-" for the missing ones
    kelurahan_padded = kelurahan + ["-"] * (max_len - len(kelurahan))
    
    # Pad the kecamatan list to match the same length
    kecamatan_padded = [kecamatan_name] * max_len
    
    data_dict = {
        'KECAMATAN': kecamatan_padded,
        'KELURAHAN': kelurahan_padded,
    }
    
    # 3. Process each party and pad their votes to match max_len
    for party_data in suara:
        party_name = party_data[0]
        party_votes = party_data[1:]
        
        current_length = len(party_votes)
        
        # If this specific party has fewer votes than the max_len, pad with 0s
        if current_length < max_len:
            missing_count = max_len - current_length
            party_votes.extend([0] * missing_count)
            
        # Add the corrected array to the dictionary
        data_dict[party_name] = party_votes

    # Create the DataFrame
    df = pd.DataFrame(data_dict)
    
    return df


def main():
    # --- Streamlit UI Setup ---
    st.set_page_config(page_title="Pemilu Parser", layout="wide")
    st.title("📊 Rekapitulasi Suara Pemilu")
    st.write("Upload file DA1 untuk menggabungkannya menjadi satu tabel.")

    # --- NEW: Drag and Drop Uploader ---
    # accept_multiple_files=True lets you select or drag in dozens of files at once
    uploaded_files = st.file_uploader(
        "Tarik dan lepas file Excel di sini (.xlsx)", 
        type=["xlsx"], 
        accept_multiple_files=True
    )
    
    # Stop the app here if no files are uploaded yet
    if not uploaded_files:
        st.info("Silakan upload file .xlsx untuk memulai pemrosesan.")
        return 
        
    st.success(f"Ditemukan {len(uploaded_files)} file. Memproses data...")
    
    all_dfs = [] 
    progress_bar = st.progress(0)
    
    # Loop through the uploaded files directly
    for index, file in enumerate(uploaded_files):
        # file.name gives us the original file name (e.g., 'data_kecamatan.xlsx')
        st.write(f"⏳ Membaca: **{file.name}**...") 
        
        # openpyxl can read the Streamlit uploaded file directly from memory!
        worksheet = load_excel_sheet(file)
        
        if worksheet:
            kecamatan, kelurahan, suara = extract_vote_by_party(worksheet)
            df = format_to_dataframe(kelurahan, suara, kecamatan)
            all_dfs.append(df)
            
        progress_bar.progress((index + 1) / len(uploaded_files))
            
    # Combine and Display
    if all_dfs:
        master_df = pd.concat(all_dfs, ignore_index=True)
        
        st.success("✅ Semua file berhasil diproses!")
        st.subheader("Data Master (Gabungan)")
        st.dataframe(master_df, use_container_width=True)
        
        # --- EXCEL DOWNLOAD BUTTON ---
        st.write("---") 
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            master_df.to_excel(writer, index=False, sheet_name='Rekap_Suara')
            
        st.download_button(
            label="📥 Download Excel File",
            data=buffer.getvalue(),
            file_name="Master_Rekap_Suara_DA1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Gagal mengekstrak data dari file mana pun.")
    

if __name__ == "__main__":
    main()