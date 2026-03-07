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
    dpt_values = []
    suara_sah_values = []
    pdip_candidates = []
    total_suara_pdip = []
    
    found_dpt_section = False
    in_pdip_section = False
    
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
            if cell_str == "B.":
                print("B. | index: ", current_party_index)
            #print(cell_str)
            
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
            # NEW: Find DPT Section and Extract "JML"
            # ---------------------------------------------------------
            elif "1. JUMLAH PEMILIH DALAM DPT" in cell_str:
                # We found the header, flip the switch ON!
                found_dpt_section = True
                
            elif found_dpt_section and cell_str == "JML" and not dpt_values:
                # We are in the DPT section AND we found the JML row
                temp_dpt = []
                
                # Extract the numbers to the right
                for j in range(i + 1, len(row)):
                    right_val = row[j]
                    if right_val is not None and str(right_val).strip() != "":
                        try:
                            dpt_int = int(float(str(right_val).strip()))
                            temp_dpt.append(dpt_int)
                        except ValueError:
                            temp_dpt.append(0)
                
                if temp_dpt:
                    # Drop the final "JUMLAH AKHIR" total
                    dpt_values = temp_dpt[:-2]
                    #print("before extend", dpt_values)
                    
                # Flip the switch OFF so we don't accidentally grab JML from other sections later
                found_dpt_section = False
                break # Move to the next row
            
            elif found_dpt_section and cell_str == "JML" and dpt_values:
                # We are in the DPT section AND we found the JML row
                temp_dpt = []
                
                # Extract the numbers to the right
                for j in range(i + 1, len(row)):
                    right_val = row[j]
                    if right_val is not None and str(right_val).strip() != "":
                        try:
                            dpt_int = int(float(str(right_val).strip()))
                            temp_dpt.append(dpt_int)
                        except ValueError:
                            temp_dpt.append(0)
                
                if temp_dpt:
                    # Drop the final "JUMLAH AKHIR" total
                    dpt_values.extend(temp_dpt[1:-2])
                    #print("after extend: ", dpt_values)
                    
                # Flip the switch OFF so we don't accidentally grab JML from other sections later
                found_dpt_section = False
                break # Move to the next row
            
            # ---------------------------------------------------------
            # NEW: Find Total Valid Votes (Suara Sah)
            # ---------------------------------------------------------
            elif "JUMLAH SELURUH SUARA SAH (IV.1.B" in cell_str and not suara_sah_values:
                temp_suara_sah = []
                
                # Extract the numbers to the right
                for j in range(i + 1, len(row)):
                    right_val = row[j]
                    if right_val is not None and str(right_val).strip() != "":
                        try:
                            sah_int = int(float(str(right_val).strip()))
                            temp_suara_sah.append(sah_int)
                        except ValueError:
                            pass
                
                if temp_suara_sah:
                    # Drop the final "JUMLAH AKHIR" total
                    suara_sah_values = temp_suara_sah[:-1] 
                    
                break # Move to the next row
            
            elif "JUMLAH SELURUH SUARA SAH (IV.1.B" in cell_str and suara_sah_values:
                temp_suara_sah = []
                
                # Extract the numbers to the right
                for j in range(i + 1, len(row)):
                    right_val = row[j]
                    if right_val is not None and str(right_val).strip() != "":
                        try:
                            sah_int = int(float(str(right_val).strip()))
                            temp_suara_sah.append(sah_int)
                        except ValueError:
                            pass
                
                if temp_suara_sah:
                    # Drop the final "JUMLAH AKHIR" total
                    suara_sah_values.extend(temp_suara_sah[1:-1]) 
                    
                break # Move to the next row
            
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
                        
                        # --- NEW: Check if this party is PDIP ---
                        if "DEMOKRASI INDONESIA PERJUANGAN" in clean_party_name.upper():
                            in_pdip_section = True
                            
                            temp_pdip_votes = []
                            
                            # Scan to the right of the party name to get the votes
                            for k in range(j + 1, len(row)):
                                v = row[k]
                                if v is not None and str(v).strip() != "":
                                    try:
                                        temp_pdip_votes.append(int(float(str(v).strip())))
                                    except ValueError:
                                        pass # Ignore empty cells or non-numbers
                                        
                            if temp_pdip_votes:
                                temp_pdip_votes = temp_pdip_votes[:-1] # Drop JUMLAH AKHIR
                                
                            if not total_suara_pdip:
                                total_suara_pdip.extend(temp_pdip_votes)
                            else:
                                total_suara_pdip.extend(temp_pdip_votes[1:])
                            
                        else:
                            in_pdip_section = False
                        
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
                            #print(clean_party_name)
                            current_party_index = len(suara) - 1 # Update current party tracker
                        
                        print("Parpol: ", clean_party_name, " | index: ", current_party_index, " | stored: ", suara[current_party_index][0])
                        break # Stop looking right once we found the party name
                
                break # Move to the next row
            
            # ---------------------------------------------------------
            # NEW: Find PDIP Candidates (A.2)
            # ---------------------------------------------------------
            elif cell_str != "B." and in_pdip_section:
                cand_name = "UNKNOWN"
                temp_votes = []
                print("current party index: ", current_party_index, " | ", cell_str)
                
                # Scan to the right to find the candidate name and votes
                for j in range(i + 1, len(row)):
                    val = row[j]
                    if val is not None and str(val).strip() != "":
                        val_str = str(val).strip()
                        
                        # Skip if it's just the candidate number (e.g., "1")
                        if val_str.isdigit() or re.match(r'^\d+\.?$', val_str):
                            continue 
                            
                        # Found the name! Clean it up
                        cand_name = re.sub(r'^\d+\s*\.?\s*', '', val_str).strip()
                        
                        # Now get the votes to the right of this name
                        for k in range(j + 1, len(row)):
                            v = row[k]
                            if v is not None and str(v).strip() != "":
                                try:
                                    temp_votes.append(int(float(str(v).strip())))
                                except ValueError:
                                    # We use pass here so we don't accidentally append 0 for empty merged cells
                                    pass 
                                    
                        break # Break the j loop since we found the name and votes
                        
                if temp_votes:
                    temp_votes = temp_votes[:-1] # Remove JUMLAH AKHIR
                    #print("candidate: ", cand_name, " | ", temp_votes)
                    
                # Append EVERY candidate as a list containing their name and votes
                candidate_found = False
                
                
                for candidate in pdip_candidates:
                    if candidate[0] == cand_name:
                        candidate.extend(temp_votes[1:])
                        candidate_found = True
                        break
                if not candidate_found:
                    print("candidate not")
                    pdip_candidates.append([cand_name] + temp_votes)
                break # Move to the next row
            
            # ---------------------------------------------------------
            # 1.2.2d: Find Total Valid Votes for Party + Candidates
            # ---------------------------------------------------------
            elif cell_str == "B." and "JUMLAH SUARA SAH PARTAI POLITIK DAN CALON" in row[i+1] and "(A.1+A.2)" in row[i+1]:
                if current_party_index != -1:
                    print("Parpol cari suara: ", suara[current_party_index][0], " | index: ", current_party_index)
                    temp_votes = []
                    
                    # Extract every value right side of it
                    for j in range(i + 2, len(row)):
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
                        
                        # print("After cleaning:", temp_votes)
                        suara[current_party_index].extend(temp_votes)
                #print(suara[current_party_index])
                break # Move to the next row
    
    #print("Semua Kandidat:\n", pdip_candidates)
    
    return kecamatan, kelurahan, dpt_values, suara_sah_values, suara, pdip_candidates, total_suara_pdip


# --- NEW: Added total_suara_pdip to the arguments ---
def format_to_dataframe(kelurahan, dpt_values, suara_sah_values, suara, pdip_candidates, total_suara_pdip, kecamatan_name="UNKNOWN"):
    # 1. Find max_len
    max_len = len(kelurahan)
    if len(dpt_values) > max_len: max_len = len(dpt_values)
    if len(suara_sah_values) > max_len: max_len = len(suara_sah_values)
    
    # Check our new PDIP party votes length
    if len(total_suara_pdip) > max_len: max_len = len(total_suara_pdip)
        
    for party_data in suara:
        if len(party_data[1:-1]) > max_len: max_len = len(party_data[1:-1])
        
    for cand_data in pdip_candidates:
        if len(cand_data[1:]) > max_len: max_len = len(cand_data[1:])
        
    # 2. Pad Geographical, DPT, and Suara Sah
    kelurahan_padded = kelurahan + ["-"] * (max_len - len(kelurahan))
    kecamatan_padded = [kecamatan_name] * max_len
    
    dpt_padded = dpt_values + [0] * (max_len - len(dpt_values)) if len(dpt_values) < max_len else dpt_values[:max_len]
    sah_padded = suara_sah_values + [0] * (max_len - len(suara_sah_values)) if len(suara_sah_values) < max_len else suara_sah_values[:max_len]
    
    # --- NEW: Pad the PDIP Party Votes ---
    pdip_partai_padded = total_suara_pdip + [0] * (max_len - len(total_suara_pdip)) if len(total_suara_pdip) < max_len else total_suara_pdip[:max_len]
    
    # --- NEW: Add 'SUARA PARTAI PDIP' to the dictionary ---
    data_dict = {
        'KECAMATAN': kecamatan_padded,
        'KELURAHAN': kelurahan_padded,
        'DPT': dpt_padded,
        'SUARA SAH': sah_padded
    }
    
    # 3. Process each party and pad their votes
    for party_data in suara:
        party_name = party_data[0]
        party_votes = party_data[1:]
        
        if len(party_votes) < max_len:
            party_votes.extend([0] * (max_len - len(party_votes)))
            
        data_dict[party_name] = party_votes
        
    data_dict["SUARA PARTAI PDIP"] = pdip_partai_padded
        
    # 4. Process PDIP Candidates and pad their votes
    for cand_data in pdip_candidates:
        cand_name = f"PDIP - {cand_data[0]}" 
        cand_votes = cand_data[1:] 
        
        if len(cand_votes) < max_len:
            cand_votes.extend([0] * (max_len - len(cand_votes)))
            
        data_dict[cand_name] = cand_votes
    
    

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
            # Unpack all 6 variables now
            kecamatan, kelurahan, dpt_values, suara_sah_values, suara, pdip_candidates, total_suara_pdip = extract_vote_by_party(worksheet)
            #print("Suara:\n", suara)
            
            # Pass all 6 variables to the formatter
            df = format_to_dataframe(kelurahan, dpt_values, suara_sah_values, suara, pdip_candidates, total_suara_pdip, kecamatan)
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