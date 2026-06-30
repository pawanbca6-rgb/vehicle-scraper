import streamlit as st
import time
import os
import re
import fitz  # PyMuPDF
import requests
import pandas as pd
import hashlib
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
import zipfile
import io

# --- ENHANCED UI CONFIGURATION ---
st.set_page_config(
    page_title="Royal Sundaram Image Tool", 
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Design Palette Injection
st.markdown("""
    <style>
    .main-header { font-size: 34px; font-weight: 800; color: #1E3A8A; margin-bottom: 2px; }
    .sub-header { font-size: 16px; color: #4B5563; margin-bottom: 25px; }
    .metric-box { background-color: #F3F4F6; padding: 18px; border-radius: 10px; border-left: 6px solid #2563EB; box-shadow: 0 2px 4px rgba(0,0,0,0.04); }
    .footer-text { text-align: center; font-size: 14px; color: #9CA3AF; margin-top: 50px; padding-top: 20px; border-top: 1px solid #E5E7EB; }
    .note-box { background-color: #FEF3C7; color: #92400E; padding: 12px; border-radius: 6px; border-left: 4px solid #F59E0B; margin-bottom: 15px; font-size: 14px; font-weight: 500; }
    </style>
""", unsafe_allow_html=True)

# --- SIDEBAR CONTROL PANEL ---
with st.sidebar:
    st.markdown("### 🎛️ Control & Template Center")
    st.write("Use this panel to verify specifications or pull down a template layout.")
    st.write("---")
    
    # Generate Sample Template Data Frame on the fly
    sample_data = {
        "Registration_No": ["OD01Z0269", "DL3CABC123"],
        "Link": ["https://icma.royalsundaram.in/...", "https://icma.royalsundaram.in/..."]
    }
    sample_df = pd.DataFrame(sample_data)
    
    # Convert to Excel Byte Stream
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        sample_df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    st.markdown("**1. Grab Base Blueprint Layout:**")
    st.download_button(
        label="📥 Download Sample Excel Template",
        data=buffer.getvalue(),
        file_name="sample_input_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.write("---")
    st.caption("Engine State: Production Cloud Cluster")

# --- MAIN APP LAYOUT ---
st.markdown("## 🚗 Royal Sundaram Image Tool")
st.markdown("##### Automated server-side bulk extraction utility for final survey insurance file elements.")

# Requirement Guidelines Box
st.markdown("""
    <div class="note-box">
        📌 <strong>Note:</strong> Please upload an Excel sheet containing <strong>Registration_No</strong> in the first column and your target portal <strong>Link</strong> in the second column.
    </div>
""", unsafe_allow_html=True)

layout_left, layout_right = st.columns([2, 1])

with layout_left:
    st.markdown("### 📂 Document Upload")
    uploaded_file = st.file_uploader("Drop target xlsx file context here", type=["xlsx"], label_visibility="collapsed")

# Core Engine Workers
def extract_claim_from_text(text):
    match = re.search(r'([A-Z]{2}\d{8})', str(text))
    return match.group(1) if match else None

def extract_and_clean_pdf(pdf_path, output_folder, pdf_base_name):
    try:
        if os.path.getsize(pdf_path) == 0:
            if os.path.exists(pdf_path): os.remove(pdf_path)
            return 0
        doc = fitz.open(pdf_path)
        img_count = 0
        for page_index in range(len(doc)):
            page = doc[page_index]
            images = page.get_images(full=True)
            for img_index, img in enumerate(images):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                
                # Get image dimensions to filter out small icons/stamps/slices
                width = base_image.get("width", 0)
                height = base_image.get("height", 0)
                
                # 🛠️ FILTER: Agar image bahut choti hai (watermark/logo/stamp/slice), toh skip karein
                if width < 150 or height < 150:
                    continue
                    
                img_hash = hashlib.md5(image_bytes).hexdigest()[:6]
                image_name = f"{pdf_base_name}_p{page_index+1}_i{img_index+1}_{img_hash}.{image_ext}"
                image_path = os.path.join(output_folder, image_name)
                if os.path.exists(image_path):
                    image_name = f"{pdf_base_name}_p{page_index+1}_i{img_index+1}_dup_{img_count}.{image_ext}"
                    image_path = os.path.join(output_folder, image_name)
                with open(image_path, "wb") as f:
                    f.write(image_bytes)
                img_count += 1
        doc.close()
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return img_count
    except Exception:
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return 0

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    if 'Registration_No' not in df.columns or 'Link' not in df.columns:
        st.error("❌ Column Validation Error: The sheet must contain explicit 'Registration_No' and 'Link' header tracks.")
    else:
        valid_rows = df[df['Registration_No'].notna() & df['Link'].notna()]
        total_vehicles_count = len(valid_rows)
        
        with layout_right:
            st.markdown("### 📊 Metrics Assessment")
            st.markdown(f"""
                <div class="metric-box">
                    <span style='font-size:13px; color:#6B7280; text-transform: uppercase; font-weight:bold;'>Validated Queue Load</span><br>
                    <span style='font-size:32px; font-weight:bold; color:#1E3A8A;'>{total_vehicles_count} Target Rows</span>
                </div>
            """, unsafe_allow_html=True)
            
        with st.sidebar:
            st.write(" ")
            run_engine = st.button("🚀 Execute Engine Scan", type="primary", use_container_width=True)
            
        if run_engine:
            CURRENT_BATCH_DIR = os.path.abspath("Downloaded_Images")
            if os.path.exists(CURRENT_BATCH_DIR):
                import shutil
                shutil.rmtree(CURRENT_BATCH_DIR)
            os.makedirs(CURRENT_BATCH_DIR, exist_ok=True)
            
            # Cloud Execution Environment Arguments
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument("--headless=new")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--remote-allow-origins=*")
            chrome_options.binary_location = "/usr/bin/chromium"
            
            try:
                chrome_service = Service("/usr/bin/chromedriver")
                driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
            except Exception as driver_err:
                st.error(f"Failed to load standard web driver hooks: {driver_err}")
                st.stop()
                
            report_data = []
            
            st.write("---")
            st.markdown("### ⚙️ Engine Execution Trace Logs")
            engine_progressbar = st.progress(0)
            
            with st.status("Spinning virtual browser pipelines...", expanded=True) as operation_context:
                for idx, (index, row) in enumerate(valid_rows.iterrows()):
                    reg_no = str(row['Registration_No']).strip()
                    portal_url = str(row['Link']).strip()
                    
                    if "http" not in portal_url:
                        continue
                        
                    operation_context.write(f"⏳ **Scraping Portfolio Array ({idx+1}/{total_vehicles_count}):** File Target `{reg_no}`")
                    
                    reg_folder = os.path.join(CURRENT_BATCH_DIR, reg_no)
                    os.makedirs(reg_folder, exist_ok=True)
                    
                    row_total_files, row_pdf_count, row_other_count = 0, 0, 0
                    row_downloaded_files, row_failed_files, row_extracted_images_total, row_direct_images_saved = 0, 0, 0, 0
                    row_attempts = 1
                    row_time_stamp = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                    
                    try:
                        driver.get(portal_url)
                        time.sleep(6)
                        
                        current_page_url = driver.current_url
                        discovered_claim_number = extract_claim_from_text(current_page_url) or extract_claim_from_text(portal_url)

                        selenium_cookies = driver.get_cookies()
                        session = requests.Session()
                        for cookie in selenium_cookies:
                            session.cookies.set(cookie['name'], cookie['value'])
                            
                        user_agent = driver.execute_script("return navigator.userAgent;")
                        session.headers.update({"User-Agent": user_agent})
                        
                        try:
                            dropdown_element = driver.find_element(By.ID, "selectedDocument")
                            dropdown = Select(dropdown_element)
                            options = [opt.text for opt in dropdown.options if opt.text.strip()]
                            row_total_files = len(options)
                            
                            if not discovered_claim_number:
                                for opt_name in options:
                                    found_claim = extract_claim_from_text(opt_name)
                                    if found_claim:
                                        discovered_claim_number = found_claim
                                        break
                            
                            if not discovered_claim_number:
                                discovered_claim_number = "CV00168307"
                            
                            for doc_name in options:
                                if doc_name.lower().endswith('.pdf'):
                                    row_pdf_count += 1
                                else:
                                    row_other_count += 1

                                detected_claim = extract_claim_from_text(doc_name) or discovered_claim_number
                                doc_name_lower = doc_name.lower()
                                temp_file_name = f"{reg_no}_{doc_name}"
                                file_save_path = os.path.join(reg_folder, temp_file_name)
                                
                                download_url = f"https://icma.royalsundaram.in/DocumentsViewer/viewdocuments.do?do=fetchDocumentInternalCall&Dataclass=EcmsClaims&DocIndex={doc_name}&proposalCode=&inwardCode=&claimNumber={detected_claim}&documentType=FINAL_SURVEY"
                                
                                file_success = False
                                for attempt in range(1, 4):
                                    row_attempts = max(row_attempts, attempt)
                                    try:
                                        response = session.get(download_url, timeout=60, stream=True)
                                        if response.status_code == 200:
                                            with open(file_save_path, "wb") as f:
                                                for chunk in response.iter_content(chunk_size=1024 * 1024):
                                                    if chunk: f.write(chunk)
                                            
                                            if os.path.exists(file_save_path) and os.path.getsize(file_save_path) > 100:
                                                if doc_name_lower.endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
                                                    row_downloaded_files += 1
                                                    row_direct_images_saved += 1
                                                    file_success = True
                                                    break
                                                else:
                                                    pdf_name_without_ext = os.path.splitext(doc_name)[0]
                                                    pdf_base_identifier = f"{reg_no}_{pdf_name_without_ext}"
                                                    imgs_saved = extract_and_clean_pdf(file_save_path, reg_folder, pdf_base_identifier)
                                                    
                                                    if imgs_saved > 0:
                                                        row_downloaded_files += 1
                                                        row_extracted_images_total += imgs_saved
                                                        file_success = True
                                                        break
                                        time.sleep(4)
                                    except Exception:
                                        time.sleep(4)
                                
                                if not file_success:
                                    row_failed_files += 1
                                    if os.path.exists(file_save_path): os.remove(file_save_path)
                                time.sleep(2)
                        except Exception:
                            row_failed_files = row_total_files
                        
                        report_data.append({
                            "Registration_No": reg_no, "Link": portal_url, "Timestamp": row_time_stamp,
                            "Total_Source_Files": row_total_files, "Total_PDF_Count": row_pdf_count,
                            "Total_Other_Files_Count": row_other_count, "Downloaded_Files_Success": row_downloaded_files,
                            "Failed_Files_Count": row_failed_files, "Total_Extracted_Images_From_PDF": row_extracted_images_total,
                            "Total_Direct_Images_Saved": row_direct_images_saved, 
                            "Total_Images_In_Folder": (row_extracted_images_total + row_direct_images_saved),
                            "Max_Attempts_Used": row_attempts
                        })
                    except Exception as e:
                        operation_context.write(f"⚠️ Trace anomaly recorded on row {reg_no}: {e}")
                        
                    engine_progressbar.progress((idx + 1) / total_vehicles_count)
                    
                driver.quit()
                operation_context.update(label="🚀 Data Compiling Finished!", state="complete", expanded=False)
                
            # --- CONSOLIDATED SINGLE DOWNLOAD ARCHIVE BUILDER ---
            if report_data:
                report_df = pd.DataFrame(report_data)
                report_csv_path = os.path.join(CURRENT_BATCH_DIR, "Execution_Report.csv")
                report_df.to_csv(report_csv_path, index=False)
                
                master_delivery_zip = "Master_Extracted_Package.zip"
                with zipfile.ZipFile(master_delivery_zip, 'w', zipfile.ZIP_DEFLATED) as master_zip:
                    for root, dirs, files in os.walk(CURRENT_BATCH_DIR):
                        for file in files:
                            file_full_path = os.path.join(root, file)
                            relative_archive_path = os.path.relpath(file_full_path, CURRENT_BATCH_DIR)
                            master_zip.write(file_full_path, relative_archive_path)
                
                st.balloons()
                st.success("🎉 Extraction processing batch verified. Everything is compiled neatly inside one master download archive:")
                
                with open(master_delivery_zip, "rb") as delivery_file:
                    st.download_button(
                        label="📥 DOWNLOAD EXTRACTED IMAGES & CSV REPORT (SINGLE CLICK)",
                        data=delivery_file,
                        file_name=f"Royal_Sundaram_Package_{datetime.now().strftime('%d-%m-%Y')}.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )

# --- AUTHOR FOOTER SECTION ---
st.markdown('<div class="footer-text">🛠️ Images Tool Engineered and Optimized by <b>Pawan Pandey</b></div>', unsafe_allow_html=True)
