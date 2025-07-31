import streamlit as st
import PyPDF2
import pandas as pd
import io
from openpyxl.utils.exceptions import IllegalCharacterError
import requests
import re

st.title("PDF to Excel Extractor")

st.markdown("""
This app allows you to upload multiple machine-readable PDF files or an Excel file with PDF URLs, extract their text content using PyPDF2, and download the combined content in a single Excel sheet.

**Note:** This app does not perform OCR; it assumes the PDFs contain extractable text.
""")

mode = st.radio("Choose input method", ["Upload PDF files", "Upload Excel with PDF URLs", "Paste list of URLs"])

if mode == "Upload PDF files":
    uploaded_files = st.file_uploader("Upload your PDF files", type=["pdf"], accept_multiple_files=True)

    if st.button("Extract and Download Excel"):
        if uploaded_files:
            with st.spinner("Extracting text from PDFs..."):
                errors = []
                data = []
                for uploaded_file in uploaded_files:
                    try:
                        reader = PyPDF2.PdfReader(uploaded_file)
                        text = ""
                        for page in reader.pages:
                            extracted = page.extract_text()
                            if extracted:
                                text += extracted + "\n\n"
                        full_text = text.strip()
                        
                        max_chars = 30000
                        chunks = [full_text[i:i + max_chars] for i in range(0, len(full_text), max_chars)]
                        
                        for idx, chunk in enumerate(chunks, 1):
                            data.append({"Filename": uploaded_file.name, "Part": idx, "Content": chunk})
                    except Exception as e:
                        errors.append(f"Error extracting from {uploaded_file.name}: {str(e)}")
                        continue
                
                if data:
                    df = pd.DataFrame(data)
                    good_dfs = []
                    for filename, group in df.groupby('Filename'):
                        try:
                            temp_output = io.BytesIO()
                            with pd.ExcelWriter(temp_output, engine='openpyxl') as writer:
                                group.to_excel(writer, index=False)
                            good_dfs.append(group)
                        except IllegalCharacterError as e:
                            errors.append(f"Error writing content for {filename} due to illegal characters: {str(e)}")
                    
                    if good_dfs:
                        final_df = pd.concat(good_dfs)
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            final_df.to_excel(writer, sheet_name='PDF Contents', index=False)
                        output.seek(0)
                        
                        st.download_button(
                            label="Download Excel file",
                            data=output,
                            file_name="pdf_contents.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success("Extraction complete! Click the button above to download your Excel file.")
                    else:
                        st.warning("No content could be processed successfully.")
                else:
                    st.warning("No content extracted from the uploaded files.")
                
                if errors:
                    with st.expander("Error Details"):
                        for err in errors:
                            st.error(err)
        else:
            st.warning("Please upload at least one PDF file.")

elif mode == "Upload Excel with PDF URLs":
    uploaded_excel = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

    if uploaded_excel:
        excel_file = pd.ExcelFile(uploaded_excel)
        sheet_names = excel_file.sheet_names
        selected_sheet = st.selectbox("Select the sheet", sheet_names + ["All sheets"])
        if selected_sheet != "All sheets":
            df_input = pd.read_excel(uploaded_excel, sheet_name=selected_sheet)
            columns = df_input.columns.tolist()
        else:
            first_sheet = sheet_names[0]
            df_first = pd.read_excel(uploaded_excel, sheet_name=first_sheet)
            columns = df_first.columns.tolist()
        url_column = st.selectbox("Select the column containing PDF URLs", columns)

        if st.button("Extract from URLs and Download Excel"):
            with st.spinner("Preparing to process PDFs..."):
                if selected_sheet != "All sheets":
                    sheet_list = [selected_sheet]
                else:
                    sheet_list = sheet_names
                url_list = []
                for sheet in sheet_list:
                    df_sheet = pd.read_excel(uploaded_excel, sheet_name=sheet)
                    if url_column in df_sheet.columns:
                        sheet_urls = df_sheet[url_column].dropna().tolist()
                        url_list.extend(sheet_urls)
                    else:
                        st.warning(f"Column '{url_column}' not found in sheet '{sheet}'. Skipping.")
                total_urls = len(url_list)
                if total_urls == 0:
                    st.warning("No valid URLs found to process.")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    with st.spinner("Downloading and extracting text from PDF URLs..."):
                        errors = []
                        data = []
                        for idx, url in enumerate(url_list, 1):
                            status_text.text(f"Processing {idx} of {total_urls}: {url}")
                            success = False
                            for attempt in range(2):
                                try:
                                    response = requests.get(url, timeout=10)
                                    response.raise_for_status()
                                    content_type = response.headers.get('Content-Type', '')
                                    if 'application/pdf' in content_type:
                                        pdf_stream = io.BytesIO(response.content)
                                        reader = PyPDF2.PdfReader(pdf_stream)
                                        text = ""
                                        for page in reader.pages:
                                            extracted = page.extract_text()
                                            if extracted:
                                                text += extracted + "\n\n"
                                        full_text = text.strip()
                                        max_chars = 30000
                                        chunks = [full_text[i:i + max_chars] for i in range(0, len(full_text), max_chars)]
                                        for part_idx, chunk in enumerate(chunks, 1):
                                            data.append({"URL": url, "Part": part_idx, "Content": chunk})
                                        success = True
                                        break
                                    else:
                                        raise ValueError(f"URL {url} did not return a PDF (Content-Type: {content_type})")
                                except Exception as e:
                                    if attempt == 1:
                                        errors.append(f"Error processing URL {url} after 2 attempts: {str(e)}")
                            if success:
                                progress_bar.progress(idx / total_urls)
                        status_text.text("Processing complete.")
                
                if data:
                    df = pd.DataFrame(data)
                    good_dfs = []
                    for url_group, group in df.groupby('URL'):
                        try:
                            temp_output = io.BytesIO()
                            with pd.ExcelWriter(temp_output, engine='openpyxl') as writer:
                                group.to_excel(writer, index=False)
                            good_dfs.append(group)
                        except IllegalCharacterError as e:
                            errors.append(f"Error writing content for URL {url_group} due to illegal characters: {str(e)}")
                    
                    if good_dfs:
                        final_df = pd.concat(good_dfs)
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            final_df.to_excel(writer, sheet_name='PDF Contents', index=False)
                        output.seek(0)
                        
                        st.download_button(
                            label="Download Excel file",
                            data=output,
                            file_name="pdf_contents.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success("Extraction complete! Click the button above to download your Excel file.")
                    else:
                        st.warning("No content could be processed successfully.")
                else:
                    st.warning("No content extracted from the URLs.")
                
                if errors:
                    with st.expander("Error Details"):
                        for err in errors:
                            st.error(err)
        else:
            st.warning("Please select a column and click the button to proceed.")
    else:
        st.info("Please upload an Excel file to proceed.")

else:  # Paste list of URLs
    pasted_text = st.text_area("Paste your list of PDF URLs here")
    if st.button("Extract from pasted URLs and Download Excel"):
        with st.spinner("Preparing to process pasted URLs..."):
            url_list = re.findall(r'https?://\S+', pasted_text)
            total_urls = len(url_list)
            if total_urls == 0:
                st.warning("No valid URLs found in the pasted text.")
            else:
                progress_bar = st.progress(0)
                status_text = st.empty()
                with st.spinner("Downloading and extracting text from PDF URLs..."):
                    errors = []
                    data = []
                    for idx, url in enumerate(url_list, 1):
                        status_text.text(f"Processing {idx} of {total_urls}: {url}")
                        success = False
                        for attempt in range(2):
                            try:
                                response = requests.get(url, timeout=10)
                                response.raise_for_status()
                                content_type = response.headers.get('Content-Type', '')
                                if 'application/pdf' in content_type:
                                    pdf_stream = io.BytesIO(response.content)
                                    reader = PyPDF2.PdfReader(pdf_stream)
                                    text = ""
                                    for page in reader.pages:
                                        extracted = page.extract_text()
                                        if extracted:
                                            text += extracted + "\n\n"
                                    full_text = text.strip()
                                    max_chars = 30000
                                    chunks = [full_text[i:i + max_chars] for i in range(0, len(full_text), max_chars)]
                                    for part_idx, chunk in enumerate(chunks, 1):
                                        data.append({"URL": url, "Part": part_idx, "Content": chunk})
                                    success = True
                                    break
                                else:
                                    raise ValueError(f"URL {url} did not return a PDF (Content-Type: {content_type})")
                            except Exception as e:
                                if attempt == 1:
                                    errors.append(f"Error processing URL {url} after 2 attempts: {str(e)}")
                        if success:
                            progress_bar.progress(idx / total_urls)
                    status_text.text("Processing complete.")
                    if data:
                        df = pd.DataFrame(data)
                        good_dfs = []
                        for url_group, group in df.groupby('URL'):
                            try:
                                temp_output = io.BytesIO()
                                with pd.ExcelWriter(temp_output, engine='openpyxl') as writer:
                                    group.to_excel(writer, index=False)
                                good_dfs.append(group)
                            except IllegalCharacterError as e:
                                errors.append(f"Error writing content for URL {url_group} due to illegal characters: {str(e)}")
                        if good_dfs:
                            final_df = pd.concat(good_dfs)
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                final_df.to_excel(writer, sheet_name='PDF Contents', index=False)
                            output.seek(0)
                            st.download_button(
                                label="Download Excel file",
                                data=output,
                                file_name="pdf_contents.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            st.success("Extraction complete! Click the button above to download your Excel file.")
                        else:
                            st.warning("No content could be processed successfully.")
                    else:
                        st.warning("No content extracted from the URLs.")
                    if errors:
                        with st.expander("Error Details"):
                            for err in errors:
                                st.error(err)
