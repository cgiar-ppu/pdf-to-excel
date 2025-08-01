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

        # Initialize session state variables
        if 'url_list' not in st.session_state:
            st.session_state.url_list = None
        if 'current_idx' not in st.session_state:
            st.session_state.current_idx = 0
        if 'data' not in st.session_state:
            st.session_state.data = []
        if 'errors' not in st.session_state:
            st.session_state.errors = []
        if 'total_urls' not in st.session_state:
            st.session_state.total_urls = 0
        if 'excel_data' not in st.session_state:
            st.session_state.excel_data = None
        if 'processing_mode' not in st.session_state:
            st.session_state.processing_mode = None

        progress_bar = st.progress(0)
        status_text = st.empty()

        if st.button("Extract from URLs and Download Excel"):
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
                st.session_state.url_list = url_list
                st.session_state.current_idx = 0
                st.session_state.data = []
                st.session_state.errors = []
                st.session_state.total_urls = total_urls
                st.session_state.processing_mode = "excel"
                st.session_state.excel_data = None
                st.rerun()

        # Update progress and status
        if st.session_state.processing_mode == "excel" and st.session_state.total_urls > 0:
            current_progress = min(st.session_state.current_idx / st.session_state.total_urls, 1.0)
            progress_bar.progress(current_progress)
            if st.session_state.current_idx < st.session_state.total_urls:
                status_text.text(f"Processing {st.session_state.current_idx + 1} of {st.session_state.total_urls}")
            else:
                status_text.text("Processing complete.")

        # Process one URL if needed
        if st.session_state.processing_mode == "excel" and st.session_state.url_list is not None and st.session_state.current_idx < st.session_state.total_urls:
            idx = st.session_state.current_idx
            url = st.session_state.url_list[idx]
            status_text.text(f"Processing {idx + 1} of {st.session_state.total_urls}: {url}")
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
                            st.session_state.data.append({"URL": url, "Part": part_idx, "Content": chunk})
                        success = True
                        break
                    else:
                        raise ValueError(f"URL {url} did not return a PDF (Content-Type: {content_type})")
                except Exception as e:
                    if attempt == 1:
                        st.session_state.errors.append(f"Error processing URL {url} after 2 attempts: {str(e)}")
            st.session_state.current_idx += 1
            st.rerun()

        # When processing is complete, generate Excel
        if st.session_state.processing_mode == "excel" and st.session_state.current_idx >= st.session_state.total_urls and st.session_state.excel_data is None and st.session_state.data:
            with st.spinner("Generating Excel file..."):
                df = pd.DataFrame(st.session_state.data)
                good_dfs = []
                for url_group, group in df.groupby('URL'):
                    try:
                        temp_output = io.BytesIO()
                        with pd.ExcelWriter(temp_output, engine='openpyxl') as writer:
                            group.to_excel(writer, index=False)
                        good_dfs.append(group)
                    except IllegalCharacterError as e:
                        st.session_state.errors.append(f"Error writing content for URL {url_group} due to illegal characters: {str(e)}")
                if good_dfs:
                    final_df = pd.concat(good_dfs)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, sheet_name='PDF Contents', index=False)
                    output.seek(0)
                    st.session_state.excel_data = output.getvalue()
                else:
                    st.warning("No content could be processed successfully.")

        # Show download button if excel_data is available
        if st.session_state.excel_data is not None:
            st.download_button(
                label="Download Excel file",
                data=st.session_state.excel_data,
                file_name="pdf_contents.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Extraction complete! Click the button above to download your Excel file.")

        # Show errors if any
        if st.session_state.errors:
            with st.expander("Error Details"):
                for err in st.session_state.errors:
                    st.error(err)

    else:
        st.info("Please upload an Excel file to proceed.")

else:  # Paste list of URLs
    pasted_text = st.text_area("Paste your list of PDF URLs here")

    # Initialize session state variables (using different keys for this mode)
    if 'paste_url_list' not in st.session_state:
        st.session_state.paste_url_list = None
    if 'paste_current_idx' not in st.session_state:
        st.session_state.paste_current_idx = 0
    if 'paste_data' not in st.session_state:
        st.session_state.paste_data = []
    if 'paste_errors' not in st.session_state:
        st.session_state.paste_errors = []
    if 'paste_total_urls' not in st.session_state:
        st.session_state.paste_total_urls = 0
    if 'paste_excel_data' not in st.session_state:
        st.session_state.paste_excel_data = None
    if 'processing_mode' not in st.session_state:
        st.session_state.processing_mode = None

    progress_bar = st.progress(0)
    status_text = st.empty()

    if st.button("Extract from pasted URLs and Download Excel"):
        url_list = re.findall(r'https?://\S+', pasted_text)
        total_urls = len(url_list)
        if total_urls == 0:
            st.warning("No valid URLs found in the pasted text.")
        else:
            st.session_state.paste_url_list = url_list
            st.session_state.paste_current_idx = 0
            st.session_state.paste_data = []
            st.session_state.paste_errors = []
            st.session_state.paste_total_urls = total_urls
            st.session_state.processing_mode = "paste"
            st.session_state.paste_excel_data = None
            st.rerun()

    # Update progress and status
    if st.session_state.processing_mode == "paste" and st.session_state.paste_total_urls > 0:
        current_progress = min(st.session_state.paste_current_idx / st.session_state.paste_total_urls, 1.0)
        progress_bar.progress(current_progress)
        if st.session_state.paste_current_idx < st.session_state.paste_total_urls:
            status_text.text(f"Processing {st.session_state.paste_current_idx + 1} of {st.session_state.paste_total_urls}")
        else:
            status_text.text("Processing complete.")

    # Process one URL if needed
    if st.session_state.processing_mode == "paste" and st.session_state.paste_url_list is not None and st.session_state.paste_current_idx < st.session_state.paste_total_urls:
        idx = st.session_state.paste_current_idx
        url = st.session_state.paste_url_list[idx]
        status_text.text(f"Processing {idx + 1} of {st.session_state.paste_total_urls}: {url}")
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
                        st.session_state.paste_data.append({"URL": url, "Part": part_idx, "Content": chunk})
                    success = True
                    break
                else:
                    raise ValueError(f"URL {url} did not return a PDF (Content-Type: {content_type})")
            except Exception as e:
                if attempt == 1:
                    st.session_state.paste_errors.append(f"Error processing URL {url} after 2 attempts: {str(e)}")
        st.session_state.paste_current_idx += 1
        st.rerun()

    # When processing is complete, generate Excel
    if st.session_state.processing_mode == "paste" and st.session_state.paste_current_idx >= st.session_state.paste_total_urls and st.session_state.paste_excel_data is None and st.session_state.paste_data:
        with st.spinner("Generating Excel file..."):
            df = pd.DataFrame(st.session_state.paste_data)
            good_dfs = []
            for url_group, group in df.groupby('URL'):
                try:
                    temp_output = io.BytesIO()
                    with pd.ExcelWriter(temp_output, engine='openpyxl') as writer:
                        group.to_excel(writer, index=False)
                    good_dfs.append(group)
                except IllegalCharacterError as e:
                    st.session_state.paste_errors.append(f"Error writing content for URL {url_group} due to illegal characters: {str(e)}")
            if good_dfs:
                final_df = pd.concat(good_dfs)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, sheet_name='PDF Contents', index=False)
                output.seek(0)
                st.session_state.paste_excel_data = output.getvalue()
            else:
                st.warning("No content could be processed successfully.")

    # Show download button if excel_data is available
    if st.session_state.paste_excel_data is not None:
        st.download_button(
            label="Download Excel file",
            data=st.session_state.paste_excel_data,
            file_name="pdf_contents.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Extraction complete! Click the button above to download your Excel file.")

    # Show errors if any
    if st.session_state.paste_errors:
        with st.expander("Error Details"):
            for err in st.session_state.paste_errors:
                st.error(err)
