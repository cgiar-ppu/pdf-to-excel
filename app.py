import streamlit as st
import PyPDF2
import pandas as pd
import io
from openpyxl.utils.exceptions import IllegalCharacterError

st.title("PDF to Excel Extractor")

st.markdown("""
This app allows you to upload multiple machine-readable PDF files, extract their text content using PyPDF2, and download the combined content in a single Excel sheet.

**Note:** This app does not perform OCR; it assumes the PDFs contain extractable text.
""")

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
