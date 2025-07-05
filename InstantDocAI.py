# To install required packages:
# pip install streamlit PyMuPDF python-docx pandas requests

import streamlit as st
import fitz  # PyMuPDF
import docx  # python-docx
import pandas as pd
import requests
import io
import json
import datetime

# Configure page and state
st.set_page_config(page_title="Document Analysis Assistant", layout="wide")

# Initialize session state for chat history and management
if "messages" not in st.session_state:
    st.session_state.messages = []
if "document_text" not in st.session_state:
    st.session_state.document_text = None
if "document_sources" not in st.session_state:
    st.session_state.document_sources = {}
if "processed_files" not in st.session_state:
    st.session_state.processed_files = []
if "chat_histories" not in st.session_state:
    st.session_state.chat_histories = {}
if "current_chat_id" not in st.session_state:
    st.session_state.current_chat_id = "chat_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
if "analysis_running" not in st.session_state:
    st.session_state.analysis_running = False

# Set the API key directly in the session state (hidden from users)
st.session_state.api_key = "DdEAPJjPQMEah3pi78cg2aKX9QrptFnI"
st.session_state.api_provider = "Mistral AI"

# App header
st.title("📄 Document Analysis Assistant")
st.markdown("Upload documents and interact with their content using AI")

# Sidebar with file upload
with st.sidebar:
    st.header("Upload Documents")
    uploaded_files = st.file_uploader(
        "Choose files (PDF, DOCX, XLSX, TXT)",
        type=["pdf", "docx", "xlsx", "txt"],
        accept_multiple_files=True,
        help="Upload one or more documents to analyze"
    )
    if uploaded_files:
        # Check if there are new files to process
        current_filenames = {f.name for f in uploaded_files}
        processed_filenames = {f['name'] for f in st.session_state.processed_files}
        new_files = [f for f in uploaded_files if f.name not in processed_filenames]
        if new_files:
            st.success(f"✅ Uploaded {len(new_files)} new file(s)")
            # Process each new file
            for uploaded_file in new_files:
                # Document parsing based on file type
                try:
                    if uploaded_file.name.endswith('.pdf'):
                        with st.spinner(f"Processing PDF: {uploaded_file.name}"):
                            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                            text = ""
                            for page in doc:
                                text += page.get_text()
                    elif uploaded_file.name.endswith('.docx'):
                        with st.spinner(f"Processing Word document: {uploaded_file.name}"):
                            doc = docx.Document(uploaded_file)
                            text = "\n".join([para.text for para in doc.paragraphs])  # Fixed line break issue
                    elif uploaded_file.name.endswith('.xlsx'):
                        with st.spinner(f"Processing Excel file: {uploaded_file.name}"):
                            text = ""
                            try:
                                # Create a buffer for the file
                                buffer = io.BytesIO(uploaded_file.getvalue())
                                # Try multiple engines
                                excel_engines = ['openpyxl', 'xlrd']
                                success = False
                                
                                for engine in excel_engines:
                                    try:
                                        # Try to open with current engine
                                        excel_file = pd.ExcelFile(buffer, engine=engine)
                                        sheet_names = excel_file.sheet_names
                                        
                                        # Extract basic workbook metadata
                                        text = f"Excel File: {uploaded_file.name}\n"
                                        text += f"Engine: {engine}\n"
                                        text += f"Number of sheets: {len(sheet_names)}\n"
                                        text += f"Sheet names: {', '.join(sheet_names)}\n\n"
                                        
                                        # Process each sheet
                                        all_sheets = []
                                        for sheet_name in sheet_names:
                                            try:
                                                # Multiple approaches to read the sheet
                                                read_methods = [
                                                    # Method 1: Standard with header
                                                    lambda: pd.read_excel(excel_file, sheet_name=sheet_name, header=0),
                                                    # Method 2: Convert all to strings
                                                    lambda: pd.read_excel(excel_file, sheet_name=sheet_name, header=0, 
                                                                         converters={i: str for i in range(1000)}),
                                                    # Method 3: No header, everything as string
                                                    lambda: pd.read_excel(excel_file, sheet_name=sheet_name, header=None, dtype=str),
                                                    # Method 4: Raw values, no processing
                                                    lambda: pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                                                ]
                                                
                                                # Try each method until one works
                                                df = None
                                                method_used = "None"
                                                for i, method in enumerate(read_methods):
                                                    try:
                                                        df = method()
                                                        method_used = f"Method {i+1}"
                                                        break
                                                    except Exception as method_error:
                                                        continue
                                                
                                                if df is None:
                                                    # Ultra fallback method using direct openpyxl access
                                                    try:
                                                        if engine == 'openpyxl':
                                                            # Direct worksheet access
                                                            raw_wb = excel_file.book
                                                            raw_ws = raw_wb[sheet_name]
                                                            
                                                            # Get basic sheet dimensions - with safety checks
                                                            max_row = getattr(raw_ws, 'max_row', 0) or 0  # Handle None safely
                                                            max_col = getattr(raw_ws, 'max_column', 0) or 0  # Handle None safely
                                                            
                                                            # If dimensions are invalid, try to estimate them
                                                            if max_row == 0 or max_col == 0 or max_row is None or max_col is None:
                                                                # Try to scan for dimensions
                                                                try:
                                                                    # Scan for non-empty cells to determine dimensions
                                                                    max_scan_row = 100  # Max rows to scan
                                                                    max_scan_col = 50   # Max columns to scan
                                                                    found_row = 0
                                                                    found_col = 0
                                                                    
                                                                    for r in range(1, max_scan_row + 1):
                                                                        for c in range(1, max_scan_col + 1):
                                                                            try:
                                                                                cell = raw_ws.cell(row=r, column=c)
                                                                                if cell and cell.value:
                                                                                    found_row = max(found_row, r)
                                                                                    found_col = max(found_col, c)
                                                                            except:
                                                                                continue
                                                                    
                                                                    max_row = found_row
                                                                    max_col = found_col
                                                                except:
                                                                    pass
                                                                    
                                                            # If we still have no dimensions, report and skip
                                                            if max_row == 0 or max_col == 0 or max_row is None or max_col is None:
                                                                all_sheets.append(f"--- Sheet: {sheet_name} --- [Empty sheet or could not determine dimensions]")
                                                                continue
                                                                
                                                            # Manual cell-by-cell extraction with robust error handling
                                                            sheet_data = "--- Sheet: {0} [{1} rows × {2} columns] (Manual cell extraction) ---\n".format(
                                                                sheet_name, max_row, max_col)
                                                                
                                                            # Limit large sheets
                                                            max_extract_rows = min(max_row, 200) if max_row is not None else 200
                                                            max_extract_cols = min(max_col, 30) if max_col is not None else 30
                                                            
                                                            # Build a text table representation
                                                            for r in range(1, max_extract_rows + 1):
                                                                row_data = []
                                                                for c in range(1, max_extract_cols + 1):
                                                                    try:
                                                                        cell = raw_ws.cell(row=r, column=c)
                                                                        cell_value = cell.value if cell and cell.value is not None else ""
                                                                        row_data.append(str(cell_value))
                                                                    except:
                                                                        row_data.append("")
                                                                sheet_data += " | ".join(row_data) + "\n"
                                                                
                                                            if max_row is not None and max_row > 200:
                                                                sheet_data += f"[Note: Large sheet - showing only first 200 rows of {max_row} total]\n"
                                                                
                                                            all_sheets.append(sheet_data)
                                                            continue
                                                        else:
                                                            all_sheets.append(f"--- Sheet: {sheet_name} --- [Failed to read with all methods]")
                                                            continue
                                                    except Exception as ultra_fallback_error:
                                                        all_sheets.append(f"--- Sheet: {sheet_name} --- [Failed to read: {str(ultra_fallback_error)}]")
                                                        continue
                                            except Exception as sheet_error:
                                                all_sheets.append(f"--- Sheet: {sheet_name} --- [Error: {str(sheet_error)}]")
                                        
                                        text += "\n\n".join(all_sheets)
                                        success = True
                                        break  # Exit engine loop if successful
                                        
                                    except Exception as engine_error:
                                        # Try next engine
                                        continue
                                
                                if not success:
                                    text = f"Failed to process Excel file {uploaded_file.name} with all available engines. The file may be corrupted or in an unsupported format."
                                    
                            except Exception as excel_error:
                                text = f"Error processing Excel file: {str(excel_error)}"
                    elif uploaded_file.name.endswith('.txt'):
                        with st.spinner(f"Processing text file: {uploaded_file.name}"):
                            text = uploaded_file.getvalue().decode("utf-8")
                    # Store the processed file info
                    file_info = {
                        'name': uploaded_file.name,
                        'text': text,
                        'size': len(text),
                        'timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    st.session_state.processed_files.append(file_info)
                except Exception as e:
                    st.error(f"Error processing {uploaded_file.name}: {str(e)}")
        # Update document text by combining all processed files
        combined_text = ""
        for idx, file_info in enumerate(st.session_state.processed_files):
            combined_text += f"\n--- DOCUMENT {idx+1}: {file_info['name']} ---\n"
            combined_text += file_info['text']
            # Store document sections for potential future use
            start_pos = len(combined_text) - len(file_info['text'])
            st.session_state.document_sources[file_info['name']] = {
                'start': start_pos,
                'end': len(combined_text)
            }
        st.session_state.document_text = combined_text
        # Display information about processed files
        st.subheader("Processed Documents")
        for idx, file_info in enumerate(st.session_state.processed_files):
            with st.expander(f"{idx+1}. {file_info['name']}"):
                st.text(f"Size: {file_info['size']} characters")
                st.text(f"Processed: {file_info['timestamp']}")
                st.text_area("Preview", file_info['text'][:500] + "..." if len(file_info['text']) > 500 else file_info['text'], height=100)
        # Remove files button
        if st.button("Clear All Documents"):
            st.session_state.processed_files = []
            st.session_state.document_text = None
            st.session_state.document_sources = {}
            st.rerun()

    # Analysis options
    st.header("Analysis Options")
    analysis_type = st.selectbox(
        "Select analysis type:",
        ["General Analysis", "Summarize", "Bullet Points", "Simplify", "Extract Key Insights"]
    )
    
    # Stop Analysis button
    if st.session_state.analysis_running:
        if st.button("⛔ Stop Analysis", type="primary"):
            st.session_state.analysis_running = False
            st.info("Analysis stopped by user.")
            st.rerun()

    # Chat management section
    st.header("Chat Management")
    # New chat button
    if st.button("Start New Chat"):
        # Save current chat if it has messages
        if st.session_state.messages and st.session_state.current_chat_id:
            st.session_state.chat_histories[st.session_state.current_chat_id] = {
                "messages": st.session_state.messages.copy(),
                "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "title": f"Chat {len(st.session_state.chat_histories) + 1}"
            }
        # Create a new chat
        st.session_state.current_chat_id = "chat_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        st.session_state.messages = []
        st.rerun()

    # Clear chat history button
    if st.button("Clear Chat History"):
        st.session_state.messages = []
        st.rerun()

    # Show past chats if they exist
    if st.session_state.chat_histories:
        st.subheader("Past Conversations")
        selected_chat = st.selectbox(
            "Select a past conversation",
            options=list(st.session_state.chat_histories.keys()),
            format_func=lambda x: f"{st.session_state.chat_histories[x]['title']} ({st.session_state.chat_histories[x]['timestamp']})"
        )
        if st.button("Load Selected Chat"):
            # Save current chat first
            if st.session_state.messages and st.session_state.current_chat_id:
                st.session_state.chat_histories[st.session_state.current_chat_id] = {
                    "messages": st.session_state.messages.copy(),
                    "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "title": f"Chat {len(st.session_state.chat_histories)}"
                }
            # Load selected chat
            st.session_state.messages = st.session_state.chat_histories[selected_chat]["messages"].copy()
            st.session_state.current_chat_id = selected_chat
            st.rerun()

# Main chat area
st.header("Chat with your Documents")

# Display document count
if st.session_state.processed_files:
    st.info(f"Currently analyzing {len(st.session_state.processed_files)} document(s) with a total of {len(st.session_state.document_text)} characters")

# Display chat title
if st.session_state.current_chat_id in st.session_state.chat_histories:
    chat_title = st.session_state.chat_histories[st.session_state.current_chat_id]["title"]
    st.subheader(f"Current: {chat_title}")

# Display chat messages
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Update the API function to use Mistral AI
def call_ai_api(prompt, document_text, analysis_type):
    # Get API key from session state
    api_key = st.session_state.get("api_key", "")
    # Create system message based on analysis type
    system_messages = {
        "General Analysis": "You are a helpful document analysis assistant specialized in understanding structured data from various document formats including Excel files. Analyze the documents and respond to the user's query with specific information from the documents. When Excel files are presented, focus on analyzing the tabular data and providing insights based on the column values and structure. Never claim you can't access the file - all relevant content has been extracted and provided to you in plain text format.",
        "Summarize": "You are a document summarization expert. Provide a concise summary of the documents, including any structured data they contain such as Excel spreadsheets. When Excel data is included, summarize the data found in each sheet, focusing on column headers and the types of information present.",
        "Bullet Points": "You are a document structuring expert. Convert the key points of the documents into bullet points. For Excel data, create bullet points for each sheet, highlighting the key columns and data patterns.",
        "Simplify": "You are a simplification expert. Rewrite the document content in simpler, more accessible language, including explanations of any structured data or Excel content.",
        "Extract Key Insights": "You are a data insights expert. Extract and explain the most important insights from these documents, especially focusing on patterns in any tabular data from Excel files."
    }
    system_message = system_messages.get(analysis_type, system_messages["General Analysis"])
    # Truncate document if too long
    max_doc_length = 8000
    doc_text = document_text[:max_doc_length]
    # Mistral AI API configuration
    api_url = "https://api.mistral.ai/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "mistral-medium",  # You can change to other Mistral models as needed
        "messages": [
            {"role": "system", "content": f"{system_message}"},
            {"role": "user", "content": f"DOCUMENT CONTENT:\n{doc_text}\nQUERY: {prompt}"}
        ],
        "temperature": 0.7
    }
    try:
        response = requests.post(api_url, headers=headers, json=payload)
        response.raise_for_status()
        result = response.json()
        return result["choices"][0]["message"]["content"]
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 402 or e.response.status_code == 429:
            return f"Error: API rate limit or payment required. Status code: {e.response.status_code}. Your API key may be valid, but your account doesn't have sufficient credits or has reached its rate limit."
        else:
            return f"Error calling Mistral AI API: {str(e)}"
    except Exception as e:
        return f"Error: {str(e)}"

# Handle user input
if st.session_state.document_text:
    # Get user query
    if prompt := st.chat_input("Ask something about your documents..."):
        # Add user message to chat history
        st.session_state.messages.append({"role": "user", "content": prompt})
        # Display in chat
        with st.chat_message("user"):
            st.markdown(prompt)
        # Get AI response
        with st.spinner("Analyzing..."):
            # Set analysis running flag
            st.session_state.analysis_running = True
            # Add stop button in the main area too for visibility during analysis
            stop_col1, stop_col2, stop_col3 = st.columns([1, 1, 1])
            with stop_col2:
                if st.button("⛔ Stop Current Analysis", key="stop_main"):
                    st.session_state.analysis_running = False
                    st.info("Analysis stopped by user.")
                    st.rerun()
            
            # Check if analysis was stopped before proceeding
            if st.session_state.analysis_running:
                response = call_ai_api(prompt, st.session_state.document_text, analysis_type)
                # Add response to chat history
                st.session_state.messages.append({"role": "assistant", "content": response})
                # Display in chat
                with st.chat_message("assistant"):
                    st.markdown(response)
                # Auto-save current chat to histories
                st.session_state.chat_histories[st.session_state.current_chat_id] = {
                    "messages": st.session_state.messages.copy(),
                    "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "title": f"Chat {len(st.session_state.chat_histories)}"
                }
            else:
                # If analysis was stopped, add a message indicating that
                stop_message = "Analysis was stopped by the user."
                st.session_state.messages.append({"role": "assistant", "content": stop_message})
                with st.chat_message("assistant"):
                    st.warning(stop_message)
            
            # Reset the analysis flag
            st.session_state.analysis_running = False
else:
    st.info("👋 Please upload at least one document to start the analysis")

# Footer
st.markdown("---")
st.caption("Document Analysis Assistant powered by AI")
