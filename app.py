"""
DOCX Anonymizer - Streamlit User Interface

A professional-grade tool for anonymizing sensitive information in Word documents.
"""

import streamlit as st
from processor import (
    anonymize_docx,
    load_dictionary,
    clear_dictionary,
    create_zip_from_files,
    MAX_KEYWORDS_COUNT,
    MAX_KEYWORD_LENGTH,
    MAX_FILE_SIZE_MB,
    MAX_FILES_COUNT,
    DictionaryLockError,
    DictionarySaveError,
    FileTooLargeError,
)

# =============================================================================
# Page Configuration
# =============================================================================

st.set_page_config(
    page_title="DOCX Anonymizer",
    page_icon="🔒",
    layout="centered",
    initial_sidebar_state="expanded"
)

# =============================================================================
# Cached Dictionary Loading
# =============================================================================

@st.cache_data(ttl=5)
def get_cached_dictionary():
    """Load dictionary with caching."""
    try:
        keywords, next_num = load_dictionary()
        return keywords, next_num, None
    except DictionaryLockError as e:
        return {}, 1, str(e)

# =============================================================================
# Custom Styling
# =============================================================================

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    .stApp {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        font-family: 'Inter', sans-serif;
    }
    
    .main .block-container {
        background: rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        padding: 2rem 3rem;
        margin-top: 2rem;
    }
    
    [data-testid="stSidebar"] {
        background: rgba(15, 52, 96, 0.95);
        border-right: 1px solid rgba(255, 255, 255, 0.1);
    }
    
    h1 {
        background: linear-gradient(90deg, #e94560, #ff6b6b);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 700;
        text-align: center;
        margin-bottom: 0.5rem !important;
    }
    
    .subtitle {
        text-align: center;
        color: rgba(255, 255, 255, 0.7);
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }
    
    .stFileUploader > div > div {
        background: rgba(255, 255, 255, 0.08);
        border: 2px dashed rgba(233, 69, 96, 0.5);
        border-radius: 15px;
    }
    
    .stTextArea textarea {
        background: rgba(255, 255, 255, 0.08) !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        border-radius: 12px !important;
        color: white !important;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #e94560 0%, #ff6b6b 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        width: 100%;
    }
    
    .stDownloadButton > button {
        background: linear-gradient(135deg, #00d9ff 0%, #00a8cc 100%);
        color: white;
        border: none;
        border-radius: 12px;
        width: 100%;
    }
    
    .info-box {
        background: rgba(0, 217, 255, 0.1);
        border-left: 4px solid #00d9ff;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
        color: rgba(255, 255, 255, 0.8);
    }
    
    .dict-box {
        background: rgba(233, 69, 96, 0.1);
        border-left: 4px solid #e94560;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
        color: rgba(255, 255, 255, 0.8);
        font-size: 0.85rem;
    }
    
    .footer {
        text-align: center;
        color: rgba(255, 255, 255, 0.4);
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid rgba(255, 255, 255, 0.1);
        font-size: 0.85rem;
    }
    
    #MainMenu, footer, header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# =============================================================================
# Sidebar
# =============================================================================

with st.sidebar:
    st.markdown("## 📚 Keyword Dictionary")
    
    dictionary, _, dict_error = get_cached_dictionary()
    if dict_error:
        st.error(f"⚠️ {dict_error}")
    
    dict_count = len(dictionary)
    
    st.markdown(f"""
    <div class="dict-box">
        <strong>📖 Saved Keywords:</strong> {dict_count}
    </div>
    """, unsafe_allow_html=True)
    
    if dict_count > 0:
        with st.expander("View Dictionary"):
            for kw, replacement in dictionary.items():
                st.text(f"{kw} → {replacement}")
        
        if st.button("🗑️ Clear Dictionary", use_container_width=True):
            try:
                clear_dictionary()
                st.cache_data.clear()
                st.success("Dictionary cleared!")
                st.rerun()
            except (DictionaryLockError, DictionarySaveError) as e:
                st.error(f"Could not clear: {e}")
    
    st.markdown("---")
    st.markdown("### ⚙️ Options")
    
    include_dictionary = st.checkbox("📚 Include dictionary keywords", value=True)
    anonymize_financial = st.checkbox("💰 Anonymize financial data", value=False)
    anonymize_pii = st.checkbox("🔐 Anonymize PII data", value=False)
    
    if anonymize_pii:
        st.caption("Detects: Email, Phone, SSN, Credit Card, IP, Date")
    
    st.markdown("---")
    st.caption(f"Limits: {MAX_FILE_SIZE_MB}MB/file, {MAX_FILES_COUNT} files")

# =============================================================================
# Main Application
# =============================================================================

st.title("🔒 DOCX Anonymizer")
st.markdown('<p class="subtitle">Securely redact sensitive information from Word documents</p>', unsafe_allow_html=True)

# File Upload
uploaded_files = st.file_uploader(
    "📄 Upload Word document(s)",
    type=["docx"],
    accept_multiple_files=True,
    help=f"Max {MAX_FILES_COUNT} files, {MAX_FILE_SIZE_MB}MB each"
)

# Validate file count
if uploaded_files and len(uploaded_files) > MAX_FILES_COUNT:
    st.error(f"⚠️ Too many files. Maximum is {MAX_FILES_COUNT}.")
    uploaded_files = uploaded_files[:MAX_FILES_COUNT]

if uploaded_files:
    st.info(f"📁 {len(uploaded_files)} file(s) selected")

# Keywords Input
st.markdown("### 🔑 Additional Keywords")
st.markdown("""
<div class="info-box">
    <strong>💡 Tip:</strong> Enter keywords separated by commas. Auto-saved to dictionary.
</div>
""", unsafe_allow_html=True)

keywords_input = st.text_area(
    "Enter keywords (comma-separated)",
    placeholder="John Doe, Acme Corporation, secret@email.com",
    height=80
)

# Active options summary
active_options = []
if include_dictionary and dict_count > 0:
    active_options.append(f"{dict_count} dictionary keywords")
if anonymize_financial:
    active_options.append("financial data")
if anonymize_pii:
    active_options.append("PII data")

if active_options:
    st.info(f"ℹ️ Will anonymize: {', '.join(active_options)}")

# Process Button
st.markdown("<br>", unsafe_allow_html=True)

if st.button("🚀 Process Document(s)", use_container_width=True):
    if not uploaded_files:
        st.error("⚠️ Please upload at least one DOCX file.")
    else:
        new_keywords = [kw.strip() for kw in keywords_input.split(",") if kw.strip()] if keywords_input.strip() else []
        has_work = new_keywords or (include_dictionary and dict_count > 0) or anonymize_financial or anonymize_pii
        
        if not has_work:
            st.warning("⚠️ Enable at least one anonymization option.")
        else:
            processed_files = []
            all_stats = []
            errors = []
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, uploaded_file in enumerate(uploaded_files):
                status_text.text(f"Processing {uploaded_file.name}...")
                progress_bar.progress((i + 1) / len(uploaded_files))
                
                try:
                    anonymized_file, stats = anonymize_docx(
                        uploaded_file,
                        keywords=new_keywords if new_keywords else None,
                        include_dictionary=include_dictionary,
                        anonymize_financial=anonymize_financial,
                        anonymize_pii=anonymize_pii
                    )
                    processed_files.append((f"anonymized_{uploaded_file.name}", anonymized_file))
                    all_stats.append((uploaded_file.name, stats))
                    
                except FileTooLargeError as e:
                    errors.append((uploaded_file.name, f"File too large: {e}"))
                except ValueError as e:
                    errors.append((uploaded_file.name, f"Invalid file: {e}"))
                except DictionaryLockError as e:
                    errors.append((uploaded_file.name, f"Dictionary locked: {e}"))
                except IOError as e:
                    errors.append((uploaded_file.name, f"Read error: {e}"))
                except Exception as e:
                    errors.append((uploaded_file.name, f"Unexpected error: {type(e).__name__}"))
                    
            progress_bar.empty()
            status_text.empty()
            st.cache_data.clear()
            
            # Show errors
            for filename, error in errors:
                st.error(f"❌ {filename}: {error}")
            
            if processed_files:
                st.success(f"✅ Processed {len(processed_files)}/{len(uploaded_files)} file(s)")
                
                # Stats
                for filename, stats in all_stats:
                    with st.expander(f"📊 {filename}"):
                        cols = st.columns(3)
                        cols[0].metric("Keywords", stats.keywords_replaced)
                        cols[1].metric("Financial", stats.financial_replaced)
                        cols[2].metric("PII", sum(stats.pii_replaced.values()))
                        if stats.pii_replaced:
                            st.caption("PII: " + ", ".join(f"{k}: {v}" for k, v in stats.pii_replaced.items()))
                
                st.markdown("---")
                
                # Downloads
                if len(processed_files) == 1:
                    filename, data = processed_files[0]
                    st.download_button(
                        label=f"⬇️ Download {filename}",
                        data=data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                else:
                    zip_data = create_zip_from_files(processed_files)
                    st.download_button(
                        label="📦 Download All (ZIP)",
                        data=zip_data,
                        file_name="anonymized_documents.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    
                    with st.expander("Individual files"):
                        for filename, data in processed_files:
                            st.download_button(
                                label=f"⬇️ {filename}",
                                data=data,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                                key=filename
                            )

# Footer
st.markdown("""
<div class="footer">
    <p>Built with ❤️ using Python, Streamlit & python-docx</p>
    <p>© 2026 DOCX Anonymizer</p>
</div>
""", unsafe_allow_html=True)
