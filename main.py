import streamlit as st
import plotly.express as px
from groq import Groq
from dotenv import load_dotenv
import services
import os
import pandas as pd
import json
from PIL import Image
import PyPDF2
import docx
import time
import tempfile
import shutil

load_dotenv()

SUPPORTED_FILES = {
    'data': ['csv', 'xlsx', 'xls', 'json'],
    'document': ['txt', 'pdf', 'docx', 'doc'],
    'image': ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp']
}

CONVERT_OPTIONS = {
    'data': ['PDF', 'CSV', 'Excel', 'JSON'],
    'document': ['PDF', 'TXT'],
    'image': ['PDF'],
    'all': ['PDF']
}

if "messages" not in st.session_state:
    st.session_state.messages = []
if "df" not in st.session_state:
    st.session_state.df = None
if "file_path" not in st.session_state:
    st.session_state.file_path = None
if "file_type" not in st.session_state:
    st.session_state.file_type = None
if "file_content" not in st.session_state:
    st.session_state.file_content = None
if "last_user_input" not in st.session_state:
    st.session_state.last_user_input = ""
if "file_name" not in st.session_state:
    st.session_state.file_name = None
if "uploaded_file_id" not in st.session_state:
    st.session_state.uploaded_file_id = None

try:
    client = Groq(api_key=os.environ.get("GROQ_API_KEY"))
except:
    client = None

st.set_page_config(page_title="Universal Fayl Tahlilchi", page_icon="📊", layout="wide")
st.title("📊 Universal Fayl Tahlil va Konvertor Dasturi")


def safe_file_write(file_path, file_content, max_retries=3):
    for attempt in range(max_retries):
        try:
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    time.sleep(0.1)
                except:
                    pass

            with open(file_path, "wb") as f:
                f.write(file_content)

            time.sleep(0.2)
            return True

        except PermissionError:
            if attempt < max_retries - 1:
                time.sleep(0.5)
                continue
            else:
                temp_path = file_path + f".temp_{int(time.time())}"
                with open(temp_path, "wb") as f:
                    f.write(file_content)
                return temp_path
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(0.5)
                continue
            else:
                raise e
    return False


def get_file_category(filename):
    ext = filename.split('.')[-1].lower()
    for category, extensions in SUPPORTED_FILES.items():
        if ext in extensions:
            return category, ext
    return None, ext


def read_file_content(file_path, file_ext):
    max_retries = 3
    for attempt in range(max_retries):
        try:
            if file_ext in ['csv']:
                return pd.read_csv(file_path)
            elif file_ext in ['xlsx', 'xls']:
                return pd.read_excel(file_path, engine='openpyxl')
            elif file_ext == 'json':
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        return pd.DataFrame(data)
                    return data
            elif file_ext == 'txt':
                with open(file_path, 'r', encoding='utf-8') as f:
                    return f.read()
            elif file_ext == 'pdf':
                with open(file_path, 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    text = ""
                    for page in pdf_reader.pages:
                        text += page.extract_text()
                    return text
            elif file_ext in ['docx', 'doc']:
                doc = docx.Document(file_path)
                return '\n'.join([para.text for para in doc.paragraphs])
            elif file_ext in ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp']:
                return Image.open(file_path)
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(0.3)
                continue
            else:
                st.error(f"Faylni o'qishda xatolik: {str(e)}")
                return None


def send_message(user_input, context=""):
    if client is None:
        return "GROQ API key topilmadi. .env faylida GROQ_API_KEY ni sozlang."

    try:
        full_prompt = f"{context}\n\nFoydalanuvchi savoli: {user_input}" if context else user_input
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": full_prompt}],
            model="llama-3.3-70b-versatile",
            temperature=0.7,
        )
        return chat_completion.choices[0].message.content
    except Exception as e:
        return f"Xatolik yuz berdi: {str(e)}"


def display_data_visualizations(df):
    numeric_df = df.select_dtypes(include=["number"])

    if numeric_df.empty:
        st.warning("📊 Grafik chizish uchun raqamli ustunlar topilmadi.")
        return

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 📈 Asosiy Statistika")
        st.dataframe(numeric_df.describe().T, use_container_width=True)

    with col2:
        st.markdown("### 📊 O'rtacha Qiymatlar")
        avg_values = numeric_df.mean().sort_values(ascending=False)
        fig_bar = px.bar(
            x=avg_values.index,
            y=avg_values.values,
            labels={"x": "Ustun", "y": "O'rtacha"},
            title="Ustunlarning O'rtacha Qiymatlari"
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("### 📉 Line Chart")
    fig_line = px.line(numeric_df, title="Ma'lumotlar Dinamikasi")
    st.plotly_chart(fig_line, use_container_width=True)

    if len(numeric_df.columns) > 1:
        st.markdown("### 🔗 Korrelyatsiya Matritsasi")
        corr_matrix = numeric_df.corr()
        fig_heatmap = px.imshow(
            corr_matrix,
            text_auto=True,
            aspect="auto",
            title="Ustunlar O'rtasidagi Bog'liqlik"
        )
        st.plotly_chart(fig_heatmap, use_container_width=True)

    with st.expander("📊 Batafsil Histogrammalar"):
        cols = st.columns(2)
        for idx, col in enumerate(numeric_df.columns):
            with cols[idx % 2]:
                fig_hist = px.histogram(
                    numeric_df, x=col, nbins=30,
                    title=f"{col} Taqsimoti"
                )
                st.plotly_chart(fig_hist, use_container_width=True)


def convert_file(file_name, format_type):
    try:
        if format_type == "PDF":
            return services.to_pdf(file_name)
        elif format_type == "CSV":
            return services.to_csv(file_name)
        elif format_type == "Excel":
            return services.to_excel(file_name)
        elif format_type == "JSON":
            return services.to_json(file_name)
        elif format_type == "TXT":
            return services.to_txt(file_name)
        else:
            raise ValueError(f"Noma'lum format: {format_type}")
    except Exception as e:
        raise Exception(f"Konvertatsiya xatosi: {str(e)}")


st.sidebar.title("🛠️ Sozlamalar")

all_extensions = []
for exts in SUPPORTED_FILES.values():
    all_extensions.extend(exts)

file_input = st.sidebar.file_uploader(
    "📁 Fayl yuklang",
    type=all_extensions,
    help="CSV, Excel, JSON, TXT, PDF, DOCX, va rasm fayllarini qo'llab-quvvatlaydi"
)

if file_input is not None:
    current_file_id = f"{file_input.name}_{file_input.size}"

    if st.session_state.uploaded_file_id != current_file_id:
        category, file_ext = get_file_category(file_input.name)

        if category:
            os.makedirs("Files", exist_ok=True)
            os.makedirs("Output", exist_ok=True)

            timestamp = int(time.time())
            safe_filename = f"{timestamp}_{file_input.name}"
            save_path = os.path.join("Files", f"saved_{safe_filename}")

            try:
                result = safe_file_write(save_path, file_input.getbuffer())

                if isinstance(result, str):
                    save_path = result

                st.session_state.file_path = save_path
                st.session_state.file_type = category
                st.session_state.file_name = safe_filename
                st.session_state.uploaded_file_id = current_file_id

                time.sleep(0.2)
                content = read_file_content(save_path, file_ext)
                st.session_state.file_content = content

                if category == 'data' and isinstance(content, pd.DataFrame):
                    st.session_state.df = content

                st.sidebar.success(f"✅ {file_input.name} yuklandi!")
                st.sidebar.info(f"📋 Turi: {category.upper()} ({file_ext})")

            except Exception as e:
                st.sidebar.error(f"❌ Fayl yuklashda xatolik: {str(e)}")
                st.session_state.file_path = None
                st.session_state.file_type = None
                st.session_state.file_name = None
        else:
            st.sidebar.error("❌ Qo'llab-quvvatlanmaydigan fayl turi!")

if st.session_state.file_path and st.session_state.file_type and st.session_state.file_name:
    st.sidebar.header("🔄 Format O'zgartirish")

    available_formats = CONVERT_OPTIONS.get(st.session_state.file_type, CONVERT_OPTIONS['all'])

    select_format = st.sidebar.selectbox(
        "Format tanlang",
        available_formats,
        help="Faylingiz turidan kelib chiqqan holda mavjud formatlar"
    )

    if st.sidebar.button("🔄 O'zgartirish", use_container_width=True):
        with st.sidebar:
            with st.spinner(f"{select_format} ga aylantirilmoqda..."):
                try:
                    output_path = convert_file(st.session_state.file_name, select_format)

                    mime_types = {
                        'PDF': 'application/pdf',
                        'CSV': 'text/csv',
                        'Excel': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        'JSON': 'application/json',
                        'TXT': 'text/plain'
                    }

                    time.sleep(0.3)
                    if os.path.exists(output_path):
                        with open(output_path, "rb") as f:
                            file_data = f.read()

                        st.download_button(
                            label=f"⬇️ {select_format} yuklab olish",
                            data=file_data,
                            file_name=os.path.basename(output_path),
                            mime=mime_types.get(select_format, 'application/octet-stream'),
                            use_container_width=True
                        )
                        st.success(f"✅ {select_format} tayyor!")
                    else:
                        st.error("❌ Fayl yaratilmadi!")

                except Exception as e:
                    st.error(f"❌ Xatolik: {str(e)}")

with st.sidebar.expander("ℹ️ Konvertatsiya imkoniyatlari"):
    st.markdown("""
    **Ma'lumotlar:**
    - Excel ↔ CSV
    - JSON ↔ CSV ↔ Excel
    - Barchasi → PDF

    **Hujjatlar:**
    - PDF → TXT
    - DOCX → TXT, PDF
    - TXT → PDF

    **Rasmlar:**
    - Barcha formatlar → PDF
    """)

if st.session_state.file_content is not None:
    file_type = st.session_state.file_type

    if file_type == 'data':
        st.subheader("📊 Ma'lumotlar Ko'rinishi")

        tabs = st.tabs(["📋 Jadval", "📈 Vizualizatsiya", "ℹ️ Ma'lumot"])

        with tabs[0]:
            st.dataframe(st.session_state.df, use_container_width=True)

        with tabs[1]:
            display_data_visualizations(st.session_state.df)

        with tabs[2]:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Qatorlar soni", len(st.session_state.df))
            with col2:
                st.metric("Ustunlar soni", len(st.session_state.df.columns))
            with col3:
                st.metric("Xotira", f"{st.session_state.df.memory_usage(deep=True).sum() / 1024:.2f} KB")

            st.markdown("#### Ustunlar Ma'lumoti")
            st.dataframe(st.session_state.df.dtypes.to_frame('Turi'), use_container_width=True)

    elif file_type == 'document':
        st.subheader("📄 Hujjat Matni")
        st.text_area("Matn", st.session_state.file_content, height=400)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.info(f"📝 Belgilar: {len(st.session_state.file_content)}")
        with col2:
            st.info(f"📄 So'zlar: {len(st.session_state.file_content.split())}")
        with col3:
            st.info(f"📑 Qatorlar: {len(st.session_state.file_content.splitlines())}")

    elif file_type == 'image':
        st.subheader("🖼️ Rasm Ko'rinishi")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.image(st.session_state.file_content, use_container_width=True)

        img = st.session_state.file_content
        col1, col2, col3 = st.columns(3)
        with col1:
            st.info(f"📐 Kenglik: {img.size[0]} px")
        with col2:
            st.info(f"📐 Balandlik: {img.size[1]} px")
        with col3:
            st.info(f"🎨 Format: {img.format}")

st.divider()
st.subheader("💬 AI Yordamchisi")

if client is None:
    st.warning("⚠️ GROQ API key topilmadi. .env faylida GROQ_API_KEY ni sozlang.")

user_input = st.chat_input("Savolingizni yozing...")

if user_input and user_input != st.session_state.last_user_input:
    context = ""

    if st.session_state.file_type == 'data' and st.session_state.df is not None:
        context = f"Ma'lumotlar haqida:\n{st.session_state.df.head(10).to_string()}"
    elif st.session_state.file_type == 'document':
        context = f"Hujjat matni:\n{st.session_state.file_content[:1000]}"

    st.session_state.messages.append({"role": "user", "content": user_input})
    response = send_message(user_input, context)
    st.session_state.messages.append({"role": "assistant", "content": response})
    st.session_state.last_user_input = user_input

for msg in st.session_state.messages:
    with st.chat_message(msg['role']):
        st.write(msg['content'])

# Footer
st.sidebar.divider()
st.sidebar.markdown("### 📌 Qo'llab-quvvatlanadigan formatlar")
st.sidebar.markdown("""
- **Ma'lumotlar**: CSV, Excel, JSON
- **Hujjatlar**: TXT, PDF, DOCX
- **Rasmlar**: JPG, PNG, GIF, BMP
""")