import streamlit as st
from docx import Document
from docx.shared import Inches, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
import io

def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    """Mengatur margin dalam cell tabel menjadi 0 agar stiker bisa mepet rapi"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m_name, m_val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m_name}')
        node.set(qn('w:w'), str(m_val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def remove_table_borders(table):
    """Menghilangkan garis tepi tabel agar tidak ikut tercetak"""
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

# --- Konfigurasi Halaman Streamlit ---
st.set_page_config(page_title="Sticker to Word Layout-er", page_icon="🖨️")
st.title("🖨️ Generator Layout Stiker A4")
st.write("Masukkan gambar dan ukuran stiker untuk memenuhi kertas A4 secara otomatis dalam format Word.")

# --- Input User ---
uploaded_file = st.file_uploader("Upload Gambar Stiker (PNG)", type=["png"])
sticker_size_cm = st.number_input("Ukuran Sisi Stiker Persegi (cm)", min_value=1.0, max_value=20.0, value=5.0, step=0.1)

if uploaded_file is not None:
    # Tampilkan preview gambar
    image = Image.open(uploaded_file)
    st.image(image, caption="Preview Gambar", use_container_width=True)
    
    if st.button("Generate File Word"):
        with st.spinner("Sedang memproses layout..."):
            # 1. Inisialisasi Dokumen Word
            doc = Document()
            
            # 2. Atur Ukuran Kertas ke A4 (21 x 29.7 cm) dan Margin 1 cm
            sections = doc.sections
            for section in sections:
                section.page_width = Cm(21.0)
                section.page_height = Cm(29.7)
                section.top_margin = Cm(1.0)
                section.bottom_margin = Cm(1.0)
                section.left_margin = Cm(1.0)
                section.right_margin = Cm(1.0)

            # 3. Hitung Area Cetak Bersih
            printable_width = 21.0 - 2.0  # 19.0 cm
            printable_height = 29.7 - 2.0 # 27.7 cm
            
            # 4. Hitung Jumlah Baris dan Kolom yang Muat
            cols_count = int(printable_width // sticker_size_cm)
            rows_count = int(printable_height // sticker_size_cm)
            
            if cols_count == 0 or rows_count == 0:
                st.error("Ukuran stiker terlalu besar untuk ukuran kertas A4!")
            else:
                st.info(f"Layout Terdeteksi: {rows_count} Baris x {cols_count} Kolom (Total {rows_count * cols_count} Stiker)")
                
                # 5. Buat Tabel Sesuai Hitungan
                table = doc.add_table(rows=rows_count, cols=cols_count)
                remove_table_borders(table)
                
                # Konversi gambar upload-an ke BytesIO agar bisa dibaca python-docx tanpa save local
                img_byte_arr = io.BytesIO()
                image.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)
                
                # 6. Isi Setiap Cell dengan Gambar
                for r in range(rows_count):
                    # Atur tinggi baris secara spesifik
                    table.rows[r].height = Cm(sticker_size_cm)
                    for c in range(cols_count):
                        cell = table.cell(r, c)
                        cell.width = Cm(sticker_size_cm)
                        set_cell_margins(cell) # Nol-kan margin cell
                        
                        # Masukkan gambar ke dalam cell
                        paragraph = cell.paragraphs[0]
                        run = paragraph.add_run()
                        run.add_picture(img_byte_arr, width=Cm(sticker_size_cm), height=Cm(sticker_size_cm))
                
                # 7. Simpan Dokumen ke Memory Buffer untuk Download
                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)
                
                # 8. Tombol Download
                st.success("File Word berhasil dibuat!")
                st.download_button(
                    label="📥 Download File Word (.docx)",
                    data=doc_buffer,
                    file_name="layout_stiker_ready.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  )
