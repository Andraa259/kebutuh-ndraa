import streamlit as st
from docx import Document
from docx.shared import Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
import io

# Library untuk PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas

def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    """Mengatur margin internal cell tabel menjadi 0 agar ukuran gambar presisi"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m_name, m_val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m_name}')
        node.set(qn('w:w'), str(m_val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def remove_table_borders(table):
    """Menghilangkan garis tepi tabel di Word"""
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

# --- Konfigurasi Halaman Streamlit ---
st.set_page_config(page_title="Sticker Layout Generator", page_icon="🖨️")
st.title("🖨️ Generator Layout Stiker A4 (Word & PDF)")
st.write("Masukkan gambar stiker dan ukurannya. Sistem akan menyusunnya otomatis di kertas A4 dengan jeda 1 cm.")

# --- Input User ---
uploaded_file = st.file_uploader("Upload Gambar Stiker (PNG / JPG)", type=["png", "jpg", "jpeg"])
sticker_size_cm = st.number_input("Ukuran Sisi Stiker Persegi (cm)", min_value=1.0, max_value=15.0, value=5.0, step=0.1)

if uploaded_file is not None:
    # Membuka Gambar (Mendukung PNG & JPG)
    image = Image.open(uploaded_file)
    st.image(image, caption="Preview Gambar Stiker", use_container_width=True)
    
    # 1. Parameter Dasar Lembar Kerja (A4 dengan Margin 1cm sekeliling)
    margin_cm = 1.0
    gap_cm = 1.0
    printable_width = 21.0 - (2 * margin_cm)   # 19.0 cm
    printable_height = 29.7 - (2 * margin_cm)  # 27.7 cm
    
    # 2. Hitung jumlah kolom & baris maksimal menggunakan rumus deret spasial:
    # (count * size) + ((count - 1) * gap) <= printable_area
    cols_count = int((printable_width + gap_cm) // (sticker_size_cm + gap_cm))
    rows_count = int((printable_height + gap_cm) // (sticker_size_cm + gap_cm))
    
    # Proteksi jika ukuran terlalu besar
    if cols_count <= 0 or rows_count <= 0:
        st.error("Ukuran stiker terlalu besar untuk diletakkan di kertas A4 dengan jarak antar gambar 1 cm!")
    else:
        st.info(f"✨ Layout Optimal: **{rows_count} Baris** × **{cols_count} Kolom** (Total: {rows_count * cols_count} Stiker per halaman)")
        
        # Sediakan dua kolom tombol pilihan format output
        col1, col2 = st.columns(2)
        
        # --- PROSES GENERATE WORD ---
        with col1:
            if st.button("🚀 Buat File Word (.docx)", use_container_width=True):
                with st.spinner("Menyusun layout Word..."):
                    doc = Document()
                    
                    # Atur ukuran halaman A4 & Margin halaman
                    for section in doc.sections:
                        section.page_width = Cm(21.0)
                        section.page_height = Cm(29.7)
                        section.top_margin = Cm(margin_cm)
                        section.bottom_margin = Cm(margin_cm)
                        section.left_margin = Cm(margin_cm)
                        section.right_margin = Cm(margin_cm)
                    
                    # Total kolom/baris tabel Word termasuk baris/kolom pembatas (spacer)
                    word_cols = (cols_count * 2) - 1
                    word_rows = (rows_count * 2) - 1
                    
                    table = doc.add_table(rows=word_rows, cols=word_cols)
                    remove_table_borders(table)
                    
                    # Konversi gambar ke format byte buffer agar python-docx bisa membaca
                    img_byte_arr = io.BytesIO()
                    image.save(img_byte_arr, format='PNG')
                    img_byte_arr.seek(0)
                    
                    # Atur dimensi grid tabel dan masukkan gambar
                    for r in range(word_rows):
                        is_row_gap = (r % 2 != 0)
                        table.rows[r].height = Cm(gap_cm) if is_row_gap else Cm(sticker_size_cm)
                        
                        for c in range(word_cols):
                            is_col_gap = (c % 2 != 0)
                            cell = table.cell(r, c)
                            set_cell_margins(cell)
                            
                            if is_col_gap:
                                cell.width = Cm(gap_cm)
                            else:
                                cell.width = Cm(sticker_size_cm)
                                # Jika ini bukan baris kosong maupun kolom kosong, tempatkan gambar
                                if not is_row_gap:
                                    paragraph = cell.paragraphs[0]
                                    run = paragraph.add_run()
                                    run.add_picture(img_byte_arr, width=Cm(sticker_size_cm), height=Cm(sticker_size_cm))
                    
                    # Simpan file ke memory buffer
                    word_buffer = io.BytesIO()
                    doc.save(word_buffer)
                    word_buffer.seek(0)
                    
                    st.success("File Word siap diunduh!")
                    st.download_button(
                        label="📥 Download File Word",
                        data=word_buffer,
                        file_name="layout_stiker.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

        # --- PROSES GENERATE PDF ---
        with col2:
            if st.button("🚀 Buat File PDF (.pdf)", use_container_width=True):
                with st.spinner("Menyusun layout PDF..."):
                    # Konversi gambar ke mode RGB jika gambar bertipe PNG transparan supaya ReportLab tidak error
                    pdf_img = image.convert("RGB") if image.mode != "RGB" else image
                    img_byte_arr = io.BytesIO()
                    pdf_img.save(img_byte_arr, format='JPEG')
                    img_byte_arr.seek(0)
                    
                    # Buat objek canvas reportlab di memory buffer
                    pdf_buffer = io.BytesIO()
                    c = canvas.Canvas(pdf_buffer, pagesize=A4)
                    
                    # Baca ukuran halaman A4 sesungguhnya dalam satuan poin internal ReportLab
                    width_a4, height_a4 = A4 
                    
                    # Buat objek gambar ReportLab dari buffer objek Pillow Image
                    from reportlab.lib.utils import ImageReader
                    reader = ImageReader(img_byte_arr)
                    
                    # Gambar stiker dengan koordinat cartesius (titik 0,0 dimulai dari Kiri Bawah halaman)
                    for r in range(rows_count):
                        for col in range(cols_count):
                            # Hitung posisi X dan Y dari sisi kiri dan atas kertas dalam cm
                            pos_x_cm = margin_cm + (col * (sticker_size_cm + gap_cm))
                            pos_y_cm = margin_cm + (r * (sticker_size_cm + gap_cm))
                            
                            # Konversikan centimeter ke unit ReportLab (points)
                            # Karena sumbu Y ReportLab bergerak dari bawah ke atas, kita balik titik mulanya dari atas halaman
                            x = pos_x_cm * cm
                            y = height_a4 - (pos_y_cm * cm) - (sticker_size_cm * cm)
                            
                            c.drawImage(reader, x, y, width=sticker_size_cm*cm, height=sticker_size_cm*cm)
                    
                    c.showPage()
                    c.save()
                    pdf_buffer.seek(0)
                    
                    st.success("File PDF siap diunduh!")
                    st.download_button(
                        label="📥 Download File PDF",
                        data=pdf_buffer,
                        file_name="layout_stiker.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
