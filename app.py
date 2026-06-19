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
    """Mengatur margin internal cell tabel dalam satuan dxa (1 cm = 567 dxa)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m_name, m_val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m_name}')
        node.set(qn('w:w'), str(m_val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def remove_table_borders(table):
    """Menghilangkan garis tepi tabel agar bersih saat dicetak"""
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

# --- Konfigurasi Halaman Streamlit ---
st.set_page_config(page_title="Sticker Layout Generator", page_icon="🖨️")
st.title("🖨️ Generator Layout Stiker A4 Pro")
st.write("Layout stiker otomatis untuk memenuhi kertas A4 dengan jeda presisi dalam format Word maupun PDF.")

# --- Input User ---
uploaded_file = st.file_uploader("Upload Gambar Stiker (PNG / JPG / JPEG)", type=["png", "jpg", "jpeg"])
sticker_size_cm = st.number_input("Ukuran Sisi Stiker Persegi (cm)", min_value=1.0, max_value=20.0, value=5.0, step=0.1)

# Jeda antar gambar diatur tetap 0.6 cm sesuai evaluasi
gap_cm = 0.6

if uploaded_file is not None:
    # Tampilkan preview gambar
    image = Image.open(uploaded_file)
    st.image(image, caption="Preview Gambar", use_container_width=True)
    
    # Perhitungan Matematika Layout Dinamis
    # Lebar A4 = 21cm, Tinggi A4 = 29.7cm. Margin kertas kiri/kanan/atas/bawah = 1cm
    margin_page_cm = 1.0
    printable_width = 21.0 - (margin_page_cm * 2)   # 19.0 cm
    printable_height = 29.7 - (margin_page_cm * 2)  # 27.7 cm
    
    # Formula matematika untuk mencari jumlah stiker yang muat dengan jeda (gap)
    # Lebar_bersih >= (cols * sticker_size) + ((cols - 1) * gap)
    cols_count = int((printable_width + gap_cm) // (sticker_size_cm + gap_cm))
    rows_count = int((printable_height + gap_cm) // (sticker_size_cm + gap_cm))
    
    if cols_count == 0 or rows_count == 0:
        st.error("Ukuran stiker atau jeda terlalu besar untuk area cetak A4!")
    else:
        st.info(f"📊 **Analisis Layout:** Terdeteksi **{rows_count} Baris** x **{cols_count} Kolom** (Total: {rows_count * cols_count} Stiker per halaman A4).")
        
        # Sediakan 2 Kolom untuk Tombol Cetak Dokumen
        col1, col2 = st.columns(2)
        
        # Konversi gambar upload ke format bytes agar kompatibel dengan docx & reportlab
        img_byte_arr = io.BytesIO()
        # Jika gambar RGBA (PNG transparan), ubah ke RGB jika ingin save sebagai JPEG, 
        # namun untuk cari aman kita simpan tetap sebagai PNG di memory buffer.
        image.save(img_byte_arr, format='PNG')
        img_byte_arr.seek(0)
        
        # ----------------------------------------------------
        # PROSES GENERATE WORD (.DOCX)
        # ----------------------------------------------------
        with col1:
            if st.button("🔧 Generate File Word"):
                with st.spinner("Menyusun file Word..."):
                    doc = Document()
                    
                    # Atur spesifikasi halaman A4 & margin luar 1 cm
                    for section in doc.sections:
                        section.page_width = Cm(21.0)
                        section.page_height = Cm(29.7)
                        section.top_margin = Cm(margin_page_cm)
                        section.bottom_margin = Cm(margin_page_cm)
                        section.left_margin = Cm(margin_page_cm)
                        section.right_margin = Cm(margin_page_cm)
                    
                    # Buat tabel pembungkus stiker
                    table = doc.add_table(rows=rows_count, cols=cols_count)
                    remove_table_borders(table)
                    
                    # Konversi 0.6 cm gap ke satuan dxa untuk Word cell margin (1 cm = 567 dxa)
                    gap_in_dxa = int(gap_cm * 567)
                    
                    for r in range(rows_count):
                        # Tinggi total sel = ukuran stiker + gap bawah (kecuali baris terakhir jika ingin mepet)
                        table.rows[r].height = Cm(sticker_size_cm + gap_cm)
                        for c in range(cols_count):
                            cell = table.cell(r, c)
                            cell.width = Cm(sticker_size_cm + gap_cm)
                            
                            # Berikan jeda 0.6 cm di sisi kanan dan bawah sel menggunakan margin internal
                            set_cell_margins(cell, top=0, bottom=gap_in_dxa, left=0, right=gap_in_dxa)
                            
                            paragraph = cell.paragraphs[0]
                            run = paragraph.add_run()
                            run.add_picture(img_byte_arr, width=Cm(sticker_size_cm), height=Cm(sticker_size_cm))
                    
                    # Kirim data ke tombol unduh
                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    
                    st.success("Word siap diunduh!")
                    st.download_button(
                        label="📥 Download Word (.docx)",
                        data=doc_buffer,
                        file_name="layout_stiker.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

        # ----------------------------------------------------
        # PROSES GENERATE PDF (NATIVE REPORTLAB)
        # ----------------------------------------------------
        with col2:
            if st.button("📕 Generate File PDF"):
                with st.spinner("Mengonversi ke PDF..."):
                    pdf_buffer = io.BytesIO()
                    
                    # Buat canvas ReportLab dengan ukuran kertas A4 baku
                    c_pdf = canvas.Canvas(pdf_buffer, pagesize=A4)
                    width_a4, height_a4 = A4  # dalam satuan poin (points)
                    
                    # ReportLab membaca koordinat dari kiri bawah (0,0). Kita harus melakukan translasi koordinat ke kiri atas.
                    # Margin halaman diubah ke poin (1 cm = 1 * cm)
                    start_x = margin_page_cm * cm
                    start_y = height_a4 - (margin_page_cm * cm)
                    
                    # Simpan data gambar mentah ke objek ReportLab ImageReader
                    from reportlab.lib.utils import ImageReader
                    img_reader = ImageReader(img_byte_arr)
                    
                    for r in range(rows_count):
                        for c in range(cols_count):
                            # Hitung posisi koordinat X dan Y secara matematis
                            pos_x = start_x + c * (sticker_size_cm + gap_cm) * cm
                            # Dikurang karena koordinat Y bergerak ke bawah dari titik atas halaman
                            pos_y = start_y - (r * (sticker_size_cm + gap_cm) * cm) - (sticker_size_cm * cm)
                            
                            # Gambar stiker di koordinat tersebut
                            c_pdf.drawImage(img_reader, pos_x, pos_y, width=sticker_size_cm * cm, height=sticker_size_cm * cm)
                    
                    c_pdf.showPage()
                    c_pdf.save()
                    pdf_buffer.seek(0)
                    
                    st.success("PDF siap diunduh!")
                    st.download_button(
                        label="📥 Download PDF (.pdf)",
                        data=pdf_buffer,
                        file_name="layout_stiker.pdf",
                        mime="application/pdf"
                    )
