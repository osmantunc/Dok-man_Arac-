import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import requests

def tr_upper(text): return text.replace("i", "İ").replace("ı", "I").upper()

def tr_title(text):
    words = text.split()
    res = []
    for w in words:
        if not w: continue
        first = w[0].replace("i", "İ").replace("ı", "I").upper()
        rest = w[1:].replace("İ", "i").replace("I", "ı").lower()
        res.append(first + rest)
    return " ".join(res)

def detect_heading_level(text, style_name):
    if style_name in ['Heading 1', 'Başlık 1']: return 1
    if style_name in ['Heading 2', 'Başlık 2']: return 2
    if style_name in ['Heading 3', 'Başlık 3']: return 3
    if style_name in ['Heading 4', 'Başlık 4']: return 4
    
    text = text.strip()
    if re.match(r'^\d+\.\d+\.\d+\.\d+\.?\s+', text): return 4
    if re.match(r'^\d+\.\d+\.\d+\.?\s+', text): return 3
    if re.match(r'^\d+\.\d+\.?\s+', text): return 2
    if re.match(r'^(\d+|[A-Z])\.\s+', text): return 1
    
    if 0 < len(text.split()) < 8 and text.isupper() and not text.endswith('.'): return 1
    return 0

# Sorunsuz Çalışan Doğrudan API Bağlantısı
def check_grammar_api(text):
    try:
        response = requests.post(
            'https://api.languagetool.org/v2/check',
            data={'text': text, 'language': 'tr'},
            timeout=5
        )
        if response.status_code == 200:
            matches = response.json().get('matches', [])
            # Düzeltmeleri sondan başa doğru uygula ki metin kayması yaşanmasın
            matches.sort(key=lambda x: x['offset'], reverse=True)
            for match in matches:
                if match['replacements']:
                    rep = match['replacements'][0]['value']
                    off = match['offset']
                    length = match['length']
                    text = text[:off] + rep + text[off+length:]
    except:
        pass # İnternet kesilirse formatlamaya devam et
    return text

def process_file(input_file):
    output_file = input_file.replace(".docx", "_Duzenlenmis.docx")
    doc = docx.Document(input_file)
    
    for para in doc.paragraphs:
        if not para.text.strip(): continue

        # Yazım denetimi API'ye gönderiliyor
        para.text = check_grammar_api(para.text)

        level = detect_heading_level(para.text, para.style.name)
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.line_spacing = 1.0

        if level == 1:
            para.text = tr_upper(para.text)
            para.paragraph_format.space_before = Pt(12)
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.left_indent = Cm(1.0)
            para.paragraph_format.first_line_indent = Cm(-1.0)
        elif level in [2, 3, 4]:
            para.text = tr_title(para.text)
            para.paragraph_format.space_before = Pt(12)
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.left_indent = Cm(1.0)
            para.paragraph_format.first_line_indent = Cm(-1.0)
        else:
            para.paragraph_format.space_before = Pt(6)
            para.paragraph_format.space_after = Pt(6)
            para.paragraph_format.left_indent = Pt(0)
            para.paragraph_format.first_line_indent = Pt(0)

        for run in para.runs:
            run.font.name = 'Arial'
            run.font.size = Pt(11)

    doc.save(output_file)
    return output_file

def run_app():
    root = tk.Tk()
    root.withdraw()
    input_file = filedialog.askopenfilename(title="Düzenlenecek Word Dosyasını Seçin", filetypes=[("Word Dosyaları", "*.docx")])
    if input_file:
        try:
            messagebox.showinfo("Bilgi", "İşlem başlıyor. Lütfen bekleyin...")
            out_path = process_file(input_file)
            messagebox.showinfo("Başarılı", f"İşlem tamamlandı!\nKaydedilen dosya:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Hata", f"Bir hata oluştu:\n{e}")

if __name__ == "__main__":
    run_app()
