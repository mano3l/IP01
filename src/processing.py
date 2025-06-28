import os
import re
import subprocess
import platform
import shutil
import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Alignment
from config import HOUR_TO_COLUMN_MAP

def extract_data_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        text = "".join(page.get_text("text") for page in doc)
        doc.close()
        data_start_index = text.index("ACUMU") + len("ACUMU")
        data_end_index = text.index("Total")
        table_text = text[data_start_index:data_end_index]
        tokens = table_text.split()
        sales_data = []
        i = 0
        while i < len(tokens):
            token = tokens[i]
            if re.match(r'^\d{2}:\d{2}$', token) and i + 2 < len(tokens):
                try:
                    sales_data.append({'hora': token, 'tcs': int(tokens[i + 1]), 'vendas': float(tokens[i + 2])})
                    i += 3
                    continue
                except ValueError:
                    pass
            i += 1
        return sales_data
    except Exception as e:
        raise ValueError(f"Erro ao ler o PDF: {e}")

def create_workbook_data(extracted_data):
    data_map = {entry['hora']: entry for entry in extracted_data}
    final_data = []
    for hora_map, _ in sorted(HOUR_TO_COLUMN_MAP.items(), key=lambda item: int(item[0].split(':')[0]) % 24):
        entry = data_map.get(hora_map, {'tcs': 0, 'vendas': 0.0})
        final_data.append({'hora': hora_map, 'tcs': entry['tcs'], 'vendas': entry['vendas']})
    return final_data

def save_xlsx_file(final_data, template_path, output_path):
    try:
        data_map = {entry['hora']: entry for entry in final_data}
        workbook = openpyxl.load_workbook(template_path)
        sheet = workbook.active
        shrink_alignment = Alignment(shrink_to_fit=True, vertical='center', horizontal='center')
        for hora_map, column_letter in HOUR_TO_COLUMN_MAP.items():
            entry = data_map.get(hora_map, {'tcs': 0, 'vendas': 0.0})
            vendas_cell, tcs_cell = sheet[f'{column_letter}10'], sheet[f'{column_letter}11']
            vendas_cell.alignment = tcs_cell.alignment = shrink_alignment
            vendas_cell.value, tcs_cell.value = entry['vendas'], entry['tcs']
        workbook.save(output_path)
    except Exception as e:
        raise IOError(f"Erro ao salvar o arquivo XLSX: {e}")

def find_libreoffice_path():
    if platform.system() == "Windows":
        path = shutil.which("soffice.exe")
        if path: return path
        common_paths = [r"C:\Program Files\LibreOffice\program\soffice.exe",
                        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]
        for path in common_paths:
            if os.path.exists(path): return path
    else:  # Linux
        return shutil.which("soffice")

def convert_to_pdf_with_libreoffice(command_path, xlsx_path):
    try:
        result = subprocess.run(
            [command_path, "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(xlsx_path), xlsx_path],
            capture_output=True, text=True, check=True, timeout=60
        )
        return os.path.splitext(xlsx_path)[0] + ".pdf"
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"O LibreOffice retornou um erro:\n{e.stderr}")
    except Exception as e:
        raise RuntimeError(f"Ocorreu um erro inesperado na conversão para PDF:\n{e}")