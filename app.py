"""
Servico de Geracao de Orcamentos - Rogama SL
Recebe JSON com orcamento -> preenche Excel -> converte PDF -> retorna base64
"""
from flask import Flask, request, jsonify
import json
import os
import shutil
import base64
import subprocess
import tempfile
from datetime import datetime
from openpyxl import load_workbook

app = Flask(__name__)

TEMPLATES = {
    'ROGAMA':   '/app/templates/Ppto Rogama - 2022.xlsx',
    'MULTIMAP': '/app/templates/Presupuesto Multimap Logo Nuevo.xlsm'
}

def preencher_rogama(orcamento: dict, dest: str):
    wb = load_workbook(dest)
    ws = wb['PRESUPUESTO']

    ws['F42'] = orcamento.get('expediente', '')
    ws['F43'] = orcamento.get('cliente', '')
    ws['F44'] = orcamento.get('direccion', '')
    ws['F45'] = orcamento.get('localidad', '')
    ws['F46'] = orcamento.get('cp', '')
    ws['F47'] = orcamento.get('telefono', '')
    ws['F48'] = orcamento.get('fecha', datetime.now().strftime('%d/%m/%Y'))
    ws['I64'] = orcamento.get('expediente', '')
    ws['I65'] = f"{orcamento.get('direccion', '')} - {orcamento.get('localidad', '')}"

    items = orcamento.get('items', [])
    for i, item in enumerate(items[:11]):
        row = 71 + i
        ws[f'A{row}'] = item.get('codigo', '')
        ws[f'E{row}'] = item.get('concepto_corto') or item.get('concepto', '')
        ws[f'J{row}'] = item.get('cantidad', 1)
        ws[f'K{row}'] = item.get('precio_unitario', 0)
        ws[f'L{row}'] = round((item.get('cantidad', 1)) * (item.get('precio_unitario', 0)), 2)

    ws['L82'] = round(orcamento.get('total_sem_iva', 0), 2)
    ws['K83'] = 0.21
    ws['L83'] = round(orcamento.get('iva', 0), 2)
    ws['L84'] = round(orcamento.get('total_com_iva', 0), 2)
    ws['J86'] = datetime.now().strftime('%d/%m/%Y')

    wb.save(dest)


def preencher_multimap(orcamento: dict, dest: str):
    wb = load_workbook(dest)
    ws = wb['PRESUPUESTO']

    ws['E5']  = orcamento.get('expediente', '')
    ws['E6']  = orcamento.get('cliente', '')
    ws['E7']  = orcamento.get('direccion', '')
    ws['E8']  = orcamento.get('localidad', '')
    ws['E9']  = orcamento.get('cp', '')
    ws['E10'] = orcamento.get('fecha', datetime.now().strftime('%d/%m/%Y'))
    ws['E11'] = orcamento.get('telefono', '')

    items = orcamento.get('items', [])
    for i, item in enumerate(items[:50]):
        row = 136 + i
        ws[f'B{row}'] = item.get('unidad', 'ud')
        ws[f'C{row}'] = item.get('concepto_corto') or item.get('concepto', '')
        qty   = item.get('cantidad', 1)
        preco = item.get('precio_unitario', 0)
        ws[f'D{row}'] = qty
        ws[f'E{row}'] = preco
        ws[f'F{row}'] = round(qty * preco, 2)

    wb.save(dest)


def excel_para_pdf(excel_path: str) -> str:
    """Converte Excel para PDF usando LibreOffice headless"""
    try:
        output_dir = os.path.dirname(excel_path)
        subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'pdf',
            '--outdir', output_dir, excel_path
        ], capture_output=True, text=True, timeout=60)
        pdf_path = excel_path.rsplit('.', 1)[0] + '.pdf'
        return pdf_path if os.path.exists(pdf_path) else None
    except Exception:
        return None  # LibreOffice nao disponivel


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'rogama-orcamentos'})


@app.route('/gerar-orcamento', methods=['POST'])
def gerar_orcamento():
    try:
        data = request.get_json()
        orcamento = data.get('orcamento', data)

        exp      = orcamento.get('expediente', 'SEM_EXP')
        template = orcamento.get('template', 'ROGAMA').upper()
        if not template or template not in TEMPLATES:
            template = 'MULTIMAP' if exp.startswith('A') else 'ROGAMA'

        localidad = orcamento.get('localidad', '').replace('/', '-')[:20]
        nome_base = f"{exp} - {localidad}"

        with tempfile.TemporaryDirectory() as tmpdir:
            ext = '.xlsm' if template == 'MULTIMAP' else '.xlsx'
            excel_dest = os.path.join(tmpdir, f"{nome_base}{ext}")
            shutil.copy(TEMPLATES[template], excel_dest)

            if template == 'MULTIMAP':
                preencher_multimap(orcamento, excel_dest)
            else:
                preencher_rogama(orcamento, excel_dest)

            # Tenta converter para PDF
            pdf_path = excel_para_pdf(excel_dest)

            with open(excel_dest, 'rb') as f:
                excel_b64 = base64.b64encode(f.read()).decode('utf-8')

            if pdf_path and os.path.exists(pdf_path):
                with open(pdf_path, 'rb') as f:
                    pdf_b64 = base64.b64encode(f.read()).decode('utf-8')
                nome_ficheiro = f"{nome_base}.pdf"
            else:
                pdf_b64 = None
                nome_ficheiro = f"{nome_base}{ext}"

            return jsonify({
                'success': True,
                'expediente': exp,
                'template': template,
                'nome_ficheiro': nome_ficheiro,
                'pdf_base64': pdf_b64,
                'excel_base64': excel_b64
            })

    except Exception as e:
        return jsonify({'success': False, 'erro': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 80))
    app.run(host='0.0.0.0', port=port, debug=False)
