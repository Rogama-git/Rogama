"""
Serviço de Geração de Orçamentos - Rogama SL
Recebe JSON com orçamento → preenche Excel → converte PDF → retorna base64
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
    'ROGAMA':   '/app/templates/Ppto_Rogama_-_2022.xlsx',
    'MULTIMAP': '/app/templates/Presupuesto_Multimap_Logo_Nuevo.xlsm'
}

def preencher_rogama(orcamento: dict, dest: str):
    wb = load_workbook(TEMPLATES['ROGAMA'])
    ws = wb['PRESUPUESTO']

    # Cabeçalho
    ws['F42'] = orcamento.get('expediente', '')
    ws['F43'] = orcamento.get('cliente', '')
    ws['F44'] = orcamento.get('direccion', '')
    ws['F45'] = orcamento.get('localidad', '')
    ws['F46'] = orcamento.get('cp', '')
    ws['F47'] = orcamento.get('telefono', '')
    ws['F48'] = orcamento.get('fecha', datetime.now().strftime('%d/%m/%Y'))

    # Nº expediente e direção
    ws['I64'] = orcamento.get('expediente', '')
    ws['I65'] = f"{orcamento.get('direccion', '')} - {orcamento.get('localidad', '')}"

    # Items — linhas 71-81
    items = orcamento.get('items', [])
    for i, item in enumerate(items[:11]):
        row = 71 + i
        ws[f'A{row}'] = item.get('codigo', '')
        ws[f'E{row}'] = item.get('concepto', '')
        ws[f'J{row}'] = item.get('cantidad', 1)
        ws[f'K{row}'] = item.get('precio_unitario', 0)
        ws[f'L{row}'] = round((item.get('cantidad', 1)) * (item.get('precio_unitario', 0)), 2)

    # Totais
    ws['L82'] = round(orcamento.get('total_sem_iva', 0), 2)
    ws['K83'] = 0.21
    ws['L83'] = round(orcamento.get('iva', 0), 2)
    ws['L84'] = round(orcamento.get('total_com_iva', 0), 2)
    ws['J86'] = datetime.now().strftime('%d/%m/%Y')

    wb.save(dest)


def preencher_multimap(orcamento: dict, dest: str):
    wb = load_workbook(TEMPLATES['MULTIMAP'])
    ws = wb['PRESUPUESTO']

    # Cabeçalho
    ws['E5']  = orcamento.get('expediente', '')
    ws['E6']  = orcamento.get('cliente', '')
    ws['E7']  = orcamento.get('direccion', '')
    ws['E8']  = orcamento.get('localidad', '')
    ws['E9']  = orcamento.get('cp', '')
    ws['E10'] = orcamento.get('fecha', datetime.now().strftime('%d/%m/%Y'))
    ws['E11'] = orcamento.get('telefono', '')

    # Items — linhas 136+
    items = orcamento.get('items', [])
    for i, item in enumerate(items[:50]):
        row = 136 + i
        ws[f'B{row}'] = item.get('unidad', 'ud')
        ws[f'C{row}'] = item.get('concepto', '')
        qty   = item.get('cantidad', 1)
        preco = item.get('precio_unitario', 0)
        ws[f'D{row}'] = qty
        ws[f'E{row}'] = preco
        ws[f'F{row}'] = round(qty * preco, 2)

    wb.save(dest)


def excel_para_pdf(excel_path: str) -> str:
    """Converte Excel para PDF usando LibreOffice headless"""
    output_dir = os.path.dirname(excel_path)
    result = subprocess.run([
        'libreoffice', '--headless', '--convert-to', 'pdf',
        '--outdir', output_dir, excel_path
    ], capture_output=True, text=True, timeout=60)

    pdf_path = excel_path.rsplit('.', 1)[0] + '.pdf'
    if not os.path.exists(pdf_path):
        # Tenta com extensão xlsm
        pdf_path = excel_path.replace('.xlsm', '.pdf').replace('.xlsx', '.pdf')

    return pdf_path if os.path.exists(pdf_path) else None


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

        # Cria ficheiro temporário
        with tempfile.TemporaryDirectory() as tmpdir:
            if template == 'MULTIMAP':
                excel_dest = os.path.join(tmpdir, f"{nome_base}.xlsx")
                shutil.copy(TEMPLATES['MULTIMAP'], excel_dest)
                preencher_multimap(orcamento, excel_dest)
            else:
                excel_dest = os.path.join(tmpdir, f"{nome_base}.xlsx")
                shutil.copy(TEMPLATES['ROGAMA'], excel_dest)
                preencher_rogama(orcamento, excel_dest)

            # Converte para PDF
            pdf_path = excel_para_pdf(excel_dest)

            if pdf_path and os.path.exists(pdf_path):
                with open(pdf_path, 'rb') as f:
                    pdf_b64 = base64.b64encode(f.read()).decode('utf-8')

                with open(excel_dest, 'rb') as f:
                    excel_b64 = base64.b64encode(f.read()).decode('utf-8')

                return jsonify({
                    'success': True,
                    'expediente': exp,
                    'template': template,
                    'nome_ficheiro': f"{nome_base}.pdf",
                    'pdf_base64': pdf_b64,
                    'excel_base64': excel_b64
                })
            else:
                # Retorna só o Excel se PDF falhar
                with open(excel_dest, 'rb') as f:
                    excel_b64 = base64.b64encode(f.read()).decode('utf-8')

                return jsonify({
                    'success': True,
                    'expediente': exp,
                    'template': template,
                    'nome_ficheiro': f"{nome_base}.xlsx",
                    'pdf_base64': None,
                    'excel_base64': excel_b64,
                    'aviso': 'PDF não gerado, só Excel disponível'
                })

    except Exception as e:
        return jsonify({'success': False, 'erro': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 80))
    app.run(host='0.0.0.0', port=port, debug=False)
