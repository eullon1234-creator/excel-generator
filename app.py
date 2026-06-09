import os
import sys
import uuid
import traceback
import subprocess
import pandas as pd
import openpyxl
from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename
import google.generativeai as genai
from dotenv import load_dotenv

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = 'excel-report-generator-secret'
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['GENERATED_FOLDER'] = os.path.join(os.getcwd(), 'generated_reports')

# Garantir que as pastas existam
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)

def get_gemini_client(custom_api_key=None):
    """Configura e retorna o cliente Gemini usando a chave do .env ou chave customizada da UI."""
    key = custom_api_key or os.getenv("GEMINI_API_KEY")
    if not key:
        return None
    genai.configure(api_key=key)
    return genai.GenerativeModel('gemini-2.5-flash')

def extract_file_preview(filepath, filename):
    """Extrai uma prévia estruturada dos dados de entrada para servir de contexto à IA."""
    ext = os.path.splitext(filename)[1].lower()
    preview = f"--- ARQUIVO DE ENTRADA: {filename} ---\n"
    
    try:
        if ext in ['.xlsx', '.xls']:
            xl = pd.ExcelFile(filepath)
            preview += f"Tipo: Planilha Excel (.xlsx/.xls)\nAbas (Sheets): {xl.sheet_names}\n"
            for sheet in xl.sheet_names:
                df = pd.read_excel(filepath, sheet_name=sheet)
                preview += f"\n[Aba: {sheet}]\n"
                preview += f"Colunas ({len(df.columns)}): {list(df.columns)}\n"
                preview += f"Dimensões: {df.shape[0]} linhas por {df.shape[1]} colunas\n"
                preview += "Primeiras 8 linhas de exemplo:\n"
                preview += df.head(8).to_string(index=False) + "\n"
        elif ext == '.csv':
            # Detectar delimitador
            try:
                df = pd.read_csv(filepath, nrows=10)
                # Recarregar completo se for pequeno
                df = pd.read_csv(filepath)
            except Exception:
                df = pd.read_csv(filepath, sep=';')
            
            preview += f"Tipo: Arquivo CSV\n"
            preview += f"Colunas ({len(df.columns)}): {list(df.columns)}\n"
            preview += f"Dimensões: {df.shape[0]} linhas por {df.shape[1]} colunas\n"
            preview += "Primeiras 8 linhas de exemplo:\n"
            preview += df.head(8).to_string(index=False) + "\n"
        elif ext in ['.txt', '.json', '.xml']:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(5000) # Ler primeiros 5000 caracteres
                preview += f"Tipo: Arquivo de texto ({ext})\n"
                preview += f"Conteúdo (primeiros 5000 caracteres):\n{content}\n"
        else:
            preview += f"Tipo de arquivo não reconhecido para pré-visualização de dados automática.\n"
    except Exception as e:
        preview += f"Erro ao extrair prévia dos dados: {str(e)}\n"
        
    preview += "--------------------------------------\n\n"
    return preview

def run_python_code(code_str, output_path):
    """Executa o código gerado em um processo isolado, capturando logs e erros."""
    temp_script = os.path.join(app.config['UPLOAD_FOLDER'], f"script_{uuid.uuid4().hex}.py")
    with open(temp_script, 'w', encoding='utf-8') as f:
        f.write(code_str)
        
    try:
        # Executar no mesmo ambiente python rodando o Flask
        result = subprocess.run(
            [sys.executable, temp_script],
            capture_output=True,
            text=True,
            timeout=60,
            encoding='utf-8',
            errors='ignore'
        )
        
        # Deletar script temporário
        if os.path.exists(temp_script):
            os.remove(temp_script)
            
        return result.returncode == 0, result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        if os.path.exists(temp_script):
            os.remove(temp_script)
        return False, "", "Tempo limite esgotado (60 segundos) durante a geração da planilha."
    except Exception as e:
        if os.path.exists(temp_script):
            os.remove(temp_script)
        return False, "", str(e)

@app.route('/')
def index():
    # Verificar se a chave está configurada no .env
    has_key = bool(os.getenv("GEMINI_API_KEY"))
    return render_template('index.html', has_key=has_key)

@app.route('/generate', methods=['POST'])
def generate():
    user_prompt = request.form.get('prompt', '').strip()
    custom_api_key = request.form.get('apiKey', '').strip()
    
    # Obter cliente Gemini
    model = get_gemini_client(custom_api_key)
    if not model:
        return jsonify({
            'success': False, 
            'error': 'Chave de API do Gemini não configurada. Configure o arquivo .env ou informe a chave na interface.'
        }), 400
        
    if not user_prompt:
        return jsonify({'success': False, 'error': 'Por favor, informe as instruções de como quer o relatório.'}), 400

    uploaded_files = request.files.getlist('files')
    file_contexts = ""
    file_mappings = {} # Nome do arquivo original -> Caminho físico absoluto no servidor
    
    saved_files = []
    
    try:
        # Salvar arquivos enviados e extrair prévias
        for f in uploaded_files:
            if f and f.filename:
                sec_filename = secure_filename(f.filename)
                unique_name = f"{uuid.uuid4().hex}_{sec_filename}"
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_name)
                f.save(filepath)
                saved_files.append(filepath)
                
                # Mapear para passar ao script gerado pela IA
                file_mappings[f.filename] = filepath.replace('\\', '\\\\')
                # Também mapear sem acentos/caracteres especiais se necessário
                file_mappings[sec_filename] = filepath.replace('\\', '\\\\')
                
                # Extrair prévia estruturada dos dados
                file_contexts += extract_file_preview(filepath, f.filename)
                
        # Preparar nome do arquivo Excel final
        output_filename = f"relatorio_{uuid.uuid4().hex[:8]}.xlsx"
        output_filepath = os.path.join(app.config['GENERATED_FOLDER'], output_filename)
        output_filepath_escaped = output_filepath.replace('\\', '\\\\')
        
        # Montar o prompt instruindo o Gemini
        system_instruction = (
            "Você é um engenheiro de dados especialista em Python, Pandas e OpenPyXL.\n"
            "Sua tarefa é ler dados de entrada fornecidos e gerar um script Python independente "
            "que processe esses dados e gere uma planilha Excel (.xlsx) altamente profissional e formatada.\n\n"
            "Diretrizes estritas para o script Python:\n"
            "1. Use apenas as bibliotecas já instaladas: pandas, openpyxl, os, datetime, math.\n"
            "2. O script deve salvar o resultado final EXATAMENTE no caminho de destino fornecido.\n"
            "3. IMPORTANTE: Os caminhos dos arquivos de entrada no servidor estão mapeados no dicionário `INPUT_FILES`.\n"
            "   Use-os no seu código para carregar os dados! Exemplo: `pd.read_excel(INPUT_FILES['vendas.xlsx'])`.\n"
            "4. A planilha de saída deve ser profissional:\n"
            "   - Escolha uma paleta de cores moderna (por exemplo, cabeçalhos em Azul Escuro '#1F4E78' com texto branco, linhas alternadas em cinza claro).\n"
            "   - Ajuste a largura das colunas automaticamente para que nenhum texto fique cortado.\n"
            "   - Adicione grades/linhas de grade visíveis na planilha final usando openpyxl (ex: `ws.views.sheetView[0].showGridLines = True`).\n"
            "   - Formate números de forma apropriada: moedas como R$ #.##0,00, porcentagens como 0,0%, e datas.\n"
            "   - Se apropriado, crie fórmulas Excel nativas (ex: `=SUM(...)`) em vez de valores estáticos para totais.\n"
            "   - Crie abas separadas se o usuário pedir ou se fizer sentido para organizar o relatório.\n"
            "5. Retorne APENAS o código Python válido envolto em blocos de código ```python e ```. Sem explicações ou markdown adicionais."
        )
        
        prompt_content = f"""
Caminhos dos arquivos de entrada mapeados no servidor (INPUT_FILES):
{file_mappings}

Caminho de saída onde a planilha deve ser salva:
'{output_filepath_escaped}'

Prévia e dados dos arquivos enviados pelo usuário:
{file_contexts}

Instruções do usuário para o relatório final:
"{user_prompt}"

Por favor, escreva o código Python completo que realiza a transformação e formatação conforme as regras descritas.
"""

        # Chamar Gemini
        response = model.generate_content(
            contents=[system_instruction, prompt_content]
        )
        
        ai_response = response.text
        
        # Extrair código Python
        code_str = ""
        if "```python" in ai_response:
            code_str = ai_response.split("```python")[1].split("```")[0].strip()
        elif "```" in ai_response:
            code_str = ai_response.split("```")[1].split("```")[0].strip()
        else:
            code_str = ai_response.strip()
            
        # Adicionar o mapeamento de arquivos e importações essenciais no topo do código
        header_code = f"import pandas as pd\nimport openpyxl\nimport os\nfrom openpyxl.styles import Font, PatternFill, Alignment, Border, Side\nfrom openpyxl.utils import get_column_letter\n\nINPUT_FILES = {repr(file_mappings)}\n\n"
        full_code = header_code + code_str
        
        # Executar código
        success, stdout, stderr = run_python_code(full_code, output_filepath)
        
        # Auto-correção simples caso dê erro
        if not success:
            retry_prompt = f"""
O código que você gerou falhou com o seguinte erro:
--- ERRO ---
{stderr}
--- SAÍDA ---
{stdout}

Aqui está o código gerado originalmente:
```python
{code_str}
```

Por favor, corrija o código de forma que ele não cause este erro e salve corretamente o Excel em '{output_filepath_escaped}'. 
Retorne apenas o código Python completo corrigido dentro do bloco ```python.
"""
            response_retry = model.generate_content(
                contents=[system_instruction, retry_prompt]
            )
            ai_response_retry = response_retry.text
            
            if "```python" in ai_response_retry:
                code_str = ai_response_retry.split("```python")[1].split("```")[0].strip()
            elif "```" in ai_response_retry:
                code_str = ai_response_retry.split("```")[1].split("```")[0].strip()
            else:
                code_str = ai_response_retry.strip()
                
            full_code = header_code + code_str
            success, stdout, stderr = run_python_code(full_code, output_filepath)
            
        # Limpar arquivos temporários de entrada
        for filepath in saved_files:
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass
                    
        if success and os.path.exists(output_filepath):
            return jsonify({
                'success': True,
                'filename': output_filename,
                'code': code_str,
                'logs': stdout
            })
        else:
            return jsonify({
                'success': False,
                'error': 'Falha na execução do código gerado pela IA.',
                'details': stderr,
                'code': code_str
            }), 500
            
    except Exception as e:
        # Limpar arquivos temporários em caso de exceção geral
        for filepath in saved_files:
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass
        return jsonify({
            'success': False,
            'error': f'Ocorreu um erro no servidor: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@app.route('/download/<filename>')
def download(filename):
    # Evitar directory traversal
    filename = secure_filename(filename)
    return send_from_directory(app.config['GENERATED_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
