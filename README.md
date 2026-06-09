# Gerador Inteligente de Relatórios Excel 📊🚀

Este é um aplicativo web local desenvolvido em Python (Flask) e HTML/CSS/JS moderno que utiliza a inteligência artificial **Google Gemini** para criar e formatar planilhas Excel automaticamente a partir de arquivos de dados enviados (Excel, CSV, TXT) e instruções em linguagem natural.

O aplicativo gera dinamicamente código Python customizado utilizando as bibliotecas `pandas` e `openpyxl` para estruturar e estilizar planilhas de forma profissional.

## 🛠️ Tecnologias
- **Backend:** Python 3, Flask
- **Processamento de Dados & Excel:** Pandas, OpenPyXL
- **IA:** Google Generative AI (Gemini 2.5 Flash)
- **Frontend:** HTML5, CSS3 (Design Moderno com Tema Escuro), JavaScript Vanilla
- **Repositório:** [GitHub - excel-generator](https://github.com/eullon1234-creator/excel-generator)

## 🚀 Como Executar o Projeto Localmente

### Pré-requisitos
Certifique-se de ter o Python instalado em seu computador.

### Passo 1: Clonar o repositório
```bash
git clone https://github.com/eullon1234-creator/excel-generator.git
cd excel-generator
```

### Passo 2: Instalar as dependências
```bash
pip install -r requirements.txt
```

### Passo 3: Configurar a Chave de API do Gemini
Crie um arquivo `.env` na raiz do projeto contendo sua chave do Google Gemini:
```env
GEMINI_API_KEY=sua_chave_de_api_aqui
```
*(Você também pode colar a chave diretamente na tela do aplicativo ao executá-lo)*.

### Passo 4: Iniciar o Servidor
```bash
python app.py
```
O aplicativo estará disponível no seu navegador em: **`http://127.0.0.1:5000`**

## 💡 Como Usar
1. Faça o upload das suas planilhas de origem (ou arquivos de texto/CSV).
2. Na caixa de texto, escreva suas instruções (ex: *"Mescle as planilhas de vendas e produtos pela coluna ID_Produto. Ordene por Faturamento de forma decrescente, pinte o cabeçalho de azul marinho com texto branco e adicione uma coluna de média calculada"*).
3. Clique em **Gerar Relatório Excel**.
4. Acompanhe as etapas de geração e, ao terminar, clique em **Baixar Planilha Excel Gerada**.
