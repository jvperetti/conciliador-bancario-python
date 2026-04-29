# ⚡ Conciliador Bancário Automático (Mini-ERP)

## 📌 Sobre o Projeto
O **Conciliador Bancário** é uma aplicação desktop/web híbrida desenvolvida para automatizar e otimizar o fechamento financeiro de empresas. Ele cruza dados de uma planilha de Fluxo de Caixa (Master) com múltiplos extratos bancários em formato `.OFX`, identificando divergências financeiras em segundos.

O sistema transforma um trabalho manual de horas, sujeito a erros humanos e falhas de digitação, em um processo de poucos cliques com geração de relatórios automatizados e auditados.

## 🚀 Principais Funcionalidades e Problemas Resolvidos
Durante o desenvolvimento, solucionamos diversos desafios reais de regras de negócio financeiras:

* **Cruzamento Inteligente de Dados:** O algoritmo não faz apenas uma comparação "1 para 1". Se uma fatura de cartão foi lançada agrupada com encargos no fluxo, mas separada no banco, o sistema valida pelo saldo diário, evitando **falsos positivos** na auditoria.
* **Filtro de Período Dinâmico (Range Filter):** Bancos costumam "vazar" transações de meses subsequentes em arquivos OFX de períodos fechados. O sistema lê as extremidades de datas do fluxo de caixa e cria uma barreira automática, ignorando transações fora do período analisado.
* **Leitura Multi-Banco:** Suporte nativo para múltiplos arquivos OFX simultâneos (Bradesco, Banco do Brasil, Banrisul, etc.), separando as conciliações e gerando abas individuais por banco.
* **Auditoria de Divergências:** Painel limpo que exibe estritamente os dias onde o dinheiro não bateu, indicando a origem do erro ("Faltou cair na conta" ou "Não encontrado no fluxo").
* **Exportação para Excel:** Geração automática de uma planilha multi-abas detalhada utilizando `pandas` e `openpyxl`.

## 🛠️ Tecnologias Utilizadas
* **Backend:** Python 3
* **Processamento de Dados:** `pandas`, `ofxparse`, `datetime`, `re` (Expressões Regulares)
* **Manipulação de Excel:** `openpyxl`
* **Frontend / UI:** HTML5, CSS3 avançado (Animações de partículas no Canvas, Efeitos SVG), JavaScript, Bootstrap
* **Comunicação Front/Back:** `Eel` (Permite construir interfaces gráficas usando web technologies para scripts Python)
* **Build / Deploy:** `PyInstaller` (Empacotamento da aplicação em um executável `.exe` único)

## 💻 Como rodar o projeto localmente

1. Clone o repositório:
   ```bash
   git clone [https://github.com/SEU_USUARIO/conciliador-bancario.git](https://github.com/SEU_USUARIO/conciliador-bancario.git)
Instale as dependências:

Bash
pip install pandas ofxparse eel openpyxl
Execute o sistema:

Bash
python main.py
(Nota: Dados financeiros reais foram omitidos deste repositório por questões de privacidade e LGPD).