# Declaração de Quitação de Débitos

## Descrição
Este projeto em Python automatiza a geração de declarações de quitação de débitos escolares em formato Word. A aplicação conecta-se a uma base de dados SQL Server, executa consultas personalizadas para obter informações financeiras de alunos, formata os dados em um documento profissional e inclui um logo personalizado.

---

## Funcionalidades
- Conexão com banco de dados SQL Server (ambientes HML ou PRD).
- Consulta SQL customizada para obter dados financeiros.
- Geração de documentos Word com formatação específica, incluindo:
  - Logo no cabeçalho.
  - Ajuste automático de margens.
  - Tabelas com colunas ajustadas uniformemente.
  - Formatação de datas e valores conforme padrão brasileiro.
- Possibilidade de customizar o caminho de salvamento do relatório.

---

## Pré-requisitos
Certifique-se de que as seguintes bibliotecas estão instaladas:
- `python-docx`: Para manipulação de documentos Word.
- `pyodbc`: Para conexão com o banco de dados SQL Server.
- `requests`: Para download do logo a partir de um link.

Para instalar as dependências, execute:
```bash
pip install python-docx pyodbc requests
