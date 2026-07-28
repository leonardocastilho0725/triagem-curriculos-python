# 📊 Analisador Inteligente de Currículos (Eixo Alto Tietê)

Este projeto é uma automação inteligente desenvolvida em Python para triagem automatizada de currículos em formato PDF. Ele foi estruturado para otimizar processos de recrutamento e seleção corporativos, aplicando filtros precisos de perfil de vaga, escolaridade e geolocalização.

## 🚀 Funcionalidades e Melhorias Implementadas

- **Filtro Geográfico Restrito e Ampliado:** Triagem automática com foco nas cidades de fácil acesso da região do Alto Tietê (Itaquaquecetuba, Poá, Mogi das Cruzes, Ferraz de Vasconcelos e Suzano).
- **Extração Robusta de Contatos:** 
  - Captura e-mails de forma automatizada.
  - Usa `finditer` com expressões regulares de múltiplos dígitos para mapear telefones/WhatsApp válidos, priorizando o formato numérico BR mais plausível e descartando matches errados (como CPFs, CEPs ou números isolados).
- **Validação Inteligente e Hierárquica de Idade:** 
  - Determina a idade seguindo uma ordem estrita de confiabilidade: Data de nascimento completa > Termo "Idade: X" > Ano isolado > Expressões genéricas (como "X anos", evitando conflito com anos de experiência).
  - Possui trava de sanidade humana para descartar automaticamente dados que resultem em idades absurdas (fora do intervalo de 14 a 90 anos).
- **Checagem Estrita de Escolaridade:** Valida o Ensino Médio de forma textual sem depender de tamanho de arquivos, evitando falsos positivos para currículos curtos. O sistema valida automaticamente o Ensino Médio caso o candidato cite estar cursando ou ter concluído Ensino Superior.
- **Relatório Corporativo Executivo:** Exportação estruturada para planilha Excel (`.xlsx`) com abas separadas de **Aprovados** e **Reprovados** (com motivos detalhados), trazendo design executivo (cabeçalho azul-marinho, efeito zebra para leitura confortável, colunas auto-ajustáveis e tags coloridas de status).

## 📁 Estrutura do Projeto

O projeto organiza-se de maneira modular:
- `analisador.py`: Código-fonte principal contendo a inteligência de processamento e os estilos visuais.
- `curriculos/`: Diretório local onde os arquivos PDF dos candidatos devem ser depositados.
- `resultado_triagem.xlsx`: Relatório em Excel gerado automaticamente após a execução da triagem.

## 🛠️ Tecnologias Utilizadas

- **Python 3** (Linguagem de programação base)
- **pdfplumber** (Leitura e extração de texto de PDFs)
- **pandas** (Tratamento, ordenação e manipulação de tabelas)
- **openpyxl** (Modelagem estética e formatação avançada do Excel)
- **regex** (Análise avançada de padrões de texto)

## 📌 Requisitos de Execução

Para rodar o código-fonte ou realizar manutenções locais, instale as dependências executando:
```bash
pip install pdfplumber regex pandas openpyxl
```

E inicie o script:
```bash
python analisador.py
```

## 🤖 Engenharia de Software Auxiliada por IA

Este projeto foi construído utilizando as melhores práticas modernas de desenvolvimento ágil assistido por Inteligência Artificial. A arquitetura lógica e o refactoring estético do sistema contaram com o suporte estratégico de:
- **Claude (Anthropic):** Responsável pelo auxílio na lógica inicial, regras de negócio e estruturação dos primeiros blocos de código.
- **Gemini (Google):** Atuação direta na resolução de conflitos de ambientes virtuais (`venv`/`uv`), compilação do executável autônomo (`.exe`), refinamento de bugs de sintaxe e estilização executiva avançada da planilha para o padrão empresarial.

A integração de prompts estruturados com a supervisão e validação humana resultou em um software rápido, seguro e pronto para produção no ambiente de RH.
