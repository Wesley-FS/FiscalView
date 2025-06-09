# Leitor de XML - NFe 🔍📄

## Descrição
Este projeto é um leitor de XML para **notas fiscais eletrônicas (NFe)** desenvolvido por **Wesley Ferreira, Analista de Sistemas da LC1 Contadores**. A aplicação permite a leitura de arquivos XML e apresenta as informações mais relevantes para que o usuário possa realizar lançamentos sem precisar abrir o PDF da nota fiscal.

## Funcionalidades
- Selecionar uma **pasta de XMLs** com arquivos de notas fiscais.
- Exibir informações detalhadas das notas fiscais em uma **tabela organizada**.
- **Copiar informações** para a área de transferência (Ctrl + C) com um duplo clique.
- Exportação dos dados para **CSV**, facilitando integração com outras ferramentas.

## Como Usar
1. **Selecionar XMLs**: Clique no botão **"Selecionar Pasta de XMLs"** na parte superior da aplicação.
2. **Escolher Pasta**: Selecione a pasta onde os arquivos XML estão armazenados.
3. **Visualizar os Arquivos**: A tela principal exibirá todos os arquivos encontrados, com suas informações essenciais.
4. **Copiar Informações**: Clique duas vezes sobre um campo desejado para copiá-lo automaticamente para a área de transferência do Windows (**Ctrl + C**).
5. **Exportar para CSV**: Caso deseje salvar os dados da tabela, clique no botão **"Exportar CSV"** na parte inferior da interface.

## Informações Exibidas
A tabela do programa apresenta os seguintes dados das notas fiscais:
- Número da Nota
- CNPJ do Cliente
- Nome do Cliente
- Código Protheus
- Quantidade
- Valor Unitário
- Valor Total
- TES
- CFOP
- IPI
- ICMS
- Natureza da Operação (NatOp)

## Requisitos do Sistema
- Windows 10 ou superior
- .NET Framework instalado
- Permissão de acesso a arquivos XML na máquina

