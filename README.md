# Sistema de Envio de Cobranças via WhatsApp

## Descrição  
Este projeto é uma aplicação desktop desenvolvida com **Python**, utilizando as bibliotecas **Tkinter** para interface gráfica, **Selenium** para automação de navegador e **Pandas** para manipulação de dados. A aplicação lê uma planilha Excel (`Cobrancas.xlsx`), identifica devedores com base em critérios fornecidos pelo usuário e envia mensagens automáticas de cobrança via WhatsApp Web. Após o envio, atualiza a planilha e salva os resultados em um novo arquivo (`Cobrancas2.xlsx`).

## Funcionalidades Principais  
- **Interface Gráfica (`Tkinter`)**:  
  - Permite que o usuário insira o nome do funcionário e o critério de situação de pagamento (ex.: "Pendente").  
  - Possui um botão para iniciar o processo de envio de mensagens.  
- **Leitura de Planilha (`Pandas`)**:  
  - Lê a planilha `Cobrancas.xlsx` com informações como nome, telefone, valor devido e situação de pagamento.  
- **Envio de Mensagens Automáticas (`Selenium`)**:  
  - Acessa o WhatsApp Web e envia mensagens personalizadas para cada devedor com valores em aberto.  
  - Cada mensagem contém o nome do devedor, o valor devido e uma solicitação para regularização.  
- **Atualização da Planilha**:  
  - Marca a situação de pagamento como "Mensagem Enviada" para os contatos que receberam a mensagem.  
  - Salva as alterações em uma nova planilha (`Cobrancas2.xlsx`).  

## Como Contribuir  
1. Faça um fork do repositório.  
2. Crie um branch para suas alterações.  
3. Commit suas mudanças.  
4. Faça push para o branch.  
5. Abra um Pull Request.  

Para mais detalhes sobre como contribuir, veja a [documentação oficial do GitHub sobre Pull Requests](https://docs.github.com/pt/pull-requests/collaborating-with-pull-requests).  