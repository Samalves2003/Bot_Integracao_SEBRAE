# Bot de Automação: Integração de Usuários SAS

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white )
![Tkinter](https://img.shields.io/badge/Tkinter-2C599D?style=for-the-badge&logo=python&logoColor=white )
![Status](https://img.shields.io/badge/Status-Em%20Desenvolvimento-brightgreen?style=for-the-badge )

## 📌 Descrição do Projeto

Este é um bot de automação desenvolvido em Python com o objetivo de otimizar e agilizar o processo de integração de usuários na plataforma **FOCO - SEBRAE**. O grande diferencial deste bot está na sua capacidade de **autocorreção**, que identifica e resolve automaticamente os erros de integração mais comuns, minimizando a necessidade de intervenção manual e garantindo um fluxo de trabalho mais eficiente.

o projeto em breve contara com uma aplicação de interface gráfica simples, desenvolvida com **Tkinter**, que guia o usuário durante a execução. Para garantir a segurança e o controle do processo, a automação só é iniciada após o login manual do usuário na plataforma.

## ✨ Funcionalidades Principais

✅ **Automação Completa do Fluxo**
- Executa todo o processo de integração de usuários, desde a busca inicial pelo CPF até o clique final no botão de integração na plataforma.

🔁 **Sistema Inteligente de Autocorreção**
- Detecta os erros mais frequentes que ocorrem durante a integração.
- Executa rotinas de correção automática personalizadas para cada tipo de problema, aumentando a taxa de sucesso sem intervenção humana.

📝 **Relatórios Detalhados (Logs)**
- Ao final do processo, gera uma planilha Excel com o status detalhado de cada usuário processado, classificado como:
  - `Integrado com Sucesso`
  - `Falha na Correção`
  - `Erro Geral`

🧪 **Modo de Depuração e Conexão Segura**
- Inicia o navegador Google Chrome em um modo que permite ao bot se conectar a uma sessão já autenticada.
- Garante que a automação ocorra apenas após o login manual do usuário, aproveitando a sessão ativa para executar as tarefas.

## 🛠️ Tecnologias Utilizadas

- **Python:** Linguagem principal do projeto.
- **Selenium:** Para a automação da interação com o navegador web.

