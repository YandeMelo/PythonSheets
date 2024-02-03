# Automação de Avaliação no Google Sheets
Solução do Desafio Tunts.Rocks 2024, integrando a planilha do google sheets.

Este script em Python automatiza o processo de avaliação para uma planilha no Google Sheets que contém informações de notas e presenças dos alunos. Ele calcula as médias das notas, avalia as situações dos alunos com base em critérios predefinidos e atualiza a planilha com os resultados.

## Pré-requisitos

Antes de executar o script, certifique-se de ter o seguinte:

1. **Python Instalado:** Garanta que você tenha o Python instalado em sua máquina. Você pode baixá-lo [aqui](https://www.python.org/downloads/).

2. **Pacotes Python Necessários:** Instale os pacotes Python necessários usando o seguinte comando:

    ```bash
    pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client
    ```

3. **Credenciais da API do Google Sheets:**

   - Habilite a API do Google Sheets para o seu projeto no Google Cloud.
   - Crie credenciais (Service Account Key) com acesso à API do Google Sheets.
   - Faça o download do arquivo de credenciais em JSON e salve-o como `client_secret.json` no diretório do projeto.

4. **Planilha no Google Sheets:**

   - Crie uma planilha no Google Sheets com as informações dos alunos. A planilha deve ter colunas para Matrícula, Aluno, Faltas, P1, P2 e P3.

## Uso

1. Clone este repositório:

    ```bash
    git clone https://github.com/seu_usuario/nome_do_repositorio.git
    ```

2. Navegue até o diretório do projeto
   
3. Execute o script:

    ```bash
    python PythonSheets.py
    ```

## Observações

- O script assume que os dados dos alunos estão armazenados na planilha chamada "engenharia_de_software".

- Certifique-se de que o arquivo `token.json` esteja presente no diretório do projeto. Caso contrário, o script irá guiá-lo pelo processo de autorização do OAuth 2.0.

- O script imprimirá informações sobre o processo de avaliação, incluindo médias e situações para cada aluno.
  
