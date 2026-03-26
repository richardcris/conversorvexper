# Conversor de Banco para Excel

Aplicacao desktop em Python para Windows com interface moderna, abertura animada e exportacao organizada para Excel.

## O que o sistema faz

- abre um arquivo de banco SQLite ou Firebird
- identifica todas as tabelas disponiveis
- apos selecionar o arquivo, inicia automaticamente a leitura completa
- converte automaticamente todas as tabelas para Excel
- mostra pre-visualizacao dos dados
- le cada tabela linha por linha e coluna por coluna
- gera um arquivo Excel com uma aba por tabela
- cria uma aba Resumo com totais e campos encontrados
- aplica cabecalho formatado, filtro e congelamento da primeira linha

## Formatos aceitos

- .db
- .sqlite
- .sqlite3
- .fdb

## Requisitos

- Python 3.10 ou superior
- Windows com Tkinter habilitado

## Instalar dependencias

```powershell
c:/python314/python.exe -m pip install -r requirements.txt
```

## Executar o sistema

```powershell
c:/python314/python.exe app.py
```

## Gerar executavel .exe

```powershell
build_exe.bat
```

O executavel sera criado em:

```text
dist\CONVERSOR - VEXPER atualizado.exe
```

## Gerar instalador

```powershell
build_installer.bat
```

Saida:

```text
installer\Instalador CONVERSOR - VEXPER.exe
```

## Atualizacao automatica via GitHub

O sistema agora pode buscar novas versoes pelo GitHub Releases.

Feed recomendado para os clientes instalados:

```text
https://github.com/USUARIO/REPOSITORIO/releases/latest/download/
```

Defina estas variaveis antes de rodar o build do instalador para publicar automaticamente no GitHub:

```powershell
$env:VEXPER_GITHUB_REPO="USUARIO/REPOSITORIO"
$env:VEXPER_GITHUB_TOKEN="SEU_TOKEN_GITHUB"
```

Depois rode:

```powershell
build_installer.bat
```

Ou publique direto pelo GitHub Actions:

```text
.github/workflows/release.yml
```

Fluxo do Actions:

- dispara em tags `v*`
- compila o `.exe`
- gera o instalador
- publica `latest.json` e o instalador na release do GitHub
- anexa os arquivos tambem como artifacts do workflow

O processo vai:

- compilar o instalador
- gerar o arquivo `latest.json`
- publicar o instalador e o manifesto na release `vVERSAO`

Para os clientes ja instalados:

- abra `Configuracao`
- ative a atualizacao automatica
- informe o feed do GitHub Releases
- na proxima abertura o sistema baixa a versao nova e instala por cima da antiga

## Fluxo de uso

1. Clique em Selecionar banco de dados.
2. Se for Firebird, confirme host, porta, usuario e senha.
3. Escolha o arquivo do banco.
4. O sistema inicia automaticamente a leitura completa.
5. O Excel e salvo automaticamente na mesma pasta do banco.
6. O nome padrao e `NOME_DO_BANCO_convertido.xlsx`.

## Fluxo manual opcional

- O botao Ler estrutura do banco continua disponivel se voce quiser apenas inspecionar as tabelas.
- O botao Exportar para Excel continua disponivel se voce quiser escolher manualmente o nome e o local do arquivo.

## Observacao tecnica

- SQLite e aberto diretamente pelo arquivo.
- Firebird .fdb usa conexao com o servidor Firebird local, por padrao em 127.0.0.1:3050.
- Os campos padrao da interface para Firebird sao: usuario `SYSDBA`, senha `masterkey` e charset `WIN1252`.
- O auto-update suporta pasta local compartilhada ou GitHub Releases.