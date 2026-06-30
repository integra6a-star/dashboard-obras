# Banco dos formularios no Google Sheets

## 1. Criar a planilha

Crie uma Planilha Google chamada, por exemplo, `Banco Integra 6A`.

Copie o ID da planilha pela URL:

`https://docs.google.com/spreadsheets/d/ID_DA_PLANILHA/edit`

## 2. Criar o Apps Script

Na planilha, abra `Extensoes > Apps Script`.

Cole o conteudo do arquivo:

`google-apps-script-banco-formularios.js`

Troque:

`const SPREADSHEET_ID = "COLE_AQUI_O_ID_DA_PLANILHA";`

pelo ID real da planilha.

## 3. Publicar como Web App

No Apps Script:

1. Clique em `Implantar`.
2. Clique em `Nova implantacao`.
3. Tipo: `Aplicativo da Web`.
4. Executar como: `Eu`.
5. Quem tem acesso: `Qualquer pessoa`.
6. Clique em `Implantar`.
7. Copie a URL do Web App.

## 4. Colar a URL no site

Abra `forms-db-config.js` e cole a URL:

```js
window.INTEGRA_FORMS_DB_CONFIG = {
  endpoint: "COLE_AQUI_A_URL_DO_WEB_APP",
  token: ""
};
```

Repita a mesma alteracao em `docs/forms-db-config.js` se o link publicado usa a pasta `docs`.

## 5. Como funciona

Os formularios salvam primeiro no banco local do navegador e tentam enviar para o Google Sheets.

O painel de Meio Ambiente e o painel de Seguranca leem o Apps Script e atualizam os indicadores com os registros novos.

Se a internet cair, o registro fica guardado no navegador e o sistema tenta sincronizar quando a conexao voltar.
