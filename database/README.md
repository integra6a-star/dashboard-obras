# Banco Supabase

Esta pasta prepara a migracao do dashboard estatico para Supabase/Postgres.

## Cuidado com dados existentes

O arquivo `schema.sql` foi escrito para ser conservador:

- usa `create table if not exists`;
- usa `create index if not exists`;
- nao usa `drop table`;
- nao usa `truncate`;
- nao apaga dados existentes.

O importador `import_to_supabase.py` tambem nao apaga tabelas. Ele so grava quando voce passa `--write`.

## 1. Criar tabelas

No Supabase:

1. Abra o projeto.
2. Va em **SQL Editor**.
3. Cole o conteudo de `database/schema.sql`.
4. Clique em **Run**.

## 2. Importar os dados atuais

No terminal, dentro da pasta do projeto:

```powershell
$env:SUPABASE_URL="https://SEU-PROJETO.supabase.co"
$env:SUPABASE_SERVICE_ROLE_KEY="SUA_SERVICE_ROLE_KEY"
& "C:\Users\micro\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" database\import_to_supabase.py --write
```

Para importar apenas uma area:

```powershell
& "C:\Users\micro\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" database\import_to_supabase.py --write --only obras
```

Opcoes aceitas em `--only`:

- `obras`
- `pds`
- `funcionarios`
- `medicao`
- `almoxarifado`
- `reclamacoes`

## 3. Seguranca

As tabelas ficam com RLS ativo. O SQL libera leitura apenas para usuarios autenticados. A chave `service_role` deve ficar somente no servidor ou no computador de administracao, nunca dentro do site publico.
