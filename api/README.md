# API Vercel

`supabase-snapshot.js` conecta o site ao Supabase sem expor a chave secreta no navegador.

Configure estas variaveis no projeto Vercel:

```text
SUPABASE_URL=https://elmvhncukxgrdqmgnltj.supabase.co
SUPABASE_SERVICE_ROLE_KEY=...
```

Use a secret key/service role somente em variavel de ambiente da Vercel. Nunca coloque essa chave em `index.html`, arquivos `.js` publicos ou no GitHub.

Depois de salvar as variaveis, faca novo deploy.
