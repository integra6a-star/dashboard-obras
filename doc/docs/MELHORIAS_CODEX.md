# Melhorias aplicadas

Esta pasta e uma copia melhorada do dashboard original. Ela pode ser testada localmente antes de substituir a pasta `docs` publicada no GitHub Pages.

## Ajustes feitos

- Corrigido o conflito do mapa na pagina inicial: havia um script duplicado tentando carregar `obra(8).geojson`, arquivo que nao existe no pacote. A pagina agora fica apenas com o carregamento correto de `obra.geojson`.
- Corrigido o fechamento invalido do `link` do Leaflet no `index.html`.
- Reforcada a responsividade da pagina inicial para celular e tablet: menu lateral vira navegacao horizontal, filtros empilham melhor, cards e tabela ficam mais estaveis.
- Links abertos em nova aba receberam `rel="noopener"`.
- Botoes `Detalhes` que ainda nao tinham acao foram marcados como indisponiveis, evitando clique sem resposta.
- Mensagem de senha incorreta da pagina de funcionarios foi profissionalizada.
- `icon-180.png` foi reduzido de 1024x1024 para 180x180, caindo de aproximadamente 831 KB para 24 KB.

## Ponto importante sobre senha

Como o dashboard esta em GitHub Pages, ele e um site estatico. Qualquer senha escrita no HTML ou JavaScript pode ser descoberta por quem abrir o codigo da pagina. Para proteger de verdade paginas como funcionarios, medicao, reclamacoes e relatorio mensal, o ideal e uma destas opcoes:

- publicar o repositorio como privado e controlar acesso pelo GitHub;
- mover as paginas sensiveis para uma ferramenta com login real;
- criar um pequeno backend/autenticacao antes de servir esses dados.

Enquanto continuar 100% estatico, a senha funciona apenas como barreira visual simples, nao como seguranca real.

## Como publicar

Depois de testar, substitua o conteudo da pasta `docs` do repositorio pelo conteudo desta pasta `docs_improved` e envie para o GitHub Pages.
