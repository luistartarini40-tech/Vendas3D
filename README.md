# Vendas3D

Calculadora de precificação e painel financeiro (ROI / payback) para impressão 3D.

Aplicação em um único `index.html` (HTML + Tailwind via CDN + JavaScript puro). Os dados ficam salvos no `localStorage` do navegador.

**Acesse:** https://luistartarini40-tech.github.io/vendas3d/

## Sincronizar com Google Planilhas (opcional)

1. Crie uma planilha nova em sheets.google.com.
2. Menu **Extensões → Apps Script**, apague o conteúdo e cole o arquivo [`google-apps-script/Code.gs`](google-apps-script/Code.gs).
3. Troque `COLE_AQUI_SUA_CHAVE` por uma senha só sua e salve.
4. **Implantar → Nova implantação → App da Web**: executar como **Eu**, acesso **Qualquer pessoa**. Autorize quando pedir.
5. Copie a URL que termina em `/exec` e cole no site em **Financeiro → Nuvem**, junto com a mesma chave.

A URL e a chave ficam salvas só no navegador de cada aparelho, nunca neste repositório.
