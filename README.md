# Silver Text Filter

Visual personalizado para Power BI que pesquisa automaticamente ao **digitar ou
colar texto**, sem clicar na lupa. Baseado no [Text Filter oficial da Microsoft](https://github.com/microsoft/PowerBI-visuals-TextFilter), sob licença MIT.

## Instalar no Power BI Desktop

1. Abra seu relatório ou o `.pbix` de exemplo fornecido.
2. No painel **Visualizações**, clique em **… → Importar um visual de um arquivo**.
3. Escolha `dist/silverTextFilterD0830F5D8E244FEDA7CC35F570AA0B31.1.0.0.1.pbiviz`.
4. Adicione **Silver Text Filter** à página e arraste uma **coluna de texto** para
   **Field**. No exemplo, use `SSA_name_top100_final[Name]`.
5. Digite ou cole um termo: a busca é aplicada após **300 ms sem novas alterações**.
   Por exemplo, `will` encontra valores que contêm esse trecho.
6. Apague o texto ou clique na borracha para remover o filtro imediatamente.

Se o Power BI solicitar modo de desenvolvedor, habilite **Desenvolver um visual**
em **Arquivo → Opções e configurações → Opções → Arquivo atual → Configurações de relatório**.

O visual possui um identificador próprio e pode coexistir com o original. Para
testar no exemplo, limpe a pesquisa do Text Filter original antes de usar o novo;
dois filtros ativos no mesmo campo podem restringir os resultados conjuntamente.
As interações entre visuais e os relacionamentos do modelo determinam quais
gráficos são filtrados. Se necessário, ajuste **Formato → Editar interações**.

Esta versão foi compilada localmente; não possui a certificação do visual original
nem foi publicada no Marketplace. Ambientes que permitem apenas visuais
certificados podem bloquear sua importação ou utilização.

## Comportamento

- Mantém o operador **Contains (contém)** do original, aplicado pela API de filtros
  do Power BI ao campo selecionado. Não pesquisa várias colunas de uma vez.
- Digitação, colagem, recorte e exclusão acionam o mesmo evento `input`.
- Enter permite antecipar a busca; não é necessário para pesquisar.
- A composição de caracteres aguarda a finalização antes de pesquisar.
- Preserva o rascunho durante atualizações do host e restaura pesquisas salvas.
- Mantém ajustes de fonte, tamanho, borda e cor da borda.
- Sem campo selecionado, a entrada fica desabilitada.
- Não declara privilégios de acesso externo e não envia o texto para serviços externos.

O prazo de 300 ms é o tempo até enviar o filtro ao Power BI; a atualização dos
gráficos depende da consulta e do modelo, especialmente em DirectQuery.

## Desenvolvimento e verificação

Use Node.js 22 e npm. As ferramentas são instaladas somente no projeto:

```powershell
npm ci
npm run typecheck
npm run lint
npm run package
npm test
```

Os testes usam o JavaScript compilado do visual, eventos DOM e um host Power BI
simulado. Verificam digitação rápida, colagem, limpeza, Enter, composição de
caracteres, filtros salvos, respostas antigas do host e mudança de campo. Após
editar o código, execute `npm run package` antes de `npm test`.

A importação e o funcionamento dentro de um relatório real do Power BI ainda
precisam de validação; os testes realizados usam o host simulado.

Para depurar com o visual de desenvolvedor no Power BI Service, use `npm start` e
configure o modo de desenvolvedor conforme a documentação oficial. O servidor
local de desenvolvimento pode exigir um certificado HTTPS; ele não é necessário
para importar o `.pbiviz` compilado.

O empacotador exige um e-mail nos metadados. `local-build@example.invalid` é um
marcador sem contato real, usado nesta distribuição local. Antes de distribuir
publicamente, preencha autor, e-mail e endereço de suporte em `pbiviz.json`.

## Origem

- Código base: Microsoft PowerBI-visuals-TextFilter, versão 2.2.9.0, branch `main`.
- Referência: sample `.pbix` 2.2.9.0 enviado pelo usuário; lido sem alteração.
- Licença e atribuição Microsoft preservadas em `LICENSE` e nos fontes.
- Identidade, busca automática, tratamento dos eventos e testes adaptados neste projeto.

Documentação Microsoft: [API de filtros](https://learn.microsoft.com/en-us/power-bi/developer/visuals/filter-api),
[empacotamento](https://learn.microsoft.com/en-us/power-bi/developer/visuals/package-visual)
e [importação de visuais](https://learn.microsoft.com/en-us/power-bi/developer/visuals/import-visual).
