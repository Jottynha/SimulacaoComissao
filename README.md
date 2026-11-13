# Simulador de Metas e Comissões das Assessorias

Aplicação estática em HTML, CSS e JavaScript para analisar o histórico de recuperação, projetar metas e testar tabelas de comissão por valor bruto recuperado. Foi preparada para publicação no GitHub Pages.

## 1. Como executar localmente

1. Extraia a pasta do projeto.
2. Crie uma pasta chamada `dados` na raiz do projeto.
3. Copie o Excel confidencial para `dados/historico.xlsx`.
4. Inicie um servidor estático na raiz do projeto. Por exemplo, com Python:

   ```bash
   python -m http.server 8000
   ```

5. Abra `http://localhost:8000` no navegador.

O sistema tenta localizar automaticamente `./dados/historico.xlsx` e, como alternativa, o nome do arquivo de origem `./dados/analise_assessorias_14-10-26(1).xlsx`. A pasta `dados/` e arquivos Excel estão no `.gitignore` e não devem ser adicionados ao Git. A leitura automática funciona quando a aplicação é servida localmente por HTTP.

Também é possível abrir o site publicado e usar **Selecionar arquivo Excel** para importar o histórico manualmente.

> Não abra o `index.html` diretamente por `file://` para o uso local automático. Use um servidor estático local, como o comando acima ou a extensão Live Server do VS Code.

## 2. Publicação no GitHub Pages

1. Crie um repositório com os arquivos deste projeto.
2. Mantenha o Excel dentro da pasta local `dados/`, sem removê-la do `.gitignore`.
3. Envie para o GitHub somente os arquivos do projeto (`index.html`, `css/`, `js/`, `README.md`, `.gitignore` e `.nojekyll`).
4. No GitHub, abra **Settings → Pages**.
5. Em **Build and deployment**, selecione **Deploy from a branch**, escolha a branch principal e a pasta `/ (root)`.
6. Acesse o endereço informado pelo GitHub Pages.

**Importante:** GitHub Pages hospeda arquivos estáticos. Um site publicado não consegue ler uma pasta privada do computador de quem o acessa. Por isso, no endereço publicado, o Excel deverá ser selecionado pelo botão de importação a cada nova sessão. O arquivo não precisa e não deve ser comitado.

## 3. Como o simulador calcula

### Histórico e filtros

- O sistema lê a aba `Histórico Assessorias`.
- Identifica assessoria, mês, cobrança, recuperado bruto, recuperado ativo, recuperado cancelado, comissão, conversão e custo de comissão sobre recuperação.
- Linhas de totalização são descartadas para não duplicar os valores.
- Meses ausentes não são tratados como recuperação zero. A média mensal de cada assessoria considera somente os meses em que ela possui registro. Isso é relevante quando uma assessoria começa a aparecer no histórico depois das demais.
- Registros identificados como `PARCIAL` são excluídos por padrão. O filtro `Incluir mês parcial` permite incluí-los deliberadamente.

### Simulador de metas

- A referência padrão é a média mensal do período filtrado.
- Também é possível usar o último mês completo disponível de cada assessoria.
- O percentual inicial de crescimento é 10%, mas pode ser alterado pelo campo numérico ou pelo controle deslizante.
- A meta mensal é calculada sobre o indicador escolhido (recuperado bruto ou recuperado ativo).
- A projeção anual multiplica a meta mensal por 12 e não incorpora sazonalidade.

### Tabelas de comissão por volume

- A tabela geral começa com a taxa efetiva histórica ponderada (comissão registrada ÷ recuperação bruta) no recorte de meses completos.
- Novas faixas podem ser adicionadas e seus limites e percentuais alterados.
- `Progressiva por faixa`: cada percentual incide apenas sobre a parcela do valor dentro da faixa.
- `Taxa da faixa atingida`: o percentual da faixa mais alta alcançada incide sobre todo o valor recuperado.
- É possível criar uma tabela específica por assessoria. Quando uma tabela individual ainda não foi personalizada, ela herda o modelo geral.
- O comparativo mostra a comissão histórica média, a comissão simulada sobre a meta, a variação de custo e a recuperação líquida após comissão.

**Limite metodológico:** a política histórica da planilha calcula comissão por faixa de atraso e tipo de contrato (B2C/B2B). A tabela nova baseada somente no valor recuperado é uma hipótese alternativa de incentivo. A taxa efetiva histórica serve como referência inicial, mas não permite reproduzir exatamente a comissão original sem a composição dos títulos por atraso e tipo de contrato.

### Exportações

As abas de metas, comissões e histórico permitem exportar CSV, compatível com Excel. O arquivo gerado contém somente os dados do recorte e da simulação atual.

## 4. Privacidade e dependências

- O Excel é lido e tratado em memória no navegador.
- O projeto não possui API, banco de dados, login, analytics ou armazenamento remoto.
- Os dados importados não são gravados em `localStorage` ou `IndexedDB`; ao fechar ou atualizar a página, será necessário importar o histórico novamente.
- A biblioteca SheetJS Community Edition 0.20.3 é carregada do CDN oficial (`https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js`) para interpretar arquivos Excel. A biblioteca recebe apenas a requisição do próprio script; o conteúdo da planilha é passado ao JavaScript local do navegador e não é enviado ao CDN.
- Uma conexão com a internet é necessária para carregar o script SheetJS hospedado no CDN.

## 5. Estrutura

```text
simulador-comissoes-assessorias/
├── index.html
├── css/
│   └── styles.css
├── js/
│   └── app.js
├── .gitignore
├── .nojekyll
└── README.md
```

## 6. Recomendações antes de usar para aprovar pagamentos

Use o simulador para testar hipóteses e comparar cenários, não como cálculo oficial de pagamento sem validação. Antes de transformar uma regra em política, valide a fórmula com Financeiro, Cobrança e Jurídico; teste a regra com meses históricos completos; avalie o custo incremental, a recuperação líquida e possíveis efeitos de concentração em títulos de maior atraso; e documente critérios de elegibilidade, estornos, cancelamentos e pagamento de títulos renegociados.
