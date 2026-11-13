# Central de Recuperação — Simulador de Assessorias

Aplicação estática em HTML, CSS e JavaScript para analisar o histórico de recuperação de assessorias e criar cenários independentes de metas e comissões. Preparada para publicação no GitHub Pages.

## 1. Executar localmente

1. Extraia a pasta do projeto.
2. Crie uma pasta `dados` na raiz do projeto.
3. Coloque o Excel confidencial em `dados/historico.xlsx` (ou mantenha o nome original `analise_assessorias_14-10-26(1).xlsx`).
4. Abra um terminal na raiz do projeto e execute:

   ```bash
   python -m http.server 8000
   ```

5. Acesse `http://localhost:8000`.

O sistema tenta carregar automaticamente o arquivo que estiver em `dados/`. Também é possível selecionar o Excel pelo botão de importação. Não abra `index.html` diretamente por `file://` para usar o carregamento automático.

## 2. Publicar no GitHub Pages

1. Crie um repositório e envie os arquivos do projeto.
2. Mantenha o Excel na pasta local `dados/`, protegida pelo `.gitignore`.
3. **Não comite o Excel nem arquivos derivados com dados confidenciais.**
4. No GitHub, abra `Settings → Pages`.
5. Em `Build and deployment`, selecione `Deploy from a branch`, escolha a branch principal e a pasta `/ (root)`.
6. No site publicado, selecione o arquivo Excel pelo navegador a cada sessão.

O GitHub Pages hospeda arquivos estáticos e não tem permissão para ler pastas privadas do computador do usuário. A leitura e os cálculos são feitos no dispositivo; a planilha não é enviada ao GitHub Pages ou a uma API.

## 3. Estrutura da aplicação

A navegação superior separa módulos. **Assessorias** contém tudo que está implementado atualmente. **B2B** aparece como módulo futuro e não compartilha os cálculos de Assessorias. As seções de Assessorias são:

- **Visão geral:** ranking, indicadores consolidados e resumo por assessoria.
- **Metas e evolução:** três modos independentes de definição de meta.
- **Comissões:** tabelas progressivas, taxa da faixa alcançada e gatilho de bônus.
- **Política atual:** leitura das faixas B2C/B2B por atraso identificadas no Excel, apenas como referência.
- **Histórico detalhado:** consulta dos registros de origem.

## 4. Construtor de metas

O indicador da meta pode ser:

- `Recuperado bruto`: usa o valor bruto recuperado informado no histórico.
- `Somente recuperado ativo`: considera somente o componente ativo.
- `Somente recuperado cancelado`: considera somente o componente cancelado.
- `Ativo + cancelado`: soma os dois componentes apenas quando ambos estiverem preenchidos no mesmo registro.

O método de referência pode usar a média mensal do período ou o último mês com dado disponível. Meses sem valor no indicador escolhido não são convertidos em zero. Se o histórico não contiver determinado componente, a interface informa a ausência e deixa o resultado como não disponível.

Há três métodos para a meta:

1. **Crescimento sobre o histórico:** aplica um percentual livre à referência e mantém a meta mensal constante durante a projeção.
2. **Valor manual por assessoria:** permite preencher um valor comum para todas as assessorias do recorte e editar cada linha individualmente. O valor não fica limitado à média histórica.
3. **Aumento percentual a cada mês:** define um aumento inicial para o primeiro mês e uma taxa composta que incide sobre o mês anterior. O horizonte pode variar de 1 a 36 meses.

A tabela individual apresenta referência, meta inicial, meta final, total projetado e aumento inicial. Uma segunda tabela apresenta o consolidado de cada mês projetado.

## 5. Construtor de comissões

As regras podem ser gerais ou específicas para uma assessoria. Uma regra individual começa como cópia do modelo geral e passa a ser independente quando editada.

- **Percentual progressivo por faixa:** cada taxa incide apenas sobre a parcela do valor dentro daquela faixa.
- **Percentual da faixa atingida:** a taxa da maior faixa alcançada incide sobre toda a base de comissão.
- **Gatilho com bônus:** define uma comissão-base, um indicador para avaliar o gatilho, o valor mínimo e um acréscimo em pontos percentuais (p.p.). Por exemplo, 8% de comissão-base + 1,5 p.p. resulta em 9,5%. Ao atingir o valor mínimo, o acréscimo é aplicado sobre toda a base de comissão do período, não apenas sobre o excedente.

A base da comissão pode acompanhar o indicador da meta ou usar recuperado bruto, ativo, cancelado ou ativo + cancelado. O indicador que libera o bônus também pode ser escolhido separadamente.

**Limite dos dados:** o histórico disponível é agregado por assessoria e mês e não possui a relação individual de títulos. Portanto, o gatilho estima o custo sobre o total do indicador escolhido. Não recalcula nem audita comissão título a título. Quando uma base de comissão for diferente do indicador da meta, a projeção dessa outra base usa a referência histórica; no modo de meta manual, o valor digitado é aplicado somente ao indicador da meta.

A comissão histórica é exibida como comparação. A taxa efetiva histórica é uma referência agregada e não reproduz necessariamente a política original por atraso e tipo de documento.

## 6. Exportações

O botão **Exportar estudo Excel** gera um arquivo `.xlsx` com várias abas:

- `Resumo`: origem, filtros, premissas e totais.
- `Metas`: referência e metas por assessoria.
- `Evolução mensal`: projeção mês a mês.
- `Comissões`: base de cálculo, gatilho, taxa, custo e líquido.
- `Regras de comissão`: configuração das tabelas e gatilhos.
- `Histórico filtrado`: registros incluídos no recorte atual.
- `Política atual`: faixas B2C/B2B, quando encontradas na planilha.

O botão **Gerar relatório PDF** abre a janela de impressão do navegador com um relatório resumido do estudo. Escolha **Salvar como PDF** no destino de impressão. A orientação do relatório é paisagem para acomodar as tabelas.

As exportações são geradas localmente pelo navegador. O usuário deve tratar os arquivos exportados como confidenciais, pois eles podem conter dados históricos da empresa.

## 7. Privacidade e dependências

- Os dados do Excel são processados em memória no navegador.
- O projeto não possui API, banco de dados, login, analytics ou sincronização de dados.
- Os dados importados e as regras editadas não são persistidos em `localStorage` ou `IndexedDB`; ao fechar/atualizar a página, será necessário importar o histórico novamente e refazer os parâmetros.
- A biblioteca SheetJS Community Edition 0.20.3 é carregada pelo CDN oficial (`https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js`) para ler e gerar `.xlsx`. O conteúdo do arquivo não é enviado ao CDN.
- É necessária conexão com a internet para carregar a biblioteca SheetJS.

## 8. Estrutura de arquivos

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

## 9. Validação antes de adotar uma política

Use a aplicação para explorar hipóteses e comparar cenários. Antes de converter uma simulação em regra oficial de pagamento, valide as premissas com as áreas responsáveis, teste os resultados com meses históricos completos e documente critérios de elegibilidade, estornos, cancelamentos, títulos renegociados, impostos e regras contratuais. O arquivo Excel exportado e o relatório PDF são estudos de apoio, não autorização de pagamento.
