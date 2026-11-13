# Central de Recuperação — Módulo Assessorias

Aplicação estática em HTML, CSS e JavaScript para estudar metas e políticas de comissão de assessorias de cobrança. Preparada para hospedagem no GitHub Pages. O módulo B2B aparece apenas como espaço futuro; os cálculos implementados estão no módulo **Assessorias**.

## 1. Executar localmente

1. Extraia a pasta do projeto.
2. Crie uma pasta `dados` na raiz e, opcionalmente, coloque nela o Excel confidencial como `historico.xlsx`.
3. Abra um terminal na raiz do projeto e execute:

   ```bash
   python -m http.server 8000
   ```

4. Acesse `http://localhost:8000`.

O sistema tenta carregar `dados/historico.xlsx` automaticamente. Também é possível usar **Importar histórico** e selecionar o arquivo manualmente. A leitura e as simulações ocorrem no navegador.

## 2. Publicar no GitHub Pages e proteger o histórico

1. Envie ao repositório os arquivos HTML, CSS, JavaScript, README, `.gitignore` e `.nojekyll`.
2. Não envie o Excel original, exportações ou qualquer arquivo confidencial. O `.gitignore` já exclui `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.csv` e `.pdf`, além das pastas locais de dados e exportações.
3. No GitHub, abra `Settings → Pages`, escolha `Deploy from a branch`, a branch principal e a pasta `/ (root)`.
4. No site publicado, selecione o Excel pelo navegador a cada sessão.

O GitHub Pages não consegue acessar pastas privadas do computador. A planilha selecionada não é enviada ao site ou a uma API. A opção de carregamento automático pela pasta `dados/` serve apenas para execução local; não publique essa pasta.

## 3. Estrutura do módulo Assessorias

- **Visão geral:** indicadores consolidados, ranking e desempenho por assessoria.
- **Metas e evolução:** metas baseadas no histórico, metas manuais e evolução composta mês a mês.
- **Comissões:** política vigente por aging, estudos condicionais por assessoria, taxas novas editáveis, estimativa de custo e gráfico comparativo.
- **Histórico detalhado:** linhas mensais usadas nas análises.
- **Exportar estudo:** único ponto de exportação para Excel ou PDF, com os dados das áreas do módulo.

## 4. Indicadores de recuperação

O simulador trabalha com três indicadores distintos:

- **Recuperado bruto:** coluna `RECUPERADO` do histórico.
- **Recuperado ativo:** coluna `RECUPERADO ATIVO`.
- **Recuperado cancelado:** coluna `RECUPERADO CANCELADO`.

Os três campos são apresentados como indicadores independentes. Valores ausentes de ativo ou cancelado permanecem indisponíveis e não são convertidos em zero.

## 5. Metas e evolução

O indicador pode ser bruto, ativo ou cancelado. A referência histórica pode ser a média mensal do período ou o último mês com dado. Há três métodos:

1. **Crescimento sobre o histórico:** adiciona um percentual livre à referência e mantém a meta mensal constante.
2. **Valor manual por assessoria:** permite informar metas sem depender do histórico; cada assessoria pode receber valor próprio.
3. **Aumento percentual a cada mês:** define uma meta inicial e crescimento composto sobre o mês anterior, com horizonte de 1 a 36 meses.

É possível criar um cenário sem histórico, cadastrar assessorias manualmente e preencher metas livres.

## 6. Laboratório de comissões por aging

A aba Comissões começa exibindo a política atual por faixa de atraso, com percentuais B2C e B2B. Cada estudo independente permite escolher assessoria e definir a condição para liberar uma nova tabela.

**Condição de desbloqueio:**

- percentual da carteira cancelada (exige informar manualmente o saldo atualizado da carteira daquela assessoria); ou
- valor fixo de recuperação (R$).

O indicador que deve atingir o gatilho pode ser recuperado bruto, ativo ou cancelado. Assim, por exemplo, é possível configurar que uma assessoria precisa recuperar 10% do saldo cancelado usando o indicador de recuperado cancelado.

**Nova tabela:** as taxas novas B2C e B2B são editadas por faixa de aging. O simulador mostra o mix histórico de cada faixa, a comissão estimada com a política atual e com a nova, a diferença por faixa, o custo incremental total e o saldo incremental após o custo adicional de comissão.

**Interpretação do gatilho:** a nova tabela é simulada sobre todo o valor recuperado no período em que o gatilho é atingido. Caso a média histórica já esteja igual ou acima da meta, a recuperação adicional necessária é zero e o custo da nova tabela é avaliado sobre o patamar médio atual.

**Premissas relevantes:**

- O histórico mensal detalhado informa recuperação por aging e a composição geral CPF/CNPJ, mas não cruza aging × CPF/CNPJ × situação ativa/cancelada por título.
- Para ponderar B2C/B2B, o simulador associa CPF a B2C e CNPJ a B2B. Essa correspondência é uma hipótese operacional que deve ser validada com a política interna.
- A distribuição de aging e a composição CPF/CNPJ são estimadas pelo conjunto de meses incluídos nos filtros. O mix é ponderado pelo volume histórico, não é uma média simples de percentuais mensais.
- O saldo de carteira cancelada por assessoria não existe no histórico utilizado; precisa ser atualizado manualmente para gatilhos percentuais.
- Por essas limitações, os custos são estimativas gerenciais e não substituem o recálculo de comissão título a título.

## 7. Exportação unificada

Há um único botão **Exportar estudo**, centralizado na barra do estudo. A partir dele, escolha Excel ou PDF.

O arquivo Excel (`.xlsx`) contém:

- `Resumo`: origem, filtros, premissas e indicadores gerais.
- `Desempenho por assessoria`: cobrança, recuperações bruto/ativo/cancelado, comissões e custo efetivo.
- `Metas`: referências e metas por assessoria.
- `Evolução mensal`: projeção de cada mês.
- `Cenários de comissão`: gatilhos, metas, custo atual/novo, patamar avaliado, recuperação adicional e saldo incremental após o custo adicional de comissão.
- `Tabelas simuladas`: todas as taxas novas por faixa e estudo, mesmo quando a memória de custo não puder ser calculada.
- `Custo por aging`: memória de cálculo detalhada de todos os estudos salvos.
- `Histórico filtrado`: registros mensais utilizados no recorte atual.
- `Política atual`: taxas B2C/B2B por aging encontradas no Excel.

O PDF é preparado como relatório de impressão e inclui a visão geral, metas, evolução, comparação e memória por aging de todos os cenários, política atual e histórico filtrado. Na janela de impressão, selecione **Salvar como PDF**. O PDF pode ocupar várias páginas dependendo do volume de histórico e estudos.

As exportações são geradas localmente. Trate os arquivos exportados como confidenciais.

## 8. Privacidade e dependências

- O Excel é processado em memória no navegador.
- O projeto não possui API, banco de dados, login, analytics ou sincronização.
- Parâmetros e dados importados não são gravados em `localStorage` ou `IndexedDB`; ao fechar ou atualizar a página será necessário importar e configurar novamente.
- SheetJS Community Edition 0.20.3 é carregada pelo CDN oficial para ler e gerar arquivos Excel. O conteúdo da planilha não é enviado ao CDN.
- É necessária conexão com a internet para carregar a biblioteca SheetJS.

## 9. Estrutura de arquivos

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

## 10. Validação antes de adotar uma política

Use o simulador para explorar hipóteses e comparar incentivos. Antes de converter uma simulação em regra oficial, valide os dados e premissas com as áreas responsáveis e teste os resultados com meses completos. Documente critérios de elegibilidade, estornos, cancelamentos, títulos renegociados, impostos e regras contratuais. Os relatórios são estudos de apoio, não autorização de pagamento.
