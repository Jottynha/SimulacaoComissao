# Simulador Cobranças Master

Aplicação web estática para estudar metas de recuperação e modelos de comissão de assessorias. Foi desenvolvida em HTML, CSS e JavaScript e pode ser publicada no GitHub Pages.

## 1. Estrutura do sistema

A navegação superior organiza o simulador por áreas:

| Página | Situação | Finalidade |
|---|---|---|
| **Assessorias** | Disponível | Análise histórica, metas, evolução e estudos de comissão por aging. |
| **B2B** | Em desenvolvimento | Área reservada para análises e regras específicas de B2B. |
| **Cobrança interna** | Em desenvolvimento | Área reservada para indicadores e simulações da operação interna. |
| **Carteira** | Em desenvolvimento | Área reservada para composição, aging e evolução da carteira. |

As páginas futuras já possuem navegação e telas próprias, mas ainda não executam cálculos.

## 2. Como iniciar

### Com histórico

1. Acesse **Assessorias**.
2. Selecione **Importar Excel**.
3. Escolha a planilha de histórico no computador.
4. O arquivo deve conter a aba `Histórico Assessorias`.
5. O simulador apresenta os indicadores e libera os filtros por assessoria e período.

A leitura é feita no navegador. O arquivo não é enviado ao GitHub Pages e não deve ser adicionado ao repositório.

### Sem histórico

1. Selecione **Criar cenário sem histórico**.
2. Informe uma ou mais assessorias, uma por linha.
3. Defina as metas na área **Metas e evolução**.

Este modo permite testar valores hipotéticos antes de ter dados históricos. As comparações com histórico e a estimativa de comissão por aging ficam limitadas quando não há uma base importada.

## 3. Indicadores de recuperação

O simulador trabalha com três indicadores independentes:

- **Recuperado bruto:** valor bruto recuperado registrado no histórico.
- **Recuperado ativo:** valor recuperado classificado como ativo.
- **Recuperado cancelado:** valor recuperado classificado como cancelado.

O bruto é um indicador próprio. Não é criado somando ativo e cancelado, pois a planilha pode possuir critérios de apuração distintos ou campos não preenchidos.

Campos sem informação são tratados como ausentes, não como zero. Isso evita reduzir médias por causa de meses ou indicadores não apurados.

## 4. Filtros e período de análise

Quando há histórico, os filtros permitem selecionar uma assessoria ou todas, o mês inicial, o mês final e a inclusão de meses parciais.

Meses parciais ficam excluídos por padrão. Alterar o recorte atualiza os indicadores e as análises que dependem do histórico. A simulação de comissão utiliza a média do período selecionado para a assessoria escolhida.

## 5. Metas e evolução

A área **Metas e evolução** permite escolher o indicador (bruto, ativo ou cancelado), a referência histórica e o método da meta.

### Referência histórica

A referência pode ser:

- **Média mensal do período:** média dos meses que possuem valor válido para o indicador.
- **Último mês completo com dado:** valor do mês mais recente disponível no recorte.

Meses sem valor válido não entram no cálculo da média.

### Métodos de definição da meta

**Crescimento sobre o histórico**

Aplica um percentual sobre a referência. Exemplo: referência mensal de R$ 100.000 e aumento de 10% resultam em uma meta mensal de R$ 110.000. A meta permanece constante ao longo do horizonte projetado.

**Valor manual por assessoria**

Permite informar a meta mensal de cada assessoria. Também é possível preencher todas com um mesmo valor e, depois, editar uma por uma. O valor não precisa acompanhar a referência histórica.

**Aumento percentual a cada mês**

Define um valor inicial e uma taxa composta mensal. Exemplo: meta inicial de R$ 100.000 e crescimento de 2% ao mês:

| Mês | Meta ilustrativa |
|---|---:|
| 1 | R$ 100.000,00 |
| 2 | R$ 102.000,00 |
| 3 | R$ 104.040,00 |
| 4 | R$ 106.120,80 |

A projeção pode ter de 1 a 36 meses. O aumento inicial sobre a referência é opcional quando existe histórico; sem histórico, o valor de partida deve ser digitado na tabela por assessoria.

### Gráfico de evolução

O gráfico apresenta a meta consolidada mês a mês. A linha azul representa a meta projetada. A linha amarela tracejada representa a referência mensal histórica; quando não há histórico, ela representa o patamar inicial do cenário. O gráfico é atualizado conforme os métodos, taxas, valores manuais e horizonte são alterados.

A tabela mensal logo abaixo apresenta os valores exatos usados no gráfico.

## 6. Comissões por aging

A área **Comissões** compara a política vigente com uma nova tabela condicionada ao alcance de uma meta.

### Política vigente

A tabela atual é lida da seção de política de comissionamento presente no Excel. As taxas B2C e B2B são exibidas por faixa de atraso (aging) e mantidas como referência. A edição da nova tabela não altera essa política.

### Criar um estudo

Cada estudo guarda, durante a sessão atual:

- nome do cenário;
- assessoria avaliada;
- indicador que libera o incentivo (bruto, ativo ou cancelado);
- condição para alcançar a meta;
- saldo da carteira cancelada, quando o gatilho for percentual;
- novas taxas B2C e B2B por faixa de aging.

É possível manter vários estudos e alternar entre eles. As alterações ficam em memória enquanto a página permanece aberta.

### Tipos de gatilho

**Percentual da carteira cancelada**

Informe o saldo da carteira cancelada da assessoria e o percentual desejado. Exemplo: saldo de R$ 2.000.000 e meta de 10% geram um gatilho de R$ 200.000. O indicador de recuperação que conta para atingir o gatilho pode ser bruto, ativo ou cancelado.

**Valor bruto de recuperação**

Informe diretamente o valor que a assessoria deve recuperar para desbloquear a nova tabela.

### Estimativa de comissão

O simulador utiliza a distribuição histórica de recuperação por aging da assessoria selecionada para estimar como o valor de recuperação se distribui entre as faixas. Em seguida:

1. calcula o peso de cada faixa no total recuperado do período;
2. aplica esse peso ao valor avaliado no gatilho (ou à média atual, caso ela já supere o gatilho);
3. calcula a comissão estimada pela política atual e pela nova tabela;
4. compara os custos e apresenta a diferença estimada.

A taxa combinada de cada faixa usa a participação histórica de CPF e CNPJ da assessoria como aproximação para B2C e B2B, respectivamente. A tabela alternativa pode ser editada em cada faixa; o resultado e o gráfico de custos são recalculados após a alteração.

Quando a média atual já alcança a meta, a recuperação adicional é apresentada como zero. Nesse caso, o custo da nova tabela é estimado no patamar médio atual, sem inventar recuperação adicional negativa.

## 7. Exportação

O botão único **Exportar estudo** fica centralizado na área de trabalho do módulo Assessorias.

### Excel (.xlsx)

O arquivo exportado reúne abas de resumo, desempenho por assessoria, metas, evolução mensal, cenários de comissão, taxas por aging, memória de cálculo e histórico filtrado. O conteúdo considera os filtros e as configurações presentes na sessão no momento da exportação.

### PDF

A opção PDF abre a janela de impressão do navegador. Selecione **Salvar como PDF**. O relatório inclui premissas, resultados, gráfico de evolução, cenários de comissão, detalhamento por aging e histórico filtrado.

## 8. Dados, privacidade e limites

- A aplicação não possui backend nem envia os dados para uma API própria.
- O Excel selecionado é interpretado na memória do navegador.
- O histórico e os cenários não são gravados automaticamente no navegador. Ao atualizar ou fechar a página, importe novamente o Excel e refaça ou reimporte os estudos necessários.
- Não inclua o Excel confidencial, arquivos CSV exportados ou relatórios PDF no repositório público.
- O arquivo `.gitignore` exclui formatos de planilha e exportações comuns, mas revise sempre o que será enviado ao GitHub.

### Limitações da estimativa por aging

A planilha mensal não cruza, por título, aging, CPF/CNPJ e situação ativa/cancelada. Assim, o custo da nova política é uma estimativa gerencial baseada na distribuição histórica por aging e na composição CPF/CNPJ. Não substitui a apuração transacional de comissão por título.

O saldo da carteira cancelada por assessoria também não é fornecido pelo histórico mensal. Por isso, no gatilho percentual, esse saldo precisa ser informado manualmente e atualizado pelo responsável pelo estudo.

Antes de transformar uma simulação em regra de pagamento, valide as premissas com a área financeira e compare o resultado com uma apuração real.

## 9. Executar localmente

É necessário servir os arquivos por HTTP para que o navegador consiga carregar a aplicação e, opcionalmente, o histórico local.

Com Python instalado, abra o terminal na pasta do projeto e execute:

```bash
python -m http.server 8000
```

Depois, acesse `http://localhost:8000`.

Se desejar carregar o Excel automaticamente durante testes locais, crie `dados/` e coloque nele `historico.xlsx`. Não use esse recurso em um repositório público e não comite a planilha.

## 10. Publicar no GitHub Pages

1. Envie para o repositório os arquivos `index.html`, `css/`, `js/`, `.gitignore`, `.nojekyll` e `README.md`.
2. Confirme que nenhum Excel, CSV ou PDF confidencial está na área de commit.
3. No GitHub, abra **Settings → Pages**.
4. Selecione a branch e a pasta que contém `index.html` como fonte de publicação.
5. Salve e aguarde a publicação do GitHub Pages.
6. Abra o endereço publicado e importe o arquivo confidencial diretamente pelo navegador.

## 11. Arquivos do projeto

```text
simulador-comissoes-assessorias/
├── index.html
├── README.md
├── .gitignore
├── .nojekyll
├── css/
│   └── styles.css
└── js/
    └── app.js
```
