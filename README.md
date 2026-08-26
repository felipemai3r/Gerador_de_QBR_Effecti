# Gerador de QBR — Effecti

Aplicação local (roda 100% no navegador, sem instalar nada e sem internet) que gera o
relatório de resultados (QBR) **replicando o modelo oficial da Effecti**: as páginas
do modelo são usadas como fundo e o app sobrepõe apenas os valores dinâmicos do cliente
nas posições exatas (extraídas da camada de texto do PDF original).

As páginas do modelo estão embutidas no próprio app (`assets/pages.js`), então funciona
com **duplo clique**, offline, e o PDF exporta sem depender de servidor.

## Como usar

1. Dê **duplo clique** em `Abrir QBR.bat` (ou abra o `index.html` no navegador).
2. Preencha:
   - **Nome do cliente** e **Período** (início/fim → vira "Abril a Junho de 2026").
   - **Mensalidade** — digite (ex.: `3.336,03`) ou clique em **Buscar no CS**
     depois de carregar o CSV de Customer Success. Escolha o **período do investimento**
     para o ROI: trimestre (3 meses), 4 meses, 5 meses ou semestre (6 meses).
3. **(Opcional)** Carregue o **CSV de disputas** (`dispute_web_*.csv`) — ou **cole a tabela do
   sistema interno** e clique em *Processar dados colados*. Clientes que **não usam o robô**
   (compliance) podem pular esta etapa: as páginas de arremate, economia e ROI são omitidas.
4. **Health Score / Percentual de Usabilidade:**
   - **Cole (Ctrl+V)** o print do Health Score na área indicada (ou arraste/clique para escolher).
     O print serve de referência.
   - Informe os **4 percentuais** (Encontrar / Cadastrar / Disputar / Monitorar).
     O app desenha as **bolinhas** (anéis de dots azuis com o % no centro) na página 5,
     replicando o layout do modelo.
5. **Como a Effecti te ajudou:** informe os números do período (Capturadas / Cadastradas /
   Disputadas / Monitoradas) que aparecem na página 6.
6. (Opcional) **E-mail do Gerente de Relacionamento** para a página de contatos.
7. Clique em **Gerar prévia**. **Todas as páginas do relatório entram marcadas** por padrão.
   Cada página tem uma **caixa de seleção** — desmarque as que não quiser enviar. Há também
   **Marcar todas / Desmarcar todas** no topo da prévia. Páginas de resultado baixo recebem
   um **aviso** (veja abaixo), mas continuam marcadas: quem cria decide removê-las.
8. Clique em **Baixar PDF** ou **Baixar PowerPoint**. Só as páginas marcadas são incluídas.
   Os dois formatos usam a mesma renderização em canvas e **continuam mesmo com a aba
   minimizada** ou outro programa em foco. O PowerPoint sai em widescreen 16:9 (`.pptx`).

## O que é calculado (a partir do CSV de disputas)

| Indicador | Origem no CSV |
|---|---|
| **Valor arrematado** | soma de `Valor Total Vencidos` |
| **Retorno financeiro / Margem de lucro** | soma de `Diferença Valor Mínimo` |
| **Disputas realizadas** | número de linhas |
| **Pregões vencidos acima do valor mínimo** | linhas com `Diferença Valor Mínimo > 0` |
| **Licitações vencidas** | linhas com `Items Vencidos > 0` |
| **Itens vencidos** | soma de `Items Vencidos` |
| **Órgãos atendidos** | `Nome UASG` distintos |

## ROI (no formato do modelo)

O modelo calcula o ROI sobre o **Valor Total Arrematado em 1º lugar** vs. o investimento
no período selecionado:

```
Investimento no período     = Mensalidade × meses (3, 4, 5 ou 6)
ROI aproximado (%)           = (Valor Arrematado − Investimento) / Investimento × 100
Retorno por R$ 1 investido   = Valor Arrematado / Investimento
```

O rótulo na página de ROI do PDF acompanha a opção escolhida (trimestral, 4/5 meses ou semestral).

### Seletor de páginas e avisos de resultado baixo

A prévia mostra cada página com uma caixa de seleção. **Todas entram marcadas** por padrão;
só as marcadas vão para o PDF. Para QBRs com resultado fraco, algumas páginas recebem um
**aviso** (mas seguem marcadas — quem cria decide remover):

- **ROI < 20%** → a página de ROI (pág. 4) recebe o aviso "ROI X% abaixo de 20% — considere remover".
- **Diferença de Valor Mínimo total = R$ 0,00** (cliente venceu tudo no limite, sem economia)
  → a página de Economia/Margem (pág. 3) recebe o aviso "Sem economia (Diferença R$ 0,00) — considere remover".
- Para *valor arrematado muito baixo* e outros casos subjetivos, use o seletor manualmente.

Observações:
- Sem mensalidade informada, a página de ROI não fica disponível (não há como renderizá-la).
- **Sem CSV/dados de disputa** (cliente sem robô), as páginas **2, 3 e 4** não ficam disponíveis.

## Páginas do relatório (réplica do modelo)

Com *"Incluir todas as páginas"* marcado, saem **15 páginas** do modelo (a tela de
*Futuras Implantações* foi removida por não haver roadmap). Desmarcado, saem só as
páginas de dados: **2, 3, 4, 5 e 6**.

| Pág. | Conteúdo dinâmico |
|---|---|
| 2 | Valor arrematado, período, nº de disputas via robô |
| 3 | Margem/retorno financeiro, pregões acima do valor limite |
| 4 | ROI (investimento, valor arrematado, ROI %, retorno por R$1) — marcada por padrão; aviso se < 20%; requer dados do robô |
| 5 | Percentual de Usabilidade — **bolinhas** desenhadas a partir dos 4 % |
| 6 | Como a Effecti ajudou — 4 números do período |
| 15 | Contatos — e-mail do Gerente de Relacionamento |

## Próximos passos (planejado)

A diretoria aprovou evoluir o gerador. Cenários em avaliação:

1. **Integração com o sistema interno** (cenário preferencial) — dados automáticos, sem CSV/colar manual.
2. **Integração com BI** — alternativa viável, porém os dados ainda precisam de validação.

O app atual segue funcionando de forma autônoma (CSV, colar dados, preenchimento manual) até a integração.

## Estrutura

- `index.html` — interface
- `styles.css` — estilo da interface + posicionamento dos overlays
- `app.js` — leitura do CSV, cálculos, overlays e geração do PDF
- `assets/pages.js` — páginas do modelo embutidas (data URLs)
- `assets/fun.js` — slides de encerramento (diversão) embutidos; a opção fica **escondida** na interface (checkbox permanece no HTML, desmarcado)
- `assets/fun/` — imagens fonte dos slides de diversão (para regenerar o `fun.js`)
- `assets/modelo-p07.png` — layout atualizado de *Últimas Implantações*
- `assets/modelo-pNN.png` — páginas do modelo em PNG (fonte para `pages.js`)
- `convert.ps1` — regenera `assets/pages.js` a partir das PNGs
- `tools/` — scripts Python que recortam as fotos do monitor e geram o `fun.js`
- `vendor/` — bibliotecas (PapaParse, html2canvas, jsPDF, PptxGenJS) + fonte Poppins, offline

## Ajuste fino do posicionamento

Os textos são posicionados por coordenadas em pt (extraídas do PDF) convertidas para px.
Se algum texto ficar levemente acima/abaixo, ajuste a constante `ASC` no topo do `app.js`
(fração da altura da fonte até a baseline; padrão `0.84`). As cores estão no objeto `COR`.

## Observação sobre datas

O `dispute_web_*.csv` não possui coluna de data — por isso o período é definido manualmente
no formulário. Exporte o CSV já filtrado pelo trimestre desejado na plataforma.

Já os **dados colados do sistema interno** têm data por disputa: nesse caso o app filtra
automaticamente pelas datas dentro do período informado.
