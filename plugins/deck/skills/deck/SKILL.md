---
name: deck
description: Monta apresentações PowerPoint (.pptx) a partir de conteúdo bruto (TAP, ata, transcrição, planilha, resumo), no modelo da marca instalada. O agente escolhe o formato de cada slide, valida, gera, olha as imagens e corrige antes de entregar. Use quando pedirem slide, apresentação, deck, PowerPoint, kick-off, diagnóstico, business case, status report ou "monta uma apresentação".
---

# Deck

Você transforma conteúdo bruto numa apresentação que parece feita por uma pessoa de design, não por um gerador. O formato de cada slide é decisão sua, guiada pelo conteúdo. O arquivo sai sobre o **modelo da marca** (layouts reais, texto editável no PowerPoint).

Comando base (use `python`; se não existir, `py -3` ou `python3`):

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/deck.py" <comando>
```

Chamado `DECK` abaixo. Comandos: `doctor`, `marcas`, `tipos`, `guia`, `exemplo`, `validar`, `gerar`, `prever`, `revisar`, `verificar`, `modelo`, `marca-instalar`, `marca-padrao`.

## 0. Antes da primeira vez na sessão

1. `DECK doctor`. Se faltar `python-pptx`, instale com `pip install python-pptx` e diga ao usuário que fez isso.
2. `DECK marcas`. Use a marca que combina com o cliente ou a organização do usuário. Só existe a `neutro`? Pergunte se há um pacote de marca da organização (pasta com `marca.json`) para instalar com `DECK marca-instalar <pasta>`. Não invente marca.
3. Se o `doctor` avisar que não há como gerar prévia (sem PowerPoint nem LibreOffice), avise o usuário desde já: você não conseguirá olhar os slides, e ele precisa conferir no PowerPoint.

## 1. Entender o pedido

Descubra, perguntando só o que falta (no máximo 3 perguntas, numa mensagem):

- **Tipo de deck**: kick-off, diagnóstico, business case, status, encerramento ou outro (`references/estruturas.md`).
- **Público e decisão**: quem vai ver e o que precisa decidir ou entender. O título dos slides nasce daqui.
- **Fonte do conteúdo**: arquivos, transcrição, planilha. Leia tudo antes de escrever. Número que você não achou na fonte não entra.
- **Tamanho**: padrão do tipo de deck, se o usuário não disser.

Modos (pelo jeito que o usuário fala):

| Modo | Quando | Comportamento |
|---|---|---|
| Rascunho | "rápido", "rascunho", "só uma v1" | Não pergunte. Gere a v1, mostre, liste as suposições. |
| Padrão | o resto | Pergunte só o que falta. Mostre o roteiro antes de gerar se tiver mais de 8 slides. |
| Executivo | "para diretoria", "importante", "vai para o board" | Padrão mais: todo número com `fonte`, revisor independente (passo 8), sem nome de pessoa nos slides. |

## 2. Roteiro: escolha o formato de cada slide

Leia `references/guia-de-escolha.md` (tabela de decisão) e `references/estruturas.md`. O formato do conteúdo escolhe o tipo:

- um número que domina → `numero`; 2 a 4 números → `cartoes`
- processo em passos, hoje contra depois → `fluxo`; onde estamos → `passos`; datas → `cronograma`
- como implantar → `implantacao`; linhas por colunas → `tabela`
- achado, custo e ação → `recomendacao`; várias frentes → `frentes`
- decisão → `decisao`; texto corrido → só como último recurso

Escreva o roteiro em uma linha por slide: título que **afirma o ponto**, tipo escolhido, por quê. Varie os formatos (nunca três iguais seguidos). Se o conteúdo pede algo que o kit não faz (gráfico de dados, diagrama BPMN, foto no meio), diga já.

## 3. Escrever o deck (JSON)

Salve como `deck-<nome>.json` na pasta onde o resultado vai ficar. Contrato completo em `references/tipos.md`; exemplos em `DECK exemplo kickoff`, `business-case`, `diagnostico`. Regras:

- **Título afirmativo, uma ideia por slide.**
- **O detalhe vai para `notas`** (texto do apresentador). Nada de texto miúdo na tela.
- **Texto acima do limite é erro de conteúdo**: encurte. Nunca reduza a fonte.
- **Número com `fonte`** (de onde vem e o período). Sem fonte, não invente: escreva "a confirmar" na `nota` e liste no resumo final.
- **Suposição vai nas `notas` do slide**, marcada como suposição.
- Português com acentos. Se a marca proíbe travessão, o validador acusa.
- Dado sensível (pessoa exposta, fraude, dado pessoal): não entra em slide; avise o usuário fora do arquivo.

## 4. Validar e gerar

```
DECK validar deck-x.json          # corrija todos os ERROS; leia os AVISOS
DECK gerar deck-x.json --saida "<pasta>" --previa
```

O `gerar` cria `<nome> vN.pptx` (nunca sobrescreve), registra a assinatura e, com `--previa`, gera um PNG por slide e uma `folha.png` de contato.

## 5. Olhar os slides (não pule)

Leia a `folha.png` e depois cada `slide-NN.png` que pareça suspeito (todos, se forem até 8). Procure: texto cortado ou fora da caixa, área vazia demais, número diferente da fonte, rótulo sobreposto, ordem confusa, repetição de formato. Corrija no JSON e gere de novo (vira v2). **No máximo 3 voltas.** Se a prévia não existe, não diga que conferiu.

## 6. Revisão estrutural

`DECK revisar "<arquivo.pptx>"`: placeholder vazio, fonte abaixo de 12 pt, possível estouro, imagem sem texto alternativo. Corrija o que for erro.

## 7. Se o arquivo foi editado à mão

Antes de gerar de novo na mesma pasta, `DECK verificar "<pasta>"`. Arquivo editado à mão **não é regerado por cima**: a nova versão sai ao lado, e a mudança do usuário só entra se ele pedir.

## 8. Revisor independente (modo executivo)

Chame o agente `revisor-de-slides`, passando: a pasta com os PNGs, o `deck-x.json`, o objetivo e o público. Ele olha com contexto limpo e devolve problemas por slide. Confira cada achado antes de aplicar; ele pode estar errado.

## 9. Entregar

Responda em poucas linhas, nesta ordem:

1. O arquivo gerado (caminho completo) e quantos slides.
2. O que foi para as notas e por quê.
3. **O que precisa de confirmação humana**: números sem fonte, suposições, dado que você não achou.
4. Avisos que sobraram e se os slides foram vistos (prévia).
5. A próxima ação possível (ajustar, gerar versão para outro público).

Não envie, não publique, não anexe em e-mail. Entregar é dizer onde está o arquivo.

## Regras de ouro

- Nunca invente número, fonte, nome ou data.
- Nunca reduza a fonte para caber texto.
- Nunca sobrescreva versão anterior nem arquivo editado à mão.
- Base clara: se pedirem fundo escuro, explique que o modelo é claro e ofereça o que dá para fazer (destaque em etiqueta ou círculo).
- Se algo falhar (Python, dependência, prévia), diga o que falhou e o que o usuário precisa fazer. Não contorne em silêncio.
