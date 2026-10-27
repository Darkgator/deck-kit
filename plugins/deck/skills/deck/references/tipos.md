# Contrato dos tipos de slide

O deck é um arquivo JSON. Campos comuns a todo slide: `tipo` (obrigatório), `eyebrow` (chamada em caixa alta, até 48 caracteres), `nota` (rodapé do slide, até 170), `notas` (texto do apresentador, sem limite), `fonte` (vira "Fonte: ..." na nota).

```json
{
  "projeto": "Nome do projeto",
  "nome": "Diagnóstico",
  "marca": "neutro",
  "tipo_deck": "diagnostico",
  "slides": [ { "tipo": "capa", "titulo": "..." } ]
}
```

`deck.py tipos [tipo]` lista os campos de cada tipo. `deck.py exemplo kickoff` mostra um deck completo.

## Limites de caracteres

`eyebrow` 48 · `titulo` 70 · `subtitulo` 110 · `numero` 14 · `rotulo` 95 · `item` 120 · `corpo` 330 · `nota` 170 · `chip` 28 · `celula` 60 · `fato` 210 · `recomendacao` 230 · `impacto` 90 · `passo` 60.

Passou do limite: o gerador recusa e diz qual campo. A resposta é reescrever mais curto.

## Campos por tipo

| Tipo | Campos |
|---|---|
| `capa` | `titulo`, `subtitulo`, `eyebrow`, `data`, `foto` (caminho de imagem, opcional) |
| `numero` | `numero`, `rotulo`, `apoios` (até 3: `numero`, `rotulo`), `frase`, `eyebrow` |
| `cartoes` | `titulo`, `cartoes` (2 a 4: `numero`, `rotulo`), `faixa`, `alerta`, `nota` |
| `fluxo` | `titulo` (frase do slide), `trilhas` (até 3: `rotulo`, `tom` = `atual` ou `novo`, `passos` de 1 a 4, até 60 caracteres cada), `marca` (`texto`, `de`, `ate`), `destaque` (`valor`, `rotulo`; só com 2 trilhas e a segunda com até 3 passos) |
| `passos` | `titulo`, `etapas` (2 a 10: `etapa`, `detalhe`, `estado` = `feito` ou `aqui`) |
| `cronograma` | `titulo`, `etapas` (2 a 6: `etapa`, `periodo`, `entrega`, `chave` ou `cor`), `proximos` (até 3) |
| `implantacao` | `titulo`, `fatos` (até 4: `rotulo`, `valor`, `detalhe`), `etapas` (até 5: `periodo`, `etapa`, `detalhe`, `aprovacao`, `depois`), `precisa`, `nao_muda`, `nao_muda_rotulo` |
| `tabela` | `titulo`, `cabecalho`, `linhas` (até 7), `larguras` (proporções) |
| `frentes` | `titulo`, `frentes` (até 4: `numero`, `titulo`, `area`, `tipo`, `situacao`, `ok`), `formato` (`lista` ou `cartoes`) |
| `recomendacao` | `titulo`, `fato`, `recomendacao`, `impactos` (até 3), `tipo_rec`, `prioridade`, `altera_fluxo` |
| `decisao` | `titulo`, `subtitulo`, `opcoes` (3: `rotulo`, `corpo`; por padrão Seguir, Ajustar, Não seguir) |
| `lista` | `titulo`, `itens` (até 6) |
| `texto` | `titulo`, `blocos` (até 3: `rotulo`, `corpo`) |
| `divisor` | `titulo`, `numero`, `subtitulo`, `itens` (até 4) |

## Como cada tipo vira layout

O modelo da marca tem 9 layouts reais (Capa, Número, Painel e conteúdo, Decisão, Conteúdo, Números, Quatro cartões, Recomendação, Tabela e status). O texto vai para os placeholders do layout, editáveis no PowerPoint. O que o layout não tem (trilhas, listas de frentes, cronograma, etapas) é desenhado com formas nativas.

Quando o conteúdo não cabe no layout, o gerador cai num desenho sobre o layout de conteúdo, sempre com base clara:

- `cartoes` com menos de 4 cartões ou sem `faixa`: cartões desenhados.
- `decisao` com opção de mais de 18 caracteres no rótulo ou 64 no texto: três cartões grandes.

## Regras de forma

- Fonte mínima de 12 pt (o rodapé legal pode ter 10).
- Base clara. Fundo escuro com letra branca só em peça pequena (etiqueta, círculo).
- Texto alternativo na foto da capa.
- Tudo é texto editável: nada vira imagem.
