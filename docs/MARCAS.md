# Pacotes de marca

Uma **marca** é uma pasta com um `marca.json` e, se houver, imagens. Dela sai o **modelo** (`modelo.pptx` e `modelo.potx`): 9 layouts reais, com o tema (cores e fonte) da marca. A marca `neutro` vem embutida.

## Criar uma marca

1. Copie `plugins/deck/marcas/neutro/` para uma pasta sua, por exemplo `minha-empresa/`.
2. Edite o `marca.json`:

```json
{
  "nome": "Minha Empresa",
  "fonte": "Arial",
  "cores": {
    "primaria": "1F3A5F",
    "secundaria": "2F6DB5",
    "destaque": "D9480F",
    "aviso": "F08C00",
    "petalas": ["C92A2A", "D9480F", "F08C00", "F2C14E"],
    "fundo": "F3F1EA",
    "tinta": "1E1E1C",
    "cinza": "5B6475"
  },
  "logo": "logo.png",
  "marca_dagua": "marca-dagua.png",
  "ornamento": "ornamento.png",
  "foto_capa": "foto.jpg",
  "foto_corte": [0.0, 0.0, 0.0, 0.0],
  "descricao_foto": "Descrição da foto para leitores de tela",
  "classificacao": "USO INTERNO",
  "estilo": {"sem_travessao": false, "exigir_fonte": false, "vedados": []}
}
```

| Campo | O que é |
|---|---|
| `cores.primaria` | Títulos e elementos principais |
| `cores.secundaria` | Barras e realces |
| `cores.destaque`, `cores.aviso` | Realces de atenção (filete lateral, passos com problema) |
| `cores.petalas` | 4 cores para numeração e cartões, na ordem de uso |
| `cores.fundo` | Fundo claro de capa, número e decisão |
| `cores.tinta`, `cores.cinza` | Texto principal e secundário |
| `painel`, `suave`, `palido`, `linha` | Opcionais: se faltarem, saem da cor primária, misturada com branco |
| `logo` | PNG com fundo transparente, colorido (todos os layouts têm base clara) |
| `marca_dagua` | PNG translúcido, para o painel e o slide de número |
| `ornamento` | PNG pequeno na capa |
| `foto_capa`, `foto_corte` | Foto da capa; corte `[esquerda, topo, direita, base]` em fração. Sem foto, a capa usa um bloco de cor |
| `classificacao` | Selo no alto de cada slide (por exemplo "INTERNO"). Sem o campo, não há selo |
| `estilo` | Regras do texto: sem travessão, exigir fonte nos números, termos vedados |

Logo, foto e marca d'água são opcionais. Uma marca só com cores e fonte já funciona.

## Gerar o modelo e instalar

```
python plugins/deck/scripts/deck.py modelo minha-empresa
python plugins/deck/scripts/deck.py marca-instalar minha-empresa
python plugins/deck/scripts/deck.py marca-padrao minha-empresa
```

O modelo (`modelo.potx`) também serve no PowerPoint, sem o plugin: dois cliques abrem um deck novo com os layouts da marca.

## Onde o plugin procura marcas

1. `./.deckkit/marcas/<nome>/` na pasta do projeto
2. `~/.deckkit/marcas/<nome>/` na sua máquina
3. as embutidas no plugin

O nome pode ser dado por `--marca <nome>`, pelo campo `marca` do deck ou pela marca padrão.

## Preferências pessoais

`deck.py config` grava em `~/.deckkit/config.json` e vale sobre qualquer marca:

```
python plugins/deck/scripts/deck.py config --sem-travessao sim --autor "Seu Nome" --cedido-a "Cliente"
```

O `--autor` escreve a nota de autoria no rodapé dos slides (`© 2026 Seu Nome. Cedido a Cliente.`).

## Direitos sobre as imagens

Coloque num pacote só o que você tem direito de usar. **Não publique** logotipo, foto ou identidade de uma empresa num repositório público sem autorização dela. Pacotes de marca de uma organização devem circular pelos canais internos dela.
