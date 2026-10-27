# deck-kit

Plugin do [Claude Code](https://claude.com/claude-code) que monta **apresentações PowerPoint** a partir de conteúdo bruto (TAP, ata, transcrição, planilha, resumo), no modelo da sua marca.

Em vez de você escolher slide por slide, o agente decide o formato de cada um, valida o conteúdo, gera o `.pptx`, **olha as imagens dos slides** e corrige antes de entregar. O resultado tem layouts reais do PowerPoint e texto editável.

## Por que existe

Decks gerados por IA saem todos iguais: cartões coloridos, muito texto miúdo, o mesmo "jeitinho". Este plugin ataca as causas:

- **Modelo de marca com layouts reais**, não formas soltas sobre um slide em branco.
- **Formato escolhido pelo conteúdo**: um número que domina vira slide de número; um processo vira fluxo; uma comparação vira tabela.
- **Texto que não cabe é erro**, não fonte menor. O detalhe vai para as notas do apresentador.
- **Conferência antes de entregar**: validação, prévia em imagem, revisão estrutural e, se você pedir, um revisor independente.
- **Nunca sobrescreve**: cada geração vira `vN`, e arquivo editado à mão é protegido.

## Instalação

Pelo terminal (instala de uma vez):

```
claude plugin marketplace add Darkgator/deck-kit
claude plugin install deck@deck-kit
```

Ou dentro de uma sessão do Claude Code: `/plugin marketplace add Darkgator/deck-kit` e depois `/plugin install deck@deck-kit` (abre a aba Discover; confirme a instalação ali).

Depois, uma vez, a dependência do gerador:

```
pip install python-pptx
```

Para testar sem instalar, a partir de um clone: `claude --plugin-dir ./plugins/deck`. Para conferir a estrutura do plugin: `claude plugin validate ./plugins/deck`.

### Requisitos

- Claude Code.
- Python 3.9 ou mais novo, com `python-pptx` (traz `lxml` e `Pillow`). Testado com Python 3.14 no Windows.
- **Prévia dos slides** (para o agente olhar o resultado): PowerPoint no Windows, ou LibreOffice + `pdftoppm`. Sem nenhum dos dois, o plugin gera o arquivo mas avisa que os slides não foram vistos.
- Rode `python plugins/deck/scripts/deck.py doctor` para conferir o ambiente.

## Como usar

Peça em linguagem natural, na pasta do projeto:

> Monta o kick-off do projeto a partir deste TAP. Público: gerência. 7 slides.

> Faz o diagnóstico com estas 4 recomendações e termina com a decisão. É para diretoria, quero o revisor.

> Refaz o slide 5 com caixas em vez de linha e gera a v3.

O agente segue a skill `deck`: entende o pedido, escolhe o formato de cada slide, escreve o JSON, valida, gera, olha as imagens e entrega com um resumo do que precisa de confirmação humana.

### Três modos

| Modo | Quando | Comportamento |
|---|---|---|
| Rascunho | "rápido", "só uma v1" | Não pergunta; gera e mostra |
| Padrão | o resto | Pergunta só o que falta |
| Executivo | "para diretoria", "importante" | Todo número com fonte, revisor independente |

## O que o plugin traz

- **14 tipos de slide**: capa, número, cartões, fluxo, passos, cronograma, implantação, tabela, frentes, recomendação, decisão, lista, texto, divisor. Contrato em [`references/tipos.md`](plugins/deck/skills/deck/references/tipos.md).
- **Guia de escolha** do tipo pelo formato do conteúdo e **estruturas** de kick-off, diagnóstico e business case.
- **Marca neutra** embutida e suporte a **marcas próprias** (cores, fonte, logotipo, foto, selo de classificação). Veja [docs/MARCAS.md](docs/MARCAS.md).
- **Agente `revisor-de-slides`**: olha as imagens com contexto limpo e aponta problemas por slide.
- **Exemplos** completos e fictícios: `python plugins/deck/scripts/deck.py exemplo kickoff`.

## Privacidade

Tudo roda na sua máquina. O plugin **não faz chamadas de rede e não envia nada a lugar nenhum**: lê seus arquivos, escreve o `.pptx` e as imagens de prévia na pasta que você indicar. Quem lê os seus documentos é o agente de código que você já usa, com as regras de privacidade dele.

## Limites

- Não desenha gráficos de dados (barras, linhas). Use número ou tabela e faça o gráfico no PowerPoint.
- Não insere imagens além da foto da capa.
- Base clara em todos os layouts: fundo escuro com letra branca só em peças pequenas.
- A escolha automática do formato pode errar; por isso há validação, prévia e revisão. Os avisos existem para você ler.

## Linha de comando

```
python plugins/deck/scripts/deck.py --help
```

`doctor`, `marcas`, `tipos`, `guia`, `exemplo`, `validar`, `gerar`, `prever`, `revisar`, `verificar`, `modelo`, `marca-instalar`, `marca-padrao`, `config`.

## Testes

```
python -m unittest discover -s tests -v
```

## Licença

MIT. Copyright (c) 2026 Alexandre Barroso de Miranda. Veja [LICENSE](LICENSE).
