# Gerador de PowerPoint — Reunião de Líderes

Monta um arquivo `.pptx` a partir das imagens da pasta
`C:\Users\glowy\Downloads\reunião de lideres`, **uma imagem por slide, na
sequência dos nomes dos arquivos**.

> **Por que um script?** A sessão do Claude roda num container na nuvem, isolado
> do seu computador — ela não enxerga o seu disco `C:`. O script roda aí na sua
> máquina, onde as imagens estão, e gera o PowerPoint direto na mesma pasta.

Os dois scripts **acham a pasta sozinhos** dentro de `Downloads` (procuram por
qualquer pasta começando com "reuni" e terminando com "deres", com ou sem
acento). Você não precisa digitar o caminho.

---

## Opção 1 — PowerShell (você tem o Office instalado)

Não precisa instalar nada. Usa o próprio PowerPoint para montar o arquivo.

1. Baixe o `Criar-PPTX-Reuniao.ps1` e salve em qualquer lugar (ex.: Área de Trabalho).
2. Abra o **PowerShell** e rode:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\caminho\ate\Criar-PPTX-Reuniao.ps1"
```

O script lista a ordem dos slides e espera você confirmar com **S**. Depois o
PowerPoint abre sozinho e monta o arquivo. Não mexa no mouse enquanto ele trabalha.

---

## Opção 2 — Python (funciona sem Office)

1. Instale o Python em <https://python.org> — na instalação, **marque
   "Add Python to PATH"**.
2. Abra o **Prompt de Comando** e rode:

```cmd
pip install python-pptx
python "C:\caminho\ate\Criar-PPTX-Reuniao.py"
```

---

## Resultado

Arquivo gerado na própria pasta das imagens:

```
C:\Users\glowy\Downloads\reunião de lideres\Reuniao_de_Lideres.pptx
```

- Slides em **widescreen 16:9** (13,333 × 7,5 pol) — padrão de projetor e Zoom.
- Uma imagem por slide, **centralizada**, com a **proporção original preservada**.
- Fundo **preto**, para que as bordas fiquem com acabamento intencional.
- Imagens **embutidas** no arquivo — o `.pptx` funciona em qualquer computador,
  sem precisar levar a pasta de imagens junto.

---

## Ordem dos slides

Os arquivos entram na **ordem natural do nome**, igual ao Windows Explorer:

```
slide1  →  slide2  →  slide3  →  slide10  →  slide11  →  slide12
```

Ou seja, `slide10` vem **depois** de `slide2`, e não antes (que é o erro comum
da ordenação alfabética simples).

O script imprime a lista numerada na tela **antes** de montar e pede confirmação.
Confira ali se a sequência está certa. Se algum slide estiver fora de ordem,
cancele (ENTER), renomeie o arquivo com número na frente (`01_capa.jpg`,
`02_agenda.jpg`, ...) e rode de novo.

Formatos aceitos: `.jpg` `.jpeg` `.png` `.gif` `.bmp` `.tif` `.tiff`
(o script PowerShell aceita também `.webp`; a versão Python não — a biblioteca
`python-pptx` não lê esse formato, e o script avisa se encontrar algum).
Arquivos que não são imagem (`.txt`, `.pdf`, `.mp4`...) são ignorados.

---

## Ajustes opcionais

### Apontar outra pasta

```powershell
powershell -ExecutionPolicy Bypass -File "Criar-PPTX-Reuniao.ps1" -Pasta "C:\outra\pasta"
```

```cmd
python Criar-PPTX-Reuniao.py "C:\outra\pasta"
```

### Fundo branco em vez de preto

```powershell
powershell -ExecutionPolicy Bypass -File "Criar-PPTX-Reuniao.ps1" -Fundo branco
```

No Python: abra o arquivo e troque a linha `FUNDO = "preto"` por `FUNDO = "branco"`.

### Preencher o slide inteiro (cortando as sobras)

Útil quando todas as imagens já são 16:9 e você não quer nenhuma tarja.

```powershell
powershell -ExecutionPolicy Bypass -File "Criar-PPTX-Reuniao.ps1" -PreencherTela
```

No Python: troque `PREENCHER_TELA = False` por `PREENCHER_TELA = True`.

### Mudar o nome do arquivo gerado

```powershell
powershell -ExecutionPolicy Bypass -File "Criar-PPTX-Reuniao.ps1" -NomeArquivo "Reuniao_Lideres_v2.pptx"
```

---

## Se der erro

| Erro | O que fazer |
|---|---|
| `nao encontrei nenhuma pasta de reuniao de lideres` | O script lista as pastas do seu Downloads. Passe o caminho certo em `-Pasta` / como argumento. |
| `encontrei mais de uma pasta possivel` | Escolha uma e passe o caminho completo. |
| `Nenhuma imagem encontrada` | As imagens podem estar numa subpasta. O script lê só o primeiro nível. |
| `Nao consegui abrir o PowerPoint` | Office não instalado ou versão da Microsoft Store (sem COM). Use a Opção 2 (Python). |
| `python-pptx nao esta instalada` | Rode `pip install python-pptx`. |
| `execução de scripts foi desabilitada` | Use o `-ExecutionPolicy Bypass` exatamente como está nos comandos acima. |
| Slides fora de ordem | Cancele, renomeie com prefixo numérico: `01_`, `02_`, `03_`... |

---

## Alternativa sem script

Se preferir não rodar nada: **anexe as imagens direto na conversa com o Claude**
e ele monta o `.pptx` e devolve o arquivo pronto.
