# Gerador de PowerPoint — Live Programa Neurobalance

Monta um arquivo `.pptx` a partir das imagens da pasta
`C:\Users\glowy\Downloads\Live_programaNeurobalance`, **uma imagem por slide,
na sequência dos nomes dos arquivos**.

> **Por que um script?** A sessão do Claude roda num container na nuvem, isolado
> do seu computador — ela não enxerga o seu disco `C:`. O script roda aí na sua
> máquina, onde as imagens estão, e gera o PowerPoint direto na mesma pasta.

---

## Opção 1 — PowerShell (você tem o Office instalado)

Não precisa instalar nada. Usa o próprio PowerPoint para montar o arquivo.

1. Baixe o `Criar-PPTX.ps1` e salve em qualquer lugar (ex.: na Área de Trabalho).
2. Abra o **PowerShell** e rode:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\caminho\ate\Criar-PPTX.ps1"
```

O PowerPoint vai abrir sozinho, montar os slides e salvar o arquivo.
Não mexa no mouse enquanto ele trabalha.

---

## Opção 2 — Python (funciona sem Office)

1. Instale o Python em <https://python.org> — na instalação, **marque
   "Add Python to PATH"**.
2. Abra o **Prompt de Comando** e rode:

```cmd
pip install python-pptx
python "C:\caminho\ate\Criar-PPTX.py"
```

---

## Resultado

Arquivo gerado na própria pasta das imagens:

```
C:\Users\glowy\Downloads\Live_programaNeurobalance\Live_Programa_Neurobalance.pptx
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

O script imprime a lista numerada na tela **antes** de montar. Confira ali se a
sequência está certa. Se algum slide estiver fora de ordem, renomeie o arquivo
com um número na frente (`01_capa.jpg`, `02_problema.jpg`, ...) e rode de novo.

Formatos aceitos: `.jpg` `.jpeg` `.png` `.gif` `.bmp` `.tif` `.tiff`
(o script PowerShell aceita também `.webp`; a versão Python não — a biblioteca
`python-pptx` não lê esse formato).
Arquivos que não são imagem (`.txt`, `.pdf`, `.mp4`...) são ignorados.

---

## Ajustes opcionais

### Usar outra pasta

```powershell
powershell -ExecutionPolicy Bypass -File "Criar-PPTX.ps1" -Pasta "C:\outra\pasta"
```

```cmd
python Criar-PPTX.py "C:\outra\pasta"
```

### Fundo branco em vez de preto

```powershell
powershell -ExecutionPolicy Bypass -File "Criar-PPTX.ps1" -Fundo branco
```

No Python: abra o arquivo e troque a linha `FUNDO = "preto"` por `FUNDO = "branco"`.

### Preencher o slide inteiro (cortando as sobras)

Útil quando todas as imagens já são 16:9 e você não quer nenhuma tarja.

```powershell
powershell -ExecutionPolicy Bypass -File "Criar-PPTX.ps1" -PreencherTela
```

No Python: troque `PREENCHER_TELA = False` por `PREENCHER_TELA = True`.

### Mudar o nome do arquivo gerado

```powershell
powershell -ExecutionPolicy Bypass -File "Criar-PPTX.ps1" -NomeArquivo "Live_Neurobalance_v2.pptx"
```

---

## Se der erro

| Erro | O que fazer |
|---|---|
| `A pasta nao foi encontrada` | Confira o caminho. Abra a pasta no Explorer, clique na barra de endereço, copie e passe em `-Pasta`. |
| `Nenhuma imagem encontrada` | As imagens podem estar numa subpasta. O script lê só o primeiro nível. |
| `Nao consegui abrir o PowerPoint` | Office não instalado ou versão do Microsoft Store (sem COM). Use a Opção 2 (Python). |
| `python-pptx nao esta instalada` | Rode `pip install python-pptx`. |
| `execução de scripts foi desabilitada` | Use o `-ExecutionPolicy Bypass` exatamente como está nos comandos acima. |
| Slides fora de ordem | Renomeie com prefixo numérico: `01_`, `02_`, `03_`... |

---

## Alternativa sem script

Se preferir não rodar nada: **anexe as imagens direto na conversa com o Claude**
e ele monta o `.pptx` e devolve o arquivo pronto.
