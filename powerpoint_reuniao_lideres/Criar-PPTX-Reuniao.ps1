<#
================================================================================
 Criar-PPTX-Reuniao.ps1  -  Glowy Life Nutrition
 Monta uma apresentacao PowerPoint (.pptx) a partir das imagens da pasta
 da Reuniao de Lideres, uma imagem por slide, na sequencia natural dos nomes.

 REQUISITO: Microsoft PowerPoint instalado no computador.
            (Se voce NAO tem o Office, use o script Criar-PPTX-Reuniao.py)

 COMO USAR:
   Abra o PowerShell e rode:
        powershell -ExecutionPolicy Bypass -File "Criar-PPTX-Reuniao.ps1"

   O script procura sozinho a pasta dentro de Downloads.
   Para apontar outra pasta:
        powershell -ExecutionPolicy Bypass -File "Criar-PPTX-Reuniao.ps1" -Pasta "C:\caminho\da\pasta"
================================================================================
#>

[CmdletBinding()]
param(
    # Pasta com as imagens. Vazio = procura sozinho dentro de Downloads.
    [string]$Pasta = "",

    # Nome do arquivo .pptx gerado (fica salvo dentro da mesma pasta)
    [string]$NomeArquivo = "Reuniao_de_Lideres.pptx",

    # Cor de fundo dos slides: "preto" ou "branco"
    [ValidateSet("preto", "branco")]
    [string]$Fundo = "preto",

    # Preenche o slide inteiro cortando as sobras, em vez de encaixar a imagem inteira
    [switch]$PreencherTela
)

$ErrorActionPreference = "Stop"

function Write-Passo($texto) { Write-Host "  $texto" -ForegroundColor Cyan }
function Write-Ok($texto)    { Write-Host "  OK  $texto" -ForegroundColor Green }
function Write-Erro($texto)  { Write-Host "  ERRO  $texto" -ForegroundColor Red }

Write-Host ""
Write-Host "=============================================" -ForegroundColor White
Write-Host " GLOWY LIFE - Gerador de PowerPoint" -ForegroundColor White
Write-Host " Reuniao de Lideres" -ForegroundColor White
Write-Host "=============================================" -ForegroundColor White
Write-Host ""

# --------------------------------------------------- 1. Descobre / valida pasta
# O nome da pasta tem acento ("reuniao"). Em vez de escrever o acento aqui
# dentro (que quebra dependendo da codificacao do arquivo), procuramos por
# padrao dentro da pasta Downloads.
if ([string]::IsNullOrWhiteSpace($Pasta)) {

    $downloads = Join-Path $env:USERPROFILE "Downloads"

    if (-not (Test-Path -LiteralPath $downloads)) {
        Write-Erro "Nao encontrei a pasta Downloads em: $downloads"
        Write-Host '  Rode informando o caminho: -Pasta "C:\sua\pasta"' -ForegroundColor Gray
        Read-Host "Pressione ENTER para fechar"
        exit 1
    }

    $candidatos = @(
        Get-ChildItem -LiteralPath $downloads -Directory |
        Where-Object { $_.Name -match '(?i)^reuni.*deres' }
    )

    if ($candidatos.Count -eq 0) {
        Write-Erro "Nao encontrei nenhuma pasta de reuniao de lideres em:"
        Write-Host "        $downloads" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Pastas disponiveis ali:" -ForegroundColor Gray
        Get-ChildItem -LiteralPath $downloads -Directory |
            ForEach-Object { Write-Host "      $($_.Name)" -ForegroundColor DarkGray }
        Write-Host ""
        Write-Host '  Rode informando o caminho: -Pasta "C:\sua\pasta"' -ForegroundColor Gray
        Read-Host "Pressione ENTER para fechar"
        exit 1
    }

    if ($candidatos.Count -gt 1) {
        Write-Erro "Encontrei mais de uma pasta possivel:"
        $candidatos | ForEach-Object { Write-Host "      $($_.FullName)" -ForegroundColor Yellow }
        Write-Host ""
        Write-Host '  Escolha uma e rode: -Pasta "caminho completo"' -ForegroundColor Gray
        Read-Host "Pressione ENTER para fechar"
        exit 1
    }

    $Pasta = $candidatos[0].FullName
}

if (-not (Test-Path -LiteralPath $Pasta)) {
    Write-Erro "A pasta nao foi encontrada:"
    Write-Host "        $Pasta" -ForegroundColor Yellow
    Read-Host "Pressione ENTER para fechar"
    exit 1
}

Write-Passo "Pasta: $Pasta"

# ------------------------------------------------- 2. Coleta e ordena imagens
$extensoes = @('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tif', '.tiff', '.emf', '.wmf')

# Ordenacao natural: slide2 vem antes de slide10 (o Windows ordena assim,
# mas o Sort-Object padrao nao). Numeros sao preenchidos com zeros a esquerda.
function Get-ChaveOrdenacao([string]$nome) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($m in [regex]::Matches($nome, '\d+|\D+')) {
        if ($m.Value -match '^\d+$') {
            [void]$sb.Append($m.Value.PadLeft(12, '0'))
        } else {
            [void]$sb.Append($m.Value.ToLowerInvariant())
        }
    }
    return $sb.ToString()
}

$imagens = Get-ChildItem -LiteralPath $Pasta -File |
           Where-Object { $extensoes -contains $_.Extension.ToLowerInvariant() } |
           Sort-Object @{ Expression = { Get-ChaveOrdenacao $_.Name } }

if ($imagens.Count -eq 0) {
    Write-Erro "Nenhuma imagem encontrada na pasta."
    Write-Host "  Formatos aceitos: $($extensoes -join ', ')" -ForegroundColor Gray
    Read-Host "Pressione ENTER para fechar"
    exit 1
}

Write-Ok "$($imagens.Count) imagens encontradas. Ordem dos slides:"
$n = 1
foreach ($img in $imagens) {
    Write-Host ("      {0,3}. {1}" -f $n, $img.Name) -ForegroundColor Gray
    $n++
}
Write-Host ""
Write-Host "  Confira a ordem acima." -ForegroundColor Yellow
$resposta = Read-Host "  Digite S para montar a apresentacao (ou ENTER para cancelar)"
if ($resposta -notmatch '^(?i)s') {
    Write-Host "  Cancelado. Renomeie os arquivos com prefixo numerico (01_, 02_...) e rode de novo." -ForegroundColor Gray
    exit 0
}
Write-Host ""

# ------------------------------------------------------- 3. Abre o PowerPoint
Write-Passo "Abrindo o PowerPoint..."

try {
    $ppt = New-Object -ComObject PowerPoint.Application
} catch {
    Write-Erro "Nao consegui abrir o PowerPoint."
    Write-Host "  O Microsoft PowerPoint precisa estar instalado nesta maquina." -ForegroundColor Gray
    Write-Host "  Se voce nao tem o Office, use o script Criar-PPTX-Reuniao.py (Python)." -ForegroundColor Gray
    Read-Host "Pressione ENTER para fechar"
    exit 1
}

# O PowerPoint exige janela visivel para automacao via COM
$ppt.Visible = -1   # msoTrue

$apresentacao = $ppt.Presentations.Add()

# Slide widescreen 16:9  ->  13,333 x 7,5 polegadas  ->  960 x 540 pontos
$apresentacao.PageSetup.SlideWidth  = 960
$apresentacao.PageSetup.SlideHeight = 540

$larguraSlide = $apresentacao.PageSetup.SlideWidth
$alturaSlide  = $apresentacao.PageSetup.SlideHeight

$corFundo = if ($Fundo -eq "branco") { 16777215 } else { 0 }   # RGB branco / preto

# ------------------------------------------------------- 4. Monta os slides
Write-Passo "Montando os slides... (nao mexa no mouse)"

$ppLayoutBlank = 12
$msoTrue       = -1
$msoFalse      = 0

$indice = 1
foreach ($img in $imagens) {

    $slide = $apresentacao.Slides.Add($indice, $ppLayoutBlank)

    # Fundo solido, para que as bordas fiquem com acabamento intencional
    $slide.FollowMasterBackground     = $msoFalse
    $slide.Background.Fill.Visible    = $msoTrue
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = $corFundo

    # Insere a imagem no tamanho original (Width/Height = -1)
    $forma = $slide.Shapes.AddPicture(
        $img.FullName,   # FileName
        $msoFalse,       # LinkToFile        (nao) -> imagem embutida no arquivo
        $msoTrue,        # SaveWithDocument  (sim)
        0, 0,            # Left, Top
        -1, -1           # Width, Height     (-1 = tamanho original)
    )

    $forma.LockAspectRatio = $msoTrue

    # Escala mantendo a proporcao original da imagem
    if ($PreencherTela) {
        # Cobre o slide inteiro (pode cortar as sobras)
        $escala = [Math]::Max($larguraSlide / $forma.Width, $alturaSlide / $forma.Height)
    } else {
        # Mostra a imagem inteira dentro do slide
        $escala = [Math]::Min($larguraSlide / $forma.Width, $alturaSlide / $forma.Height)
    }

    $forma.Width = $forma.Width * $escala

    # Centraliza
    $forma.Left = ($larguraSlide - $forma.Width)  / 2
    $forma.Top  = ($alturaSlide  - $forma.Height) / 2

    Write-Host ("      slide {0,3}  <-  {1}" -f $indice, $img.Name) -ForegroundColor DarkGray
    $indice++
}

# ------------------------------------------------------------- 5. Salva tudo
$caminhoSaida = Join-Path $Pasta $NomeArquivo

Write-Host ""
Write-Passo "Salvando..."

$ppSaveAsOpenXMLPresentation = 24
$apresentacao.SaveAs($caminhoSaida, $ppSaveAsOpenXMLPresentation)
$apresentacao.Close()
$ppt.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($apresentacao) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null

Write-Host ""
Write-Ok "Apresentacao criada com $($imagens.Count) slides:"
Write-Host "        $caminhoSaida" -ForegroundColor Yellow
Write-Host ""

Read-Host "Pressione ENTER para fechar"
