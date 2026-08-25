<#
================================================================================
 Criar-PPTX.ps1  -  Glowy Life Nutrition
 Monta uma apresentacao PowerPoint (.pptx) a partir das imagens de uma pasta,
 uma imagem por slide, na sequencia natural dos nomes dos arquivos.

 REQUISITO: Microsoft PowerPoint instalado no computador.
            (Se voce NAO tem o Office, use o script Criar-PPTX.py)

 COMO USAR:
   1. Clique com o botao direito neste arquivo > "Executar com o PowerShell"
   OU
   2. Abra o PowerShell e rode:
        powershell -ExecutionPolicy Bypass -File "Criar-PPTX.ps1"

   Para usar outra pasta:
        powershell -ExecutionPolicy Bypass -File "Criar-PPTX.ps1" -Pasta "C:\caminho\da\pasta"
================================================================================
#>

[CmdletBinding()]
param(
    # Pasta onde estao as imagens dos slides
    [string]$Pasta = "C:\Users\glowy\Downloads\Live_programaNeurobalance",

    # Nome do arquivo .pptx gerado (fica salvo dentro da mesma pasta)
    [string]$NomeArquivo = "Live_Programa_Neurobalance.pptx",

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
Write-Host "=============================================" -ForegroundColor White
Write-Host ""

# ---------------------------------------------------------------- 1. Validacao
if (-not (Test-Path -LiteralPath $Pasta)) {
    Write-Erro "A pasta nao foi encontrada:"
    Write-Host "        $Pasta" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Corrija o caminho e rode de novo, por exemplo:" -ForegroundColor Gray
    Write-Host '        powershell -ExecutionPolicy Bypass -File "Criar-PPTX.ps1" -Pasta "C:\sua\pasta"' -ForegroundColor Gray
    Write-Host ""
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
    Write-Host ""
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

# ------------------------------------------------------- 3. Abre o PowerPoint
Write-Passo "Abrindo o PowerPoint..."

try {
    $ppt = New-Object -ComObject PowerPoint.Application
} catch {
    Write-Erro "Nao consegui abrir o PowerPoint."
    Write-Host "  O Microsoft PowerPoint precisa estar instalado nesta maquina." -ForegroundColor Gray
    Write-Host "  Se voce nao tem o Office, use o script Criar-PPTX.py (Python)." -ForegroundColor Gray
    Write-Host ""
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
Write-Passo "Montando os slides..."

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
