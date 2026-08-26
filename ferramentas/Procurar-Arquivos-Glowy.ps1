<#
    Procurar-Arquivos-Glowy.ps1
    Varre o computador atras da pasta antiga do projeto Glowy Life.

    SEGURANCA: este script SOMENTE LE. Ele nao altera, nao move e nao
    apaga nenhum arquivo. O resultado vai para um .txt na sua Area de Trabalho.

    Como usar:
      Abra o PowerShell, va ate a pasta onde salvou este arquivo e rode:
          .\Procurar-Arquivos-Glowy.ps1
      Se o Windows bloquear, rode antes:
          Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#>

$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'

$saida  = Join-Path ([Environment]::GetFolderPath('Desktop')) 'RESULTADO-BUSCA-GLOWY.txt'
$linhas = New-Object System.Collections.Generic.List[string]

function Escreve([string]$texto) {
    Write-Host $texto
    $linhas.Add($texto) | Out-Null
}

Escreve "=========================================================="
Escreve " BUSCA POR ARQUIVOS DO PROJETO GLOWY LIFE"
Escreve " Computador : $env:COMPUTERNAME"
Escreve " Usuario    : $env:USERNAME"
Escreve " Data       : $(Get-Date -Format 'dd/MM/yyyy HH:mm')"
Escreve "=========================================================="
Escreve ""

# ----- Onde procurar -------------------------------------------------
$locais = New-Object System.Collections.Generic.List[string]
$locais.Add($env:USERPROFILE) | Out-Null

$nuvens = @(
    "$env:USERPROFILE\OneDrive",
    "$env:USERPROFILE\Dropbox",
    "$env:USERPROFILE\Google Drive",
    "$env:USERPROFILE\iCloudDrive"
)
foreach ($n in $nuvens) {
    if (Test-Path $n) { $locais.Add($n) | Out-Null }
}

Get-PSDrive -PSProvider FileSystem |
    Where-Object { $_.Root -match '^[D-Z]:\\' } |
    ForEach-Object { if (Test-Path $_.Root) { $locais.Add($_.Root) | Out-Null } }

$locais = @($locais | Select-Object -Unique)

Escreve "Locais que serao varridos:"
foreach ($l in $locais) { Escreve "   - $l" }
Escreve ""
Escreve "Isso pode levar alguns minutos. Aguarde ate aparecer CONCLUIDO."
Escreve ""

# ----- Filtros -------------------------------------------------------
$palavras = @('glowy', 'ebook', 'neurobalance', 'neurakix', 'aevum',
              'landing', 'obrigado', 'implementa')

$extensoes = @('.html','.htm','.md','.docx','.doc','.pptx','.ppt','.pdf',
               '.xlsx','.zip','.txt','.json','.js','.css','.psd','.ai','.png','.jpg')

$ignorar = '\\(node_modules|\.git|AppData\\Local\\Temp|\$Recycle\.Bin|Windows\\|Program Files)'

function CasaComPalavra([string]$nome) {
    $baixo = $nome.ToLower()
    foreach ($p in $palavras) {
        if ($baixo.Contains($p)) { return $true }
    }
    return $false
}

# ----- 1) Pastas -----------------------------------------------------
Escreve "----------------------------------------------------------"
Escreve " 1) PASTAS COM NOME RELACIONADO AO PROJETO"
Escreve "----------------------------------------------------------"

$pastas = @()
foreach ($local in $locais) {
    Write-Host "   ... varrendo pastas em $local" -ForegroundColor DarkGray
    $pastas += Get-ChildItem -LiteralPath $local -Directory -Recurse -Force |
               Where-Object { (CasaComPalavra $_.Name) -and ($_.FullName -notmatch $ignorar) }
}
$pastas = @($pastas | Sort-Object FullName -Unique)

if ($pastas.Count -eq 0) {
    Escreve "   Nenhuma pasta encontrada."
} else {
    foreach ($p in $pastas) {
        $qtd = @(Get-ChildItem -LiteralPath $p.FullName -File -Recurse -Force).Count
        Escreve ""
        Escreve "   PASTA : $($p.FullName)"
        Escreve "     modificada em  : $($p.LastWriteTime.ToString('dd/MM/yyyy HH:mm'))"
        Escreve "     arquivos dentro: $qtd"
    }
}
Escreve ""

# ----- 2) Arquivos ---------------------------------------------------
Escreve "----------------------------------------------------------"
Escreve " 2) ARQUIVOS COM NOME RELACIONADO AO PROJETO"
Escreve "----------------------------------------------------------"

$arquivos = @()
foreach ($local in $locais) {
    Write-Host "   ... varrendo arquivos em $local" -ForegroundColor DarkGray
    $arquivos += Get-ChildItem -LiteralPath $local -File -Recurse -Force |
                 Where-Object {
                     ($extensoes -contains $_.Extension.ToLower()) -and
                     (CasaComPalavra $_.Name) -and
                     ($_.FullName -notmatch $ignorar)
                 }
}
$arquivos = @($arquivos | Sort-Object FullName -Unique | Sort-Object LastWriteTime -Descending)

if ($arquivos.Count -eq 0) {
    Escreve "   Nenhum arquivo encontrado."
} else {
    Escreve "   Total encontrado: $($arquivos.Count)"
    Escreve ""
    foreach ($a in $arquivos) {
        $kb = [math]::Round($a.Length / 1KB, 1)
        Escreve "   $($a.LastWriteTime.ToString('dd/MM/yyyy HH:mm'))  |  $kb KB"
        Escreve "      $($a.FullName)"
    }
}
Escreve ""

# ----- 3) Trabalho recente (antes de 24/08) --------------------------
Escreve "----------------------------------------------------------"
Escreve " 3) DOCUMENTOS MODIFICADOS ENTRE 01/07 E 24/08"
Escreve "    (qualquer nome - para achar o que voce nao lembra)"
Escreve "----------------------------------------------------------"

$de   = Get-Date '2026-07-01'
$ate  = Get-Date '2026-08-24 23:59'
$docs = @('.html','.htm','.md','.docx','.pptx','.pdf','.xlsx','.zip')

$recentes = @()
foreach ($local in $locais) {
    Write-Host "   ... varrendo periodo em $local" -ForegroundColor DarkGray
    $recentes += Get-ChildItem -LiteralPath $local -File -Recurse -Force |
                 Where-Object {
                     ($docs -contains $_.Extension.ToLower()) -and
                     ($_.LastWriteTime -ge $de) -and ($_.LastWriteTime -le $ate) -and
                     ($_.FullName -notmatch $ignorar)
                 }
}
$recentes = @($recentes | Sort-Object FullName -Unique |
                          Sort-Object LastWriteTime -Descending |
                          Select-Object -First 200)

if ($recentes.Count -eq 0) {
    Escreve "   Nenhum documento nesse periodo."
} else {
    Escreve "   Mostrando os $($recentes.Count) mais recentes:"
    Escreve ""
    foreach ($r in $recentes) {
        Escreve "   $($r.LastWriteTime.ToString('dd/MM/yyyy HH:mm'))  |  $($r.FullName)"
    }
}
Escreve ""

# ----- 4) Lixeira ----------------------------------------------------
Escreve "----------------------------------------------------------"
Escreve " 4) LIXEIRA"
Escreve "----------------------------------------------------------"
try {
    $shell   = New-Object -ComObject Shell.Application
    $lixeira = $shell.Namespace(0xA)
    $itens   = @($lixeira.Items() | Where-Object { CasaComPalavra $_.Name })
    if ($itens.Count -eq 0) {
        Escreve "   Nada relacionado ao projeto na lixeira."
    } else {
        foreach ($i in $itens) {
            Escreve "   $($i.Name)"
            Escreve "      origem: $($lixeira.GetDetailsOf($i, 1))"
        }
    }
} catch {
    Escreve "   Nao foi possivel ler a lixeira."
}
Escreve ""

# ----- Fim -----------------------------------------------------------
Escreve "=========================================================="
Escreve " CONCLUIDO"
Escreve " Resultado salvo em:"
Escreve " $saida"
Escreve "=========================================================="

$linhas | Out-File -FilePath $saida -Encoding UTF8

Write-Host ""
Write-Host "Pronto! Abrindo o resultado..." -ForegroundColor Green
Start-Process notepad.exe $saida
