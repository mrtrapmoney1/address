<#
    Test-Syntax.ps1 - every shipped file must parse, and must parse on
    Windows PowerShell 5.1 specifically.

    This suite exists because of a whole class of bug that no amount of
    careful reading catches and that only shows up on the machine the tool is
    actually deployed to:

      1. A non-ASCII character in a file with no byte-order mark. 5.1 reads
         such a file as the machine's ANSI code page, so a smart quote in a
         comment - the kind an editor inserts silently - can stop the entire
         script from parsing, on some machines and not others.

      2. PowerShell 7 syntax. The ternary operator, ?? , ?. , && and || are
         all perfectly valid in the pwsh most development happens in, and all
         of them are syntax errors in the Windows PowerShell that ships with
         Windows. A file using one parses cleanly here and fails there.
#>

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'OcrTest.ps1')

$root = Split-Path -Parent $here
$files = @(Get-ChildItem -LiteralPath $root -Filter '*.ps1' -File) +
         @(Get-ChildItem -LiteralPath $here -Filter '*.ps1' -File)

Start-OcrSection ('Parsing ' + $files.Count + ' PowerShell file(s)')

foreach ($file in $files) {
    $errors = $null
    $tokens = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile(
        $file.FullName, [ref] $tokens, [ref] $errors)

    if ($errors -and $errors.Count -gt 0) {
        Assert-OcrTrue $false ($file.Name + ' line ' + $errors[0].Extent.StartLineNumber +
                               ': ' + $errors[0].Message)
    } else {
        Assert-OcrTrue $true ($file.Name + ' parses')
    }
}

Start-OcrSection 'Every file is pure ASCII'

foreach ($file in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $bad = -1
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -gt 127) { $bad = $i; break }
    }
    if ($bad -ge 0) {
        # Report the line, because "byte 4823" is not actionable.
        $text = [System.Text.Encoding]::ASCII.GetString($bytes, 0, $bad)
        $line = ($text -split "`n").Count
        Assert-OcrTrue $false ($file.Name + ': non-ASCII byte 0x' +
                               $bytes[$bad].ToString('X2') + ' at line ' + $line)
    } else {
        Assert-OcrTrue $true ($file.Name + ' is ASCII')
    }
}

Start-OcrSection 'No PowerShell 7-only syntax'

<#
    The AST is used rather than a text search, so an operator mentioned in a
    comment or inside a string is not reported. A regex over the source would
    flag this very file.
#>
foreach ($file in $files) {
    $errors = $null
    $tokens = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $file.FullName, [ref] $tokens, [ref] $errors)
    if ($errors -and $errors.Count -gt 0) { continue }

    $found = New-Object System.Collections.ArrayList

    # Ternary and null-coalescing appear as their own AST node types in PS7.
    # Looking the types up by name means this file still runs under 5.1,
    # where those types do not exist at all.
    $ternaryType = 'System.Management.Automation.Language.TernaryExpressionAst'
    $coalesceTypes = @(
        'System.Management.Automation.Language.AssignmentStatementAst'
    )

    $all = $ast.FindAll({ $true }, $true)
    foreach ($node in $all) {
        $typeName = $node.GetType().FullName
        if ($typeName -eq $ternaryType) {
            $null = $found.Add('ternary ? : at line ' + $node.Extent.StartLineNumber)
        }
    }

    # && and || between commands are pipeline chains, also PS7-only.
    foreach ($node in $all) {
        if ($node.GetType().Name -eq 'PipelineChainAst') {
            $null = $found.Add('pipeline chain at line ' + $node.Extent.StartLineNumber)
        }
    }

    # ?? and ?. are tokens, so the token stream is the reliable place to see them.
    foreach ($token in $tokens) {
        if ($token.Kind -eq 'QuestionQuestion' -or $token.Kind -eq 'QuestionQuestionEquals' -or
            $token.Kind -eq 'QuestionDot' -or $token.Kind -eq 'QuestionLBracket') {
            $null = $found.Add($token.Text + ' at line ' + $token.Extent.StartLineNumber)
        }
    }

    if ($found.Count -gt 0) {
        Assert-OcrTrue $false ($file.Name + ' uses PowerShell 7-only syntax: ' + ($found -join ', '))
    } else {
        Assert-OcrTrue $true ($file.Name + ' is 5.1-compatible')
    }
}

Start-OcrSection 'The JavaScript half is present'

$jsRoot = Join-Path $root 'js'
foreach ($name in @('ocr-cli.js', 'ocr-segment.js', 'package.json')) {
    Assert-OcrTrue (Test-Path -LiteralPath (Join-Path $jsRoot $name)) ($name + ' exists')
}
foreach ($name in @('png.js', 'xml.js', 'preprocess.js', 'segment.js', 'engine.js', 'emit.js', 'imagefile.js')) {
    Assert-OcrTrue (Test-Path -LiteralPath (Join-Path $jsRoot ('lib/' + $name))) ('lib/' + $name + ' exists')
}

Complete-OcrSuite 'Test-Syntax'
