$path = 'e:\CMX\GithubSyn-CMX\cad_plugins\convert_encoding.ps1'
$parent = Split-Path $path -Parent
$lspFile = Join-Path $parent 'JZDYZ.lsp'
$content = [System.IO.File]::ReadAllText($lspFile, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($lspFile, $content, [System.Text.Encoding]::GetEncoding('gb2312'))
Write-Host "Converted to GB2312 encoding"