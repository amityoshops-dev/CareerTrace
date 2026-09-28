$ErrorActionPreference = "Stop"
$downloads = Join-Path $HOME "Downloads"
$docs = Get-ChildItem $downloads -File | Where-Object {$_.Extension -in ".docx",".pdf"} | Sort-Object LastWriteTime -Descending
if (-not $docs) { throw "No .docx or .pdf resume found in $downloads" }
Copy-Item $docs[0].FullName (Join-Path $PSScriptRoot "data\resume$($docs[0].Extension)") -Force
Write-Host "Copied latest resume: $($docs[0].FullName)"
