param(
 [Parameter(Mandatory=$true)][string]$ProjectPath,
 [Parameter(Mandatory=$true)][ValidateSet('Inspect','Save','Close')][string]$Action,
 [string]$ExpectedHash
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word_view.ps1')
$viewPath=Join-Path $ProjectPath 'word_view.json'
$config=Get-Content -LiteralPath (Join-Path $ProjectPath 'workflow.json') -Raw -Encoding UTF8|ConvertFrom-Json
$doc=Find-TexwordDocument ([string]$config.word)
if($Action -eq 'Inspect'){
 $state=@{open=($null -ne $doc);saved=$null;content_sha256=$null;revisions=$null}
 if($doc){
  $state.saved=[bool]$doc.Saved
  $state.revisions=[int]$doc.Revisions.Count
  $hasher=[Security.Cryptography.SHA256]::Create()
  try{$state.content_sha256=([BitConverter]::ToString($hasher.ComputeHash([Text.Encoding]::UTF8.GetBytes([string]$doc.Content.Text)))).Replace('-','').ToLowerInvariant()}
  finally{$hasher.Dispose()}
 }
 $state|ConvertTo-Json -Compress
}elseif($Action -eq 'Save'){
 if(Test-Path -LiteralPath $viewPath){Remove-Item -LiteralPath $viewPath}
 if($doc){Save-TexwordView $doc $viewPath}
 if($doc -and -not $doc.Saved){$doc.Save()}
}else{
 if(-not $ExpectedHash){throw 'Closing requires the hash of the synchronized Word file.'}
 if($doc -and -not $doc.Saved){throw 'Word was edited during synchronization; the document remains open. Run the command again.'}
 $stream=[IO.File]::Open([string]$config.word,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 $hasher=[Security.Cryptography.SHA256]::Create()
 try{$actual=([BitConverter]::ToString($hasher.ComputeHash($stream))).Replace('-','')}
 finally{$stream.Dispose();$hasher.Dispose()}
 if($actual -ne $ExpectedHash){throw 'Word changed during synchronization; it has not been replaced. Run the command again.'}
 if($doc){Save-TexwordView $doc $viewPath;$doc.Close(0)}
}
