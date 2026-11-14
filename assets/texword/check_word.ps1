param([Parameter(Mandatory=$true)][string]$ProjectPath,[switch]$CloseAfterCheck)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word_view.ps1')
$viewPath=Join-Path $ProjectPath 'word_view.json'
$config=Get-Content -LiteralPath (Join-Path $ProjectPath 'workflow.json') -Raw -Encoding UTF8|ConvertFrom-Json
$doc=Find-TexwordDocument ([string]$config.word)
$app=$null
if($doc){$app=$doc.Application}
else{try{$app=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')}catch{$app=$null}}
if(-not $app){$app=New-Object -ComObject Word.Application}
Invoke-TexwordCom {$app.Visible=$true} | Out-Null
$script:documents=$null
Invoke-TexwordCom {$script:documents=$app.Documents} | Out-Null
if(-not $documents){throw 'Word document collection unavailable'}
$view=$null
if($doc){$view=Get-TexwordView $doc}
elseif(Test-Path -LiteralPath $viewPath){$view=Get-Content -LiteralPath $viewPath -Raw -Encoding UTF8|ConvertFrom-Json}
if(-not $doc){
 $doc=Invoke-TexwordCom { $documents.Open([string]$config.word,$false,$false) }
}
$hidden=0
foreach($control in $doc.ContentControls){
 if($control.Tag -like 'texword:*'){
  $control.Appearance=2 # wdContentControlHidden: retain mapping without visible tags/borders.
  $control.Title=''
  $hidden++
 }
}
$tracking=[bool]$doc.TrackRevisions
$pendingRevisions=[int]$doc.Revisions.Count
$skippedPendingFields=$false
try{
 if($tracking){$doc.TrackRevisions=$false}
 if($config.native_review -and $pendingRevisions -gt 0){
  # Word can move cached REF/SEQ results outside tracked insertions during a
  # bulk update. The generated caches stay intact until decisions are settled.
  $updated=0;$skippedPendingFields=$true
 }else{$updated=Invoke-TexwordCom {$doc.Fields.Update()}}
}finally{if($tracking){$doc.TrackRevisions=$true}}
# Page count and zoom belong to the user's Word window, not synchronization.
Invoke-TexwordCom {$doc.Save()} | Out-Null
$position=Restore-TexwordView $doc $view
if(Test-Path -LiteralPath $viewPath){Remove-Item -LiteralPath $viewPath}
@{path=$doc.FullName;saved=$doc.Saved;fields=$doc.Fields.Count;math=$doc.OMaths.Count;tables=$doc.Tables.Count;sections=$doc.Sections.Count;comments=$doc.Comments.Count;revisions=$doc.Revisions.Count;track_revisions=$doc.TrackRevisions;controls=$doc.ContentControls.Count;hidden_source_controls=$hidden;field_update_result=$updated;field_update_skipped_pending_revisions=$skippedPendingFields;position=$position}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $ProjectPath 'word_check.json') -Encoding UTF8
if($CloseAfterCheck){$doc.Close(0)}else{$doc.Activate()}
