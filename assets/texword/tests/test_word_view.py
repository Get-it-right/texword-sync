"""Opt-in integration check against desktop Word's actual selection and reload."""
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
POWERSHELL = 'C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'


class ComRetryTests(unittest.TestCase):
 def test_wrapped_busy_errors_retry_but_other_errors_and_timeout_stop(self):
  script=". '"+str(ROOT/'word_view.ps1').replace("'","''")+"'\n"+r'''
$script:attempts=0
$value=Invoke-TexwordCom {
 $script:attempts++
 if($script:attempts -lt 3){throw [Exception]::new('wrapper',[Runtime.InteropServices.COMException]::new('busy',-2147418111))}
 'ok'
}
$script:hard=0
try{Invoke-TexwordCom {$script:hard++;throw [Exception]::new('ordinary error')}}catch{}
$script:timed=0
try{Invoke-TexwordCom -TimeoutMs 0 {$script:timed++;throw [Runtime.InteropServices.COMException]::new('busy',-2147417846)}}catch{}
@{value=$value;attempts=$script:attempts;hard=$script:hard;timed=$script:timed}|ConvertTo-Json -Compress
'''
  result=subprocess.run([POWERSHELL,'-NoProfile','-Command',script],capture_output=True,text=True,check=True)
  self.assertEqual(json.loads(result.stdout),dict(value='ok',attempts=3,hard=1,timed=1))


@unittest.skipUnless(os.environ.get('TEXWORD_WORD_INTEGRATION') == '1',
                     'Set TEXWORD_WORD_INTEGRATION=1 to exercise desktop Word')
class WordViewTests(unittest.TestCase):
 def test_reload_restores_cursor_and_clamps_shorter_document(self):
  with tempfile.TemporaryDirectory(prefix='texword_view_') as temporary:
   project=Path(temporary)
   document=project/'reading.docx'
   def generate(pages):
    doc=Document()
    for page in range(1,pages+1):
     doc.add_paragraph(('Reading page %d. ' % page)*150)
     if page<pages:doc.add_page_break()
    doc.save(document)
   generate(10)
   (project/'workflow.json').write_text(json.dumps({'word':str(document)}),encoding='utf8')
   script=project/'exercise.ps1'
   script.write_text(r'''
param([string]$Project,[string]$Tools,[string]$Phase)
$ErrorActionPreference='Stop'
. (Join-Path $Tools 'word_view.ps1')
$app=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
$path=Join-Path $Project 'reading.docx'
$state=Join-Path $Project 'word_view.json'
if($Phase -eq 'capture') {
 $previous=$app.ActiveDocument.FullName
 $previous | Set-Content (Join-Path $Project 'previous.txt')
 $doc=$app.Documents.Open($path)
 $doc.Activate();$window=$doc.ActiveWindow
 $window.View.Zoom.Percentage=100
 $cursor=$doc.GoTo(1,1,7);$cursor.Select()
 $before=Get-TexwordView $doc
 if($before.page -ne 7){throw 'Fixture did not position the cursor on page 7'}
 $before.selection_start|Set-Content (Join-Path $Project 'expected.txt')
 & (Join-Path $Tools 'word_session.ps1') -ProjectPath $Project -Action Save
 $sha=[Security.Cryptography.SHA256]::Create()
 $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 try{$hash=([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-','')}finally{$stream.Dispose();$sha.Dispose()}
 & (Join-Path $Tools 'word_session.ps1') -ProjectPath $Project -Action Close -ExpectedHash $hash
} else {
 & (Join-Path $Tools 'check_word.ps1') -ProjectPath $Project
 $doc=$app.Documents.Item($path)
 $view=Get-TexwordView $doc
 $report=Get-Content (Join-Path $Project 'word_check.json') -Raw|ConvertFrom-Json
 $wanted=if($Phase -eq 'shorter'){2}else{7}
 if($view.page -ne $wanted -or $report.position.restored_page -ne $wanted){throw "Wrong restored page: $($view.page)"}
 if($Phase -eq 'reload' -and $view.selection_start -ne [int](Get-Content (Join-Path $Project 'expected.txt'))){throw 'Cursor offset changed on replacement'}
 # A check without replacement must retain the existing selection too.
 & (Join-Path $Tools 'check_word.ps1') -ProjectPath $Project
 if((Get-TexwordView $doc).page -ne $wanted){throw 'Check moved reading page'}
 Save-TexwordView $doc $state
 if($Phase -eq 'reload') {
  $saved=Get-Content $state -Raw|ConvertFrom-Json
  $saved.page=10
  $saved.selection_start=$doc.Content.End-1
  $saved.selection_end=$saved.selection_start
  $saved|ConvertTo-Json|Set-Content $state -Encoding UTF8
 }
 $doc.Close(0)
 $previous=Get-Content (Join-Path $Project 'previous.txt')
 $app.Documents.Item($previous).Activate()
}
''',encoding='utf-8-sig')
   def run(phase):
    result=subprocess.run([POWERSHELL,'-NoProfile','-File',str(script),
                           '-Project',str(project),'-Tools',str(ROOT),'-Phase',phase],
                          capture_output=True,text=True,encoding='utf8',errors='replace',timeout=90)
    self.assertEqual(result.returncode,0,result.stdout+result.stderr)
   try:
    run('capture')
    generate(10)  # Replacing the bytes discards Word's own saved reading position.
    run('reload')
    generate(2)
    run('shorter')
   finally:
    cleanup=project/'cleanup.ps1'
    cleanup.write_text("param([string]$Path)\n$a=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')\nforeach($d in @($a.Documents)){if($d.FullName -eq $Path){$d.Close(0)}}\n",encoding='utf-8-sig')
    subprocess.run([POWERSHELL,'-NoProfile','-File',str(cleanup),'-Path',str(document)],capture_output=True,timeout=20)


if __name__=='__main__':unittest.main()
