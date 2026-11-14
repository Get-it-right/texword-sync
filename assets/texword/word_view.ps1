# Save only the insertion point/selection; never add bookmarks or edit content.
function Get-TexwordView($Document) {
 $selection=$Document.ActiveWindow.Selection
 return @{
  document=[string]$Document.FullName
  selection_start=[int]$selection.Start;selection_end=[int]$selection.End
  page=[int]$selection.Information(3)
 }
}

function Save-TexwordView($Document,[string]$Path) {
 Get-TexwordView $Document | ConvertTo-Json | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Restore-TexwordView($Document,$State) {
 if (-not $State -or $State.document -ne $Document.FullName) {return $null}
 $last=[Math]::Max(0,$Document.Content.End-1)
 $start=[Math]::Max(0,[Math]::Min([int]$State.selection_start,$last))
 $end=[Math]::Max($start,[Math]::Min([int]$State.selection_end,$last))
 $range=$Document.Range($start,$end)
 $Document.Activate()
 $range.Select()
 $Document.ActiveWindow.ScrollIntoView($range,$true)
 return @{requested_page=[int]$State.page;restored_page=[int]$range.Information(3);
          selection_start=$start;selection_end=$end;clamped=($start -ne $State.selection_start -or $end -ne $State.selection_end)}
}
function Invoke-TexwordCom {
 param([Parameter(Mandatory=$true)][scriptblock]$Operation,[int]$TimeoutMs=10000)
 $deadline=[DateTime]::UtcNow.AddMilliseconds($TimeoutMs)
 while($true){
  try{return (& $Operation)}
  catch{
   $cause=$_.Exception
   while($cause.InnerException){$cause=$cause.InnerException}
   if($cause.HResult -notin @(-2147418111,-2147417846) -or [DateTime]::UtcNow -ge $deadline){throw}
   Start-Sleep -Milliseconds 250
  }
 }
}

# File monikers identify the owning document even with several Word instances.
# Enumerating the ROT is read-only; unlike BindToMoniker it cannot open a file.
function Find-TexwordDocument([string]$Path) {
 if(-not ('TexwordRunningDocuments' -as [type])){
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class TexwordRunningDocuments {
 [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved,out IRunningObjectTable table);
 [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved,out IBindCtx context);
 public static object Find(string path) {
  IRunningObjectTable table; IBindCtx context; IEnumMoniker iterator;
  Marshal.ThrowExceptionForHR(GetRunningObjectTable(0,out table));
  Marshal.ThrowExceptionForHR(CreateBindCtx(0,out context));
  table.EnumRunning(out iterator); var item=new IMoniker[1];
  try {
   while(iterator.Next(1,item,IntPtr.Zero)==0) {
    try {
     string name;item[0].GetDisplayName(context,null,out name);
     if(String.Equals(name.TrimStart('!').Replace('/','\\'),path.Replace('/','\\'),StringComparison.OrdinalIgnoreCase)) {
      object document;table.GetObject(item[0],out document);return document;
     }
    } catch(COMException) {} finally {Marshal.ReleaseComObject(item[0]);}
   }
   return null;
  } finally {Marshal.ReleaseComObject(iterator);Marshal.ReleaseComObject(context);Marshal.ReleaseComObject(table);}
 }
}
'@
 }
 $found=[TexwordRunningDocuments]::Find([IO.Path]::GetFullPath($Path))
 if($found){return $found}
 $active=$null
 try{$active=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')}
 catch{
  $cause=$_.Exception;while($cause.InnerException){$cause=$cause.InnerException}
  if($cause.HResult -ne -2147221021){throw}
 }
 if($active){foreach($item in $active.Documents){if($item.FullName -eq $Path){return $item}}}
 return $null
}
