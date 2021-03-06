param(
    [parameter(position=0)]
    [string] $Path
)

"Searching for MS Word files..."
$docFiles = (Get-ChildItem $Path -Include *.docx,*.doc -Recurse)

"Converting " + $docFiles.Length + " files..."

$word = New-Object -com Word.Application

foreach ($docFile in $docFiles) {
    
    $doc = $word.Documents.Open($docFile.FullName)
    $doc.PageSetup.PaperSize = [Microsoft.Office.Interop.Word.WdPaperSize]::wdPaperA4
    
    $doc.Save()
    $doc.Close()
    
}

$word.Quit()
"Done"