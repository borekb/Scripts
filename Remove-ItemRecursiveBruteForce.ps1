<#
Helper function to delete a non-empty folder more reliably.

(c) Bernd Kriszio, http://pauerschell.blogspot.cz/2010/05/problem-with-remove-item-folder-recurse.html
#>

param(
    [Parameter(Position = 0, Mandatory=$true, ValueFromPipeline=$false)]
    [string]$folder,
    [Parameter(Position = 1, Mandatory=$False, ValueFromPipeline=$false)]
    $maxiterations = 20
)
        
$iteration = 0
while (  $iteration++ -lt $maxiterations)
{
    if (Test-Path $folder)
    {
        Start-Sleep -m 50
        Remove-Item $folder -recurse -force -EA SilentlyContinue
    }
    else
    {
        "$folder deleted in $iteration iterations"
        break
    }
}
if (Test-Path $folder)
{
    Write-Error "$folder not empty after $iteration iterations"
}

