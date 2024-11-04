Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Dialog = New-Object System.Windows.Forms.OpenFileDialog
$Dialog.Title = "Choose an Excel file"
$Dialog.Filter = "Excel Files (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm"

$Selection = $Dialog.ShowDialog()

if ($Selection -eq [System.Windows.Forms.DialogResult]::OK){
    $FilePath = $Dialog.FileName
    $ZipPath = Join-Path -Path (Split-Path -Path $FilePath -Parent) -ChildPath "Unlocker_temp.zip"
    $UnzippedPath = Join-Path -Path (Split-Path -Path $FilePath -Parent) -ChildPath "Unlocker_temp"
    $FileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $FileExtension = [System.IO.Path]::GetExtension($FilePath)
    $FileNewName = "${FileName}_unlocked$FileExtension"
    Copy-Item -Path $FilePath -Destination $ZipPath -Force
    if (Test-Path $UnzippedPath){
        Remove-Item -Path $UnzippedPath -Recurse -Force
    }
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $UnzippedPath)
    $WorksheetsPath = Join-Path -Path $UnzippedPath -ChildPath "xl/worksheets"
    if (Test-Path $WorksheetsPath){
        Get-ChildItem -Path $WorksheetsPath -Filter "*.xml" | ForEach-Object{
                                                                $xmlFile = $_.FullName
                                                                $xmlContent = Get-Content -Path $xmlFile
                                                                $xmlContent = $xmlContent -replace "<sheetProtection.*?/>", ""
                                                                Set-Content -Path $xmlFile -Value $xmlContent
                                                                }
        Remove-Item -Path $ZipPath
        [System.IO.Compression.ZipFile]::CreateFromDirectory($UnzippedPath, $ZipPath)
        Remove-Item -Path $UnzippedPath -Recurse -Force
        $FileNewPath = Join-Path -Path (Split-Path -Path $FilePath -Parent) -ChildPath "$FileNewName"
        if (Test-Path $FileNewPath){
            Remove-Item -Path $FileNewPath -Force
        }
        Rename-Item -Path $ZipPath -NewName "$FileNewName"

    }else{
        Write-Output "No worksheets folder found in this Excel file.."
    }

}else{
    Write-Output "No file selected."
}
