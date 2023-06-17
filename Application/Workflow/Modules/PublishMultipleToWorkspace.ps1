$folder_path = $args[0]
$workspace_id = $args[1]
$publish_count = 1
$file_list = Get-ChildItem -Path $folder_path\* -Include *.pbix
$file_count = $file_list.Count

try
{
    Login-PowerBIServiceAccount -ErrorAction Stop
}
catch
{
    Write-Output "Login failed."
    Write-Error "$_"
    exit 1
}

Write-Output "Logged in succesfully."
Write-Output "Publishing first report..."

Get-ChildItem -Path $folder_path\* -Include *.pbix |
ForEach-Object {
    New-PowerBIReport -Path $_.FullName -WorkspaceId $workspace_id -ConflictAction CreateOrOverwrite
    Write-Output ($publish_count.ToString() + " of " + $file_count.ToString() + " reports published.")
    $publish_count++
}