$file_path = $args[0]
$workspace_id = $args[1]

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
Write-Output "Publishing report..."

New-PowerBIReport -Path $file_path -WorkspaceId $workspace_id -ConflictAction CreateOrOverwrite
Write-Output ("Report published to workspace.")