$ErrorActionPreference = "Stop"

$token = $env:GITHUB_TOKEN
if (-not $token) {
    Write-Host "Error: GITHUB_TOKEN environment variable is not set."
    Write-Host "Please set it before running this script: `$env:GITHUB_TOKEN = 'your_token_here'"
    exit 1
}

$username = "1755311380"
$repoName = "WpsAgent-Deploy-choose-0213"
$description = "WpsAgent Deploy Project"

$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($username):$($token)"))
$headers = @{
    Authorization = "Basic $base64AuthInfo"
}

$body = @{
    name = $repoName
    description = $description
    private = $false
} | ConvertTo-Json

try {
    Write-Host "Creating GitHub repository: $repoName..."
    $response = Invoke-RestMethod -Uri "https://api.github.com/user/repos" -Method Post -Headers $headers -Body $body -ContentType "application/json"
    
    Write-Host "Repository created successfully!"
    Write-Host "Repository URL: $($response.html_url)"
    Write-Host "Clone URL: $($response.clone_url)"
    
    $response.clone_url
}
catch {
    Write-Host "Error creating repository: $($_.Exception.Message)"
    Write-Host "Status Code: $($_.Exception.Response.StatusCode.value__)"
    Write-Host "Response: $($_.Exception.Response.GetResponseStream() | New-Object System.IO.StreamReader | ReadToEnd)"
    exit 1
}