# Test script for Aqon CLI
Write-Host "Testing Aqon CLI" -ForegroundColor Green

# Create test directories
$inputDir = ".\test_input"
$outputDir = ".\test_output"

# Create directories if they don't exist
if (-not (Test-Path $inputDir)) {
    New-Item -ItemType Directory -Path $inputDir | Out-Null
    Write-Host "Created input directory: $inputDir" -ForegroundColor Yellow
}

if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
    Write-Host "Created output directory: $outputDir" -ForegroundColor Yellow
}

# Test the convert command
Write-Host "`nTesting 'convert' command:" -ForegroundColor Cyan
Write-Host "cargo run -- convert --input $inputDir --output $outputDir"
cargo run -- convert --input $inputDir --output $outputDir

# Test the convert command with file type filter
Write-Host "`nTesting 'convert' command with file type filter:" -ForegroundColor Cyan
Write-Host "cargo run -- convert --input $inputDir --output $outputDir --type docx"
cargo run -- convert --input $inputDir --output $outputDir --type docx

# Test the watch command (will run for 10 seconds)
Write-Host "`nTesting 'watch' command (will run for 10 seconds):" -ForegroundColor Cyan
Write-Host "cargo run -- watch --input $inputDir --output $outputDir"
$watchProcess = Start-Process -FilePath "cargo" -ArgumentList "run", "--", "watch", "--input", $inputDir, "--output", $outputDir -PassThru

# Wait for 10 seconds
Write-Host "Waiting for 10 seconds..." -ForegroundColor Yellow
Start-Sleep -Seconds 10

# Stop the watch process
Stop-Process -Id $watchProcess.Id -Force
Write-Host "Stopped watch process" -ForegroundColor Yellow

Write-Host "`nAll tests completed!" -ForegroundColor Green