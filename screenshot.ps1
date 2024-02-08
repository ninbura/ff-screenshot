param(
  [string]$relativePath = "C:/repositories/ff-screenshot",
  [string]$outputDirectory = "C:/drive/pictures/ff-screenshot",
  [string]$saveToDirectory = "y",
  [string]$copyToClipboard = "y",
  [string]$captureDevice = "Game Capture 4K60 Pro MK.2",
  [string]$inputFormat,
  [string]$resolution, #= "3840x2160",
  [string]$crop, #= "718x723x0x0",
  [string]$outputFilename = [string](Get-Date -Format "yyyy-MM-dd HH-mm-ss"),
  [string]$bypassQuit = "n"
)

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

function quit(){
  write-host('Closing program, press [Enter] to exit...') -NoNewLine
  $Host.UI.ReadLine()

  exit
}

function setRelativePath {
  if($PSScriptRoot){
    $relativePath = $PSScriptRoot
  }

  return $relativePath
}

function printConfig () {
  $printObject = @{
    "output directory" = $outputDirectory
    "save to directory" = $saveToDirectory
    "copy to clipboard" = $copyToClipboard
    "capture device" = $captureDevice
    "resolution" = $resolution -eq "" ? "default" : $resolution
    "crop" = $crop -eq "" ? "none" : $crop
    "output filename" = $outputFilename
  }

  $printTable = $printObject | Format-Table -Wrap -AutoSize | Out-String -Stream | Where-Object{$_}

  Write-Host "your configuration:" -ForegroundColor Cyan

  foreach($line in $printTable){
    Write-Host $line -ForegroundColor Magenta
  }

  Write-Host ""
}

function generateDirectory ($outputDirectory) {
  Write-Host "Verifying directory exists..."

  if(!(Test-Path -Path $outputDirectory)){
    Write-Host "Directory does not exist, creating directory..."

    New-Item -Path $outputDirectory -ItemType "directory" -Force

    if(Test-Path -Path $outputDirectory){
      Write-Host "`"$outputDirectory`" has been created."
    }
    else{
      Write-Host "`"$outputDirectory`" could not be created, see log for details."
      
      exit
    }
  }
  else{
    Write-Host "`"$outputDirectory`" already exists, program will continue."
  }

  Write-Host ""
}

function generateArgumentList ($captureDevice, $resolution, $crop, $outputDirectory, $outputFilename) {
  Write-Host "Generating argument list..."

  $argumentList = @(
    "-y",
    "-loglevel", "error",
    "-stats",
    "-f", "dshow",
    "-rtbufsize", "2147.48M"
  )

  if($colorFormat){
    $argumentList += "-input_format", $colorFormat
  }

  $argumentList += "-i", "video=`"$captureDevice`"", "-map", "0"

  if($crop){
    $resolutionArray = $resolution.Split("x")
    $resolution = [ordered]@{
      width = $resolutionArray[0];
      height = $resolutionArray[1];
    }
    $cropArray = $crop.Split("x")
    $crop = [ordered]@{
      left = $cropArray[0];
      right = $cropArray[1];
      top = $cropArray[2];
      bottom = $cropArray[3];
    }

    $resolution.width = $resolution.width - $crop.left - $crop.right
    $resolution.height = $resolution.height - $crop.top - $crop.bottom

    $argumentList += "-vf", "`"crop=$($resolution.width):$($resolution.height):$($crop.left):$($crop.top)`""
  }

  $argumentList += 
    "-pix_fmt", "yuvj444p",
    "-vframes", "1",
    "-q:v", "2",
    "`"$outputDirectory\$outputFilename.jpeg`""

  return $argumentList
}

function printArgumentList($argumentList){
  Write-Host "Argument list:"

  for ($i = 0; $i -lt $argumentList.Length; $i++) {
    if (
      $argumentList[$i] -match "-" -and 
      $i + 1 -lt $argumentList.Length -and
      $argumentList[$i + 1] -notmatch "-"
    ) {
      Write-Host "$($argumentList[$i]) " -NoNewline
    }
    else {
      Write-Host $argumentList[$i]
    }
  }

  Write-Host ""
}

function runFFmpegCommand ($argumentList, $outputFilePath) {
  Write-Host "Generating screenshot..."
  Start-Process "ffmpeg" -ArgumentList $argumentList -Wait -NoNewWindow

  if(Test-Path -Path $outputFilePath){
    Write-Host "`"$outputFilePath`" has been generated."
  }
  else{
    Write-Host "Screenshot could not be generated, see above for more details..."

    quit
  }
  
  Write-Host ""
}

function copyToClipboard ($copyToClipboard, $outputFilePath) {
  if($copyToClipboard -eq "y"){
    Write-Host "Copying screenshot to clipboard..."

    $screenshot = [System.Drawing.Image]::FromFile((Get-Item -Path $outputFilePath))
    [System.Windows.Forms.Clipboard]::SetImage($screenshot)
    $screenshot.Dispose()

    Write-Host "Picture has been copied to clipboard.`n"
  }
}

function deleteScreenshot ($saveToDirectory, $outputFilePath) {
  if($saveToDirectory -eq "n"){
    Write-Host "Save to directory is disabled, screenshot will now be deleted..."

    Remove-Item -Path "$outputFilePath" -Force

    Write-Host "$outputFilePath has been deleted.`n"
  }
}

function quitOrBypass(){
  if($bypassQuit.ToLower() -eq "n"){
    quit
  } else {
    Write-Host "Process completed, program will automatically close in 10 seconds..."
    Start-Sleep 10
    exit
  }
}

try {
  Write-Host "Program is starting..."
  $argumentList = generateArgumentList $captureDevice $resolution $crop "C:\Users\$($env:UserName)\AppData\Local\Temp" $outputFilename
  $outputFilePath = $argumentList[$argumentList.Length - 1].Replace("`"", "")
  runFFmpegCommand $argumentList $outputFilePath
  printArgumentList $argumentList
  $relativePath = setRelativePath
  printConfig
  generateDirectory $outputDirectory
  Move-Item -Path $outputFilePath -Destination $outputDirectory
  $outputFilePath = "$($outputDirectory)$($outputFilePath.Substring($outputFilePath.LastIndexOf("\")))"
  copyToClipboard $copyToClipboard $outputFilePath
  deleteScreenshot $saveToDirectory $outputFilePath
  quitOrBypass
} catch {
  Write-Host "An error occurred:" -ForegroundColor red
  Write-Host $_ -ForegroundColor red
  quit
}