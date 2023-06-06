param(
  [string]$outputDirectory = "C:/drive/pictures/ff-screenshot",
  [string]$saveToDirectory = "true",
  [string]$copyToClipboard = "true",
  [string]$captureDevice = "Game Capture 4K60 Pro MK.2",
  [string]$resolution, #= "3840x2160",
  [string]$crop, #= "718x723x0x0",
  [string]$outputFilename = [string](Get-Date -Format "yyyy-MM-dd HH-mm-ss")
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
  else{
    $relativePath = "C:\Drive\Programming\Production\Powershell\FFmpeg\FFSuite\FFScreenshot"
  }

  return $relativePath
}

function printConfig ($config) {
  if($saveToDirectory -eq "true"){
    $saveToDirectory = $true
  } else {
    $saveToDirectory = $false
  }

  if($copyToClipboard -eq "true"){
    $copyToClipboard = $true
  } else {
    $copyToClipboard = $false
  }

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

function generateDirectory ($defaultDirectory) {
    Write-Host "Verifying directory exists..."

    if(!(Test-Path -Path $defaultDirectory)){
        Write-Host "Directory does not exist, creating directory..."

        New-Item -Path $defaultDirectory -ItemType "directory" -Force

        if(Test-Path -Path $defaultDirectory){
            Write-Host "`"$defaultDirectory`" has been created."
        }
        else{
            Write-Host "`"$defaultDirectory`" could not be created, see log for details."
            
            exit
        }
    }
    else{
        Write-Host "`"$defaultDirectory`" already exists, program will continue."
    }

    Write-Host ""
}

Function generateArgumentList ($captureDevice, $resolution, $crop, $defaultDirectory, $outputFilename) {
    Write-Host "Generating argument list..."

    $argumentList = @(
      "-y",
      "-loglevel", "error"
      "-stats"
      "-f", "dshow",
      "-rtbufsize", "2147.48M",
      "-i", "video=`"$captureDevice`"",
      "-map", "0"
    )

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
        "`"$defaultDirectory\$outputFilename.jpeg`""

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
    if($copyToClipboard){
        Write-Host "Copying screenshot to clipboard..."

        $screenshot = [System.Drawing.Image]::FromFile((Get-Item -Path $outputFilePath))
        [System.Windows.Forms.Clipboard]::SetImage($screenshot)
        $screenshot.Dispose()

        Write-Host "Picture has been copied to clipboard.`n"
    }
}

function deleteScreenshot ($saveToDirectory, $outputFilePath) {
    if(!$saveToDirectory){
        Write-Host "Save to directory is disabled, screenshot will now be deleted..."

        Remove-Item -Path "$outputFilePath" -Force

        Write-Host "$outputFilePath has been deleted.`n"
    }
}

try{
  Write-Host "Program is starting..."
  $argumentList = generateArgumentList $captureDevice $resolution $crop "C:\Users\$($env:UserName)\AppData\Local\Temp" $outputFilename
  $outputFilePath = $argumentList[$argumentList.Length - 1].Replace("`"", "")
  runFFmpegCommand $argumentList $outputFilePath
  printArgumentList $argumentList
  $relativePath = setRelativePath
  printConfig
  generateDirectory $config.DefaultDirectory
  Move-Item -Path $outputFilePath -Destination $config.DefaultDirectory
  $outputFilePath = "$($config.DefaultDirectory)$($outputFilePath.Substring($outputFilePath.LastIndexOf("\")))"
  copyToClipboard $config.copyToClipboard $outputFilePath
  deleteScreenshot $config.SaveToDirectory $outputFilePath
  Write-Host "Process completed, program will automatically close in 10 seconds..."
  Start-Sleep 10
  exit
} catch {
  Write-Host "An error occurred:" -ForegroundColor red
  Write-Host $_ -ForegroundColor red
  quit
}