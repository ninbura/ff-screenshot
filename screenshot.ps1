param(
    [string]$captureDevice = "Game Capture 4K60 Pro MK.2",
    [string]$resolution, #= "3840x2160",
    [string]$crop, #= "718x723x0x0",
    [string]$outputFileName = [string](Get-Date -Format "yyyy-MM-dd HH-mm-ss")
)

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

function quit(){
  write-host('Closing program, press [Enter] to exit...') -NoNewLine
  $Host.UI.ReadLine()

  exit
}

function SetRelativePath {
    if($PSScriptRoot){
        $relativePath = $PSScriptRoot
    }
    else{
        $relativePath = "C:\Drive\Programming\Production\Powershell\FFmpeg\FFSuite\FFScreenshot"
    }

    return $relativePath
}

function PrintConfig ($config) {
    $printTable = $config | Format-Table -Wrap -AutoSize | Out-String -Stream | Where-Object{$_}

    Write-Host "Current configuration:"

    foreach($line in $printTable){
        Write-Host $line
    }

    Write-Host ""
}

function GenerateDirectory ($defaultDirectory) {
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

Function GenerateArgumentList ($captureDevice, $resolution, $crop, $defaultDirectory, $outputFileName) {
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
        "`"$defaultDirectory\$outputFileName.jpeg`""

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

function RunFFmpegCommand ($argumentList, $outputFilePath) {
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

function CopyToClipboard ($copyToClipboard, $outputFilePath) {
    if($copyToClipboard){
        Write-Host "Copying screenshot to clipboard..."

        $screenshot = [System.Drawing.Image]::FromFile((Get-Item -Path $outputFilePath))
        [System.Windows.Forms.Clipboard]::SetImage($screenshot)
        $screenshot.Dispose()

        Write-Host "Picture has been copied to clipboard.`n"
    }
}

function DeleteScreenshot ($saveToDirectory, $outputFilePath) {
    if(!$saveToDirectory){
        Write-Host "Save to directory is disabled, screenshot will now be deleted..."

        Remove-Item -Path "$outputFilePath" -Force

        Write-Host "$outputFilePath has been deleted.`n"
    }
}

try{
  Write-Host "Program is starting..."
  $argumentList = GenerateArgumentList $captureDevice $resolution $crop "C:\Users\$($env:UserName)\AppData\Local\Temp" $outputFileName
  $outputFilePath = $argumentList[$argumentList.Length - 1].Replace("`"", "")
  RunFFmpegCommand $argumentList $outputFilePath
  printArgumentList $argumentList
  $relativePath = SetRelativePath
  $config = Get-Content -Path "$relativePath\config.json" | ConvertFrom-Json
  PrintConfig $config
  GenerateDirectory $config.DefaultDirectory
  Move-Item -Path $outputFilePath -Destination $config.DefaultDirectory
  $outputFilePath = "$($config.DefaultDirectory)$($outputFilePath.Substring($outputFilePath.LastIndexOf("\")))"
  CopyToClipboard $config.CopyToClipboard $outputFilePath
  DeleteScreenshot $config.SaveToDirectory $outputFilePath
  Write-Host "Process completed, program will automatically close in 10 seconds..."
  Start-Sleep 10
  exit
} catch {
  Write-Host "An error occurred:" -ForegroundColor red
  Write-Host $_ -ForegroundColor red
  quit
}