param([string]$Source="APOD", [switch]$Get)

# Get the developer access key by registering for a developer account at Unsplash https://unsplash.com/join
# Supply developer access key when prompted to cache in encrypted file

# Run hidden from Task scheduler with
# pwsh.exe -WindowStyle Hidden -File "Set-WallpaperFromUnsplash.ps1"

# https://stackoverflow.com/questions/39011252/powershell-image-drawstring-centering-string-with-left-and-right-padding

[void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][reflection.assembly]::loadwithpartialname("System.Windows.Forms")

$wallpaperDownloadPath = "$($env:TEMP)`\wallpaperdownload.jpg"

$apodApiKey = 'DEMO_KEY'
if ($Env:APOD_API_KEY -ne '') {
    $apodApiKey = $Env:APOD_API_KEY
}

Function Set-WallPaper($Filename, [string]$Style='Fit') {

    # Set the style of how the wallpaper should be fitted to the desktop resolution
    $WallpaperStyle = Switch ($Style) { 
        "Fill" {"10"}
        "Fit" {"6"}
        "Stretch" {"2"}
        "Tile" {"0"}
        "Center" {"0"}
        "Span" {"22"}    
    }

    If($Style -eq "Tile") {
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force | Out-Null
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 1 -Force  | Out-Null
    }
    Else {
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force | Out-Null
        New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 0 -Force | Out-Null
    }
    if (-not ([System.Management.Automation.PSTypeName]'User32Functions').Type)
    {    
        Add-Type -IgnoreWarnings -TypeDefinition @" 
            using System; 
            using System.Runtime.InteropServices;        
            public class User32Functions
            { 
                [DllImport("User32.dll",CharSet=CharSet.Unicode)] 
                public static extern int SystemParametersInfo (Int32 uAction, Int32 uParam, String lpvParam, Int32 fuWinIni);
            }
"@
    }
    $SPI_SETDESKWALLPAPER = 0x0014
    $updateIni = 0x01
    $fireChangeEvent = 0x02
    $winIniFlags = $updateIni -bor $fireChangeEvent
    [User32Functions]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $Filename, $winIniFlags) | Out-Null
}


if ($Source -eq "UNSPLASH") {
    $creds = (Get-CredentialFromFile.ps1 -File "$($env:USERPROFILE)/Documents/Unsplash.cr")
    $accessKey = $creds.GetNetworkCredential().password
    
    $baseUrl = "https://api.unsplash.com"
    $randomPhotoUrl = "$($baseUrl)/photos/random"
    $headers = @{ "Accept-Version" = "v1"; Authorization = "Client-ID $($accessKey)" }
    $params = @{ collections = $Collections; orientation = "landscape" }

    $content = Invoke-WebRequest $randomPhotoUrl -Method Get -Headers $headers -Body $params | ConvertFrom-Json

    Add-Type -AssemblyName System.Windows.Forms
    $screenWidth = [System.Windows.Forms.Screen]::AllScreens[0].Bounds.Width

    # Get image
    Invoke-WebRequest $content.urls.raw -Headers $headers -Body @{ fm = "jpg"; w = "$($screenWidth)"; q = "80" } -OutFile $wallpaperDownloadPath
    # and description
    $sel = $content | Select-Object   @{n = "Name"; e = { $_.user.name } }, @{n = "Location"; e = { $_.location.title } }, @{n = "Description"; e = { $_.description } }
    $sel.Name | Out-File "$($env:TEMP)`\wallpapertitle.txt"
    $sel.Description | Out-File "$($env:TEMP)`\wallpaperdescription.txt"

} elseif ($Source -eq "APOD") {
    # See https://api.nasa.gov/
    $json = curl "https://api.nasa.gov/planetary/apod?api_key=$($apodApiKey)&hd=True"
    $url = ($json | ConvertFrom-Json | Select-Object hdurl).hdurl
    if ($null -eq $url) {
        $url = ($json | ConvertFrom-Json | Select-Object url).url
    }
    if ($url.Substring(0,23) -eq 'https://www.youtube.com') {
        Write-Error "Can't set wallpaper from youtube link"
        return
    }
    $explanation = ($json | ConvertFrom-Json | Select-Object explanation).explanation
    $title = ($json | ConvertFrom-Json | Select-Object title).title

    Invoke-WebRequest $url -OutFile $wallpaperDownloadPath
    $title | Out-File "$($env:TEMP)`\wallpapertitle.txt"
    $explanation | Out-File "$($env:TEMP)`\wallpaperdescription.txt"
} elseif ($Source -eq "BING") {
    $json = (curl "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US" | ConvertFrom-Json)
    Invoke-WebRequest "https://bing.com$($json.images.url)" -OutFile $wallpaperDownloadPath
    $json.images.title | Out-File "$($env:TEMP)`\wallpapertitle.txt"
    $json.images.copyright | Out-File "$($env:TEMP)`\wallpaperdescription.txt"
}

Function AddTextToImage {
    param(
        [String] $inpath,
        [String] $outpath,
        [String] $Title,
        [String] $Description = $null
    )

    [Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
    $inImg = [System.Drawing.Image]::FromFile($inpath)
    $outImg = new-object System.Drawing.Bitmap([int]($inImg.width)),([int]($inImg.height))
    $Image = [System.Drawing.Graphics]::FromImage($outImg)
    $Rectangle = New-Object Drawing.Rectangle 0, 0, $inImg.Width, $inImg.Height
    $Image.DrawImage($inImg, $Rectangle, 0, 0, $inImg.Width, $inImg.Height, ([Drawing.GraphicsUnit]::Pixel))

    $textWidth = [int]($inImg.width / 4)
    # Title
    $TitleFont = new-object System.Drawing.Font("Verdana", 18, "Bold","Pixel")
    $title_font_size = [System.Windows.Forms.TextRenderer]::MeasureText($Title, $TitleFont)

    $titlerect = [System.Drawing.RectangleF]::FromLTRB(0, 0, $textWidth, $inImg.Height)
    $format = [System.Drawing.StringFormat]::GenericDefault
    $format.Alignment = [System.Drawing.StringAlignment]::Near
    $format.LineAlignment = [System.Drawing.StringAlignment]::Near

    # Description
    $DescFont = new-object System.Drawing.Font("Verdana", 12, "Regular","Pixel")
    $descrect = [System.Drawing.RectangleF]::FromLTRB(0, $title_font_size.Height, $textWidth, $inImg.Height)

    $inImg.Dispose()
    
    #styling font
    $Brush = New-Object Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 255, 255, 255))
    $Image.DrawString($Title, $TitleFont, $Brush, $titlerect, $format)
    $Image.DrawString($Description, $DescFont, $Brush, $descrect, $format)

    $outImg.save($outpath, [System.Drawing.Imaging.ImageFormat]::jpeg)
    $outImg.Dispose()
}

Function ResizeImage() {
    param([String]$ImagePath, [String]$OutputLocation)

    
    [Int]$Quality = 90
    $img = [System.Drawing.Image]::FromFile($ImagePath)

    #Encoder parameter for image quality
    $ImageEncoder = [System.Drawing.Imaging.Encoder]::Quality
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($ImageEncoder, $Quality)

    # get codec
    $Codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object {$_.MimeType -eq 'image/jpeg'}

    # Work out the size of the image so it fits vertically 
    # between top of the screen and the taskbar
    $screenWorkingHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height        
    $screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
    $screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
    $Percent = [double]($screenWorkingHeight / $img.Height)
    $newWidth = [int] (([double]$img.Width) * $Percent)
    $newHeight = [int] (([double]$img.Height) * $Percent)

    # Create a black background the size of the screen
    # This avoids a rather nasty Windows 11 bug with virtual desktop change :
    #   if the image dimensions aren't full screen the wallpaper will pop in full width and squashed
    #   before being reszed by the Windows "Fit" to desktop
    $bmpResized = New-Object System.Drawing.Bitmap($screenWidth, $screenHeight)
    $graph = [System.Drawing.Graphics]::FromImage($bmpResized)
    $graph.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graph.Clear([System.Drawing.Color]::Black)

    # Draw the resized image centred to fit screen height
    $offset = ($screenWidth - $newWidth) / 2
    $graph.DrawImage($img, $offset, 0, $newWidth, $newHeight)

    #save to file
    $bmpResized.Save($OutputLocation, $Codec, $($encoderParams))

    # Clean Up
    $bmpResized.Dispose()
    $img.Dispose()
}


$resizedPath = "$($env:TEMP)`\wallpaper-resized.jpg"
$finalPath = "$($env:TEMP)`\wallpaper.jpg"

$Collections = "437035,3652377,8362253"
$Source = $Source.ToUpper()

$title = Get-Content "$($env:TEMP)`\wallpapertitle.txt"
$desc = Get-Content "$($env:TEMP)`\wallpaperdescription.txt"

ResizeImage $wallpaperDownloadPath $resizedPath
AddTextToImage $resizedPath $finalPath $title $desc
Remove-Item -Path $resizedPath

$Style = 'Fit'
if (-not $Get.IsPresent) {
    Set-WallPaper $finalPath $Style
}
