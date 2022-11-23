param([string]$Source="APOD")

# Get the developer access key by registering for a developer account at Unsplash https://unsplash.com/join
# Supply developer access key when prompted to cache in encrypted file

# Run hidden from Task scheduler with
# pwsh.exe -WindowStyle Hidden -File "Set-WallpaperFromUnsplash.ps1"

# https://stackoverflow.com/questions/39011252/powershell-image-drawstring-centering-string-with-left-and-right-padding
[void][reflection.assembly]::loadwithpartialname("system.windows.forms")

Function AddTextToImage {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true)][String] $ImagePath,
        [Parameter(Mandatory=$true)][String] $Title,
        [Parameter()][String] $Description = $null
    )

    [Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
    $srcImg = [System.Drawing.Image]::FromFile($ImagePath)
    $outputIImg = new-object System.Drawing.Bitmap([int]($srcImg.width)),([int]($srcImg.height))
    $Image = [System.Drawing.Graphics]::FromImage($outputIImg)
    $Rectangle = New-Object Drawing.Rectangle 0, 0, $srcImg.Width, $srcImg.Height
    $Image.DrawImage($srcImg, $Rectangle, 0, 0, $srcImg.Width, $srcImg.Height, ([Drawing.GraphicsUnit]::Pixel))

    # Title
    $TitleFont = new-object System.Drawing.Font("Verdana", 18, "Bold","Pixel")
    $title_font_size = [System.Windows.Forms.TextRenderer]::MeasureText($Title, $TitleFont)
#    $title_font_sizewidth = $title_font_size.Width
#    $title_font_sizeheight = $title_font_size.Height

    $titlerect = [System.Drawing.RectangleF]::FromLTRB(0, 0, $srcImg.Width, $srcImg.Height)
    $format = [System.Drawing.StringFormat]::GenericDefault
    $format.Alignment = [System.Drawing.StringAlignment]::Near
    $format.LineAlignment = [System.Drawing.StringAlignment]::Near

    # Description
    $DescFont = new-object System.Drawing.Font("Verdana", 12, "Regular","Pixel")
    # $desc_font_size = [System.Windows.Forms.TextRenderer]::MeasureText($Description, $DescFont)
    # $desc_font_sizewidth = $desc_font_size.Width
    # $desc_font_sizeheight = $desc_font_size.Height
    $descrect = [System.Drawing.RectangleF]::FromLTRB(0, $title_font_size.Height, $srcImg.Width, $srcImg.Height)

    $srcImg.Dispose()
    
    #styling font
    $Brush = New-Object Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 255, 255, 255))

    #lets draw font
    $Image.DrawString($Title, $TitleFont, $Brush, $titlerect, $format)
    #lets draw font
    $Image.DrawString($Description, $DescFont, $Brush, $descrect, $format)


    Write-Verbose "Save and close the files"
    $outputIImg.save($ImagePath, [System.Drawing.Imaging.ImageFormat]::jpeg)
    $outputIImg.Dispose()
    
}
Function Set-WallPaper($Image, [string]$Style='Fit') {

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
    [User32Functions]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $Image, $winIniFlags) | Out-Null
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
    Invoke-WebRequest $content.urls.raw -Headers $headers -Body @{ fm = "jpg"; w = "$($screenWidth)"; q = "80" } -OutFile "$($env:TEMP)/wallpaper.jpg"
    # and description
    $sel = $content | Select-Object   @{n = "Name"; e = { $_.user.name } }, @{n = "Location"; e = { $_.location.title } }, @{n = "Description"; e = { $_.description } }
    $sel.Name | Out-File "$($env:TEMP)`\wallpapertitle.txt"
    $sel.Description | Out-File "$($env:TEMP)`\wallpaperdescription.txt"

} elseif ($Source -eq "APOD") {
    $json = curl 'https://api.nasa.gov/planetary/apod?api_key=DEMO_KEY&hd=True'
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

    Invoke-WebRequest $url -OutFile "$($env:TEMP)`\wallpaper.jpg"
    $title | Out-File "$($env:TEMP)`\wallpapertitle.txt"
    $explanation | Out-File "$($env:TEMP)`\wallpaperdescription.txt"
}

$Collections = "437035,3652377,8362253"


$Source = $Source.ToUpper()

$title = Get-Content "$($env:TEMP)`\wallpapertitle.txt"
$desc = Get-Content "$($env:TEMP)`\wallpaperdescription.txt"

AddTextToImage -ImagePath "$($env:TEMP)`\wallpaper.jpg" -Title $title -Description $desc

$Style = 'Fit'
Set-WallPaper "$($env:TEMP)/wallpaper.jpg" $Style

