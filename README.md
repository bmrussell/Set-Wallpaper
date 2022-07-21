# Set-Wallpaper
Powershell script to set desktop wallpaper from various sources

## Installation
1. Install [bginfo](https://docs.microsoft.com/en-us/sysinternals/downloads/bginfo) from the SysInternals suite. This is just used as a lazy way to get the title and description into the image.
2. If you want wallpapers from [Unsplash](https://unsplash.com/):
    * [Register](https://unsplash.com/join) for a developer account
    * Cache the credentials to an encrypted file with:
        ```
            ./Get-CredentialFromFile.ps1 -File "$($env:USERPROFILE)/Documents/Unsplash.cr"
        ```
      Supplying your access key for the password when prompted
3. Copy `Wallpaper.bgi` into your Documents folder

## Running

### From PowerShell prompt

```
./Set-Wallpaper.ps1 -Source [Source]
```
Where `Source` is either `APOD` or `Unsplash`, not case sensitive.

### From Task scheduler
```
 pwsh.exe -WindowStyle Hidden -File "Set-Wallpaper.ps1 -Source [Source]"
 ```


## Todo
Remove dependency on bginfo