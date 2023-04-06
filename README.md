# Set-Wallpaper
Powershell script to set desktop wallpaper from various sources

## Installation
1. If you want wallpapers from [Unsplash](https://unsplash.com/):
    * [Register](https://unsplash.com/join) for a developer account
    * Cache the credentials to an encrypted file with:
        ```
            ./Get-CredentialFromFile.ps1 -File "$($env:USERPROFILE)/Documents/Unsplash.cr"
        ```
      Supplying your access key for the password when prompted



2. If you haven't already allowed the running of PowerShell scripts on your system run the following command at a PowerShell prompt

    ```powershell
    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
    ```
---
## Running

### From PowerShell prompt

```
./Set-Wallpaper.ps1 -Source [Source]
```
Where `Source` is either `BING`, `APOD` or `Unsplash`, not case sensitive.

### From Task scheduler
```
 pwsh.exe -WindowStyle Hidden -File "Set-Wallpaper.ps1 -Source [Source]"
 ```


