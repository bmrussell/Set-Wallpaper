# Get credentials from an encrypted file, prompting for entry if not present
# Defaults to <logged on username>.cr in home directory

param([string]$File = "")

[PSCredential]$creds = $null
if ("" -eq $File) {
    $File = "$($Env:USERPROFILE)/$($env:USERNAME).cr"
}
if ($false -eq (Test-Path -Path $File -PathType Leaf))
{
    # File not there so get them
    $creds = Get-Credential -Message "Enter Credentials"
    # Encrypt and save
    $username = $creds.GetNetworkCredential().UserName
    $usernameEncrypted = (ConvertTo-SecureString -String $username -AsPlainText -Force)
    ConvertFrom-SecureString($usernameEncrypted) | Out-File $File

    $creds.GetNetworkCredential().Password
    $passwordEncrypted = (ConvertTo-SecureString -String $creds.GetNetworkCredential().Password -AsPlainText -Force)
    ConvertFrom-SecureString($passwordEncrypted) | Out-File $File -Append                
} else {
    $username = $null
    foreach($line in Get-Content $File) {
        if ($null -eq $username) {
            $usernameEncrypted = ConvertTo-SecureString -String $line
            $username = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($usernameEncrypted))
        } else {
            $passwordEncrypted = ConvertTo-SecureString -String $line
            $creds = New-Object System.Management.Automation.PsCredential($username, $passwordEncrypted)
        }
    }        
}
$creds
