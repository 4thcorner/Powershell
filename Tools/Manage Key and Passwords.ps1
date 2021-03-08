Function New-ByteKey {
<#
.SYNOPSIS
    Create a Byte Array
.DESCRIPTION
    Create an New Byte Array
.PARAMETER Lenght
    Set key Lengt, Must be 16,24 or 32 (Maximal is 32)
.INPUTS
    Lenght : Interger
.OUTPUTS
    Key    : String
.EXAMPLE
    New-EncryptedString -PlainTextString "I am about to become encrypted!" -Key $Key
#>
    Param (
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeLine=$true)]
        [ValidateSet(“16”,”24”,”32”)]
        [Alias("KeyLength")]
        [INT]$Length
    )
    BEGIN  {
        
    }
    PROCESS {
        $KeyByte = New-Object Byte[] $Length # You can use 16, 24, or 32 for AES
        [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($KeyByte)
        $Key = $KeyByte -join ","
    }
    END {
        Return $Key
    }
}

Function New-EncryptedString {
<#
.SYNOPSIS
    Encrypt String with Secret Key
.DESCRIPTION
    Create an New String from input
.PARAMETER PlainTextString
    String to Encrypt
.PARAMETER Key
    Byte array
.INPUTS
    PlainTextString  : String
    Key              : Byte array
.OUTPUTS
    EncryptedString  : String
.EXAMPLE
    New-EncryptedString -PlainTextString "I am about to become encrypted!" -Key $Key
#>
    Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeLine=$true)]
        [Alias("String")]
        [String]$PlainTextString,
        [Parameter(Mandatory=$True,Position=1)]
        [Alias("Key")]
        [byte[]]$EncryptionKey
    )
    BEGIN {
    }
    PROCESS { 
        Try{
            $secureString = Convertto-SecureString $PlainTextString -AsPlainText -Force
            $EncryptedString = ConvertFrom-SecureString -SecureString $secureString -Key $EncryptionKey
        }
        Catch{
            $EncryptedString  = $null
        }
    }
    END {
        Return $EncryptedString
    }
}

Function Get-DecryptString {
<#
.SYNOPSIS
    Dencrypt String with Key
.DESCRIPTION
    Decrypt String with Key
.PARAMETER EncryptedString
    Encrypted String 
.PARAMETER Key
    Byte array
.INPUTS
    EncryptedString : String
    Key             : Byte array
.OUTPUTS
    String           : String
.EXAMPLE
    New-EncryptedString -PlainTextString "I am about to become encrypted!" -Key $Key
#>
    Param(
        [Parameter(Mandatory=$True,Position=0,ValueFromPipeLine=$true)]
        [Alias("String")]
        [String]$EncryptedString,
    
        [Parameter(Mandatory=$True,Position=1)]
        [Alias("Key")]
        [byte[]]$EncryptionKey
    )
    BEGIN {

    }
    PROCESS { 
        Try{
            $SecureString = ConvertTo-SecureString $EncryptedString -Key $EncryptionKey
            $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
            [string]$String = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
        Catch{
            $String = $null
        }
    }
    END {
        Return $String
    }
}

$Key = Read-Host "Key to use for obfuscate. Leave Empty to create New Key"

If ([string]::IsNullOrEmpty($Key)) {
    $Length = Read-Host "Length of Byte 16, 24 or 32 for AES"
    $Key = New-ByteKey $Length
    $NewKey = $Key.Split(",")
    Write-Host "Key to use for obfuscate. Save it!: $Key"
}
Else {
    $NewKey = $Key.Split(",")
}

$String = Read-Host "Enter String to obfuscate!"
$obfuscatedString = New-EncryptedString $String $NewKey
Write-host "Obfuscated: $obfuscatedString"