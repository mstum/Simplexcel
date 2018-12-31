# [CmdletBinding()]
# -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true)

function Build-ReadProperties {
<#
.SYNOPSIS
Reads a text file with Key-Value pairs into a Hashtable

.DESCRIPTION
Expects a text file in the form

key1=value1
key2=value2

Will return a Hashtable with these associations

.EXAMPLE
PS> $props = Build-ReadProperties "propsForBuildServer.txt"
PS> Write-Host "Value of key1: $props.key1"

.LINK
ConvertFrom-StringData

.PARAMETER propFilename
The file name to read.
#>
    [CmdletBinding()]
    [OutputType([Hashtable])]
    param (
        [Parameter(Position=0,mandatory=$false)]
        [string]$propFilename="buildProperties.txt",

        [Parameter(Position=1,mandatory=$false)]
        [switch]$mergeEnvironment=$true
    )

    Process {
        $buildProps = Get-Content -Raw $propFilename | ConvertFrom-StringData

        if ($mergeEnvironment) {
            Get-ChildItem Env: | ForEach-Object -Process { $buildProps[$_.Key] = $_.Value }
        }

        return $buildProps
    }
}

function Build-WriteTitle {
    param (
        [AllowEmptyString()]
        [Parameter(Position=0,mandatory=$true)]
        [string]$message
    )
    Write-Host $message -ForegroundColor Green
}

function Build-Log-Verbose {
    param (
        [AllowEmptyString()]
        [Parameter(Position=0,mandatory=$true)]
        [string]$message
    )
    Write-Verbose $message
}

function Build-Log-Information {
    param (
        [AllowEmptyString()]
        [Parameter(Position=0,mandatory=$true,ValueFromPipeline)]
        [string]$message
    )
    Write-Information $message
}

function Build-Log-Warning {
    param (
        [AllowEmptyString()]
        [Parameter(Position=0,mandatory=$true)]
        [string]$message
    )
    Write-Warning $message
}

function Build-Log-Error {
    param (
        [AllowEmptyString()]
        [Parameter(Position=0,mandatory=$true)]
        [string]$message,
        
        [AllowEmptyString()]
        [Parameter(Position=1,mandatory=$false)]
        [string]$id="BuildError"
    )

    Write-Error $message -ErrorId $id
}

function Build-Log-Hashtable {
    param (
        [AllowEmptyString()]
        [Parameter(Position=0,mandatory=$true,ValueFromPipeline)]
        [Hashtable]$props
    )
    $props | Format-Table | Out-String | Build-Log-Information
}

function Build-GetPropertyOrDefault {
    param (
        [Parameter(Position=0,mandatory=$true)]
        [Hashtable]$props,
        
        [AllowEmptyString()]
        [Parameter(Position=1,mandatory=$true)]
        [string]$propName,
        
        [AllowEmptyString()]
        [Parameter(Position=2,mandatory=$true)]
        [string]$defaultValue
    )

    $value = $props.$propName

    If ([string]::IsNullOrWhiteSpace($value)) {
        Build-Log-Verbose "$propName was empty, returning '$defaultValue'"
        return $defaultValue
    } Else {
        Build-Log-Verbose "$propName was '$value'"
        return $value
    }
}

function Build-GetRequiredProperty {
    param (
        [Parameter(Position=0,mandatory=$true)]
        [Hashtable]$props,
        
        [AllowEmptyString()]
        [Parameter(Position=1,mandatory=$true)]
        [string]$propName
    )

    $value = $props.$propName

    If ([string]::IsNullOrWhiteSpace($value)) {
        Build-Log-Error "$propName requires a value, but none was given"
        exit

    } Else {
        return $value
    }
}