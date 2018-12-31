param (
    [switch] $Verbose = $false
)
try {
    if ($Verbose -eq $true)
    {
        $PSDefaultParameterValues['*:Verbose'] = $true
        $VerbosePreference = "Continue"
        $DebugPreference = "Continue"
    }

    $InformationPreference = "Continue"
    $WarningPreference = "Continue"
    $ErrorActionPreference = "Stop"

    Import-module .\buildFunctions.ps1 -Force

    Build-WriteTitle "Reading Build Properties and Environment"
    $buildProps = Build-ReadProperties
    Build-Log-Hashtable $buildProps  

    $solution = Build-GetRequiredProperty $buildProps "build_solution"
    $version = Build-GetRequiredProperty $buildProps "build_version"
    $artifactsDir = Build-GetPropertyOrDefault $buildProps "build_artifactsDir" "../artifacts" # Build Server Env
    $productionBranch = Build-GetPropertyOrDefault $buildProps "build_prod_branch" "master"
    $buildConfig = Build-GetPropertyOrDefault $buildProps "build_config" "Release"

    Build-WriteTitle "Getting git branch information"
    $commitCount = & git rev-list --count HEAD
    $branchName = Build-GetPropertyOrDefault $buildProps "GIT_BRANCH" "$(& git rev-parse --abbrev-ref HEAD)"
    $branchName = $branchName -ireplace "origin/", ""
    $isProdBuild = $branchName -eq $productionBranch
    Build-Log-Information "Git branch: $branchName, commit count: $commitCount, is production branch: $isProdBuild"

    Build-WriteTitle "Determining Version Number"
    $part3 = [math]::Floor($commitCount / 65535)
    $part4 = $commitCount % 65535    
    $versionSuffix = $(If ($isProdBuild) { "" } Else {  "-$branchName" })
    $version = "$version.$part3.$part4"
    Build-Log-Information "Version: $version$versionSuffix"

    Build-WriteTitle "Remove TestApp Projects from sln"
    dotnet sln $solution remove ../src/Simplexcel.MvcTestApp/Simplexcel.MvcTestApp.csproj
    dotnet sln $solution remove ../src/Simplexcel.TestApp/Simplexcel.TestApp.csproj

    Build-WriteTitle "dotnet restore"
    dotnet restore $solution

    Build-WriteTitle "dotnet pack"
    # if <Version> is set in the project file, this will not allow the use of a version suffix.
    # => Set <VersionPrefix>1.2.3</VersionPrefix> to only specify the version number.
    # The SDK will then create a Version property based on VersionPrefix and VersionSuffix
    dotnet pack -c $buildConfig -o $artifactsDir -p:PackageVersion=$version$versionSuffix -p:Version=$version$versionSuffix -p:AssemblyVersion=$version -p:FileVersion=$version $solution

    Build-WriteTitle "dotnet test"
    dotnet test "../src/Simplexcel.Tests" -c $buildConfig -r "$artifactsDir/testoutput"
}
finally
{
    $PSDefaultParameterValues.Remove('*:Verbose')
}
