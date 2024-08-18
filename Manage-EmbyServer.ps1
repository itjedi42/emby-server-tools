<#
    .SYNOPSIS
    Emby Server Management Script

     .DESCRIPTION
    Installs and Updates Emby Server on Windows with IIS Reverse Proxy and SSL Certificate via Certify the Web.

    What it does:
    - Checks if installation path exists, if not, creates it.
    - Checks if 7-Zip is installed (needed to extract .7z archives of Emby Server), if its not present, it will get the latest version and install it.
    - Checks for NSSM in installation path, if not found, fetches it and gets it where it needs to be.
    - Checks if Visual C++ 2015-2022 Redistributable Runtime is installed, if not it gets latest version and installs it.
    - Checks if IIS is installed, if not, installs it and relevant features.
    - Checks if IIS Rewrite2 Module is installed, if not, it downloads and installs it.
    - Checks if IIS AAR3.0 Module is installed, if not, it downloads and installs it.
    - Checks if Certify The Web client is installed, if not it downloads and installs it.
    - Downloads latest .7z x64 version of Emby Server (will download latest beta if `-beta` is specified) and extracts it.
    - Creates a local user to use as a service account to run Emby Server.
    - Configures Emby Server to run as a service via NSSM.
    - Configures Windows Firewall rule to allow TCP port `80` and `443` in and UDP `443` in (For TLS1.3/QUIC support).
    - Stops the default IIS website.
    - Creates a new `Emby Server Reverese Proxy` site.
    - Configures IIS WebServer farm in AAR3.0.
    - Configures IIS server variables.
    - Disables IIS caching (causes weirdness with streaming).
    - Configures IIS request filtering.
    - Configures IIS headers.
    - Configures rewrite/reverse proxy rules.
    - Configures Certify the Web client to create and maintain SSL certificate.
    - Disables OCSP stapling.
    - Disables legacy TLS.
    - Configures QUICK protocol and TLS 1.3 (if Windows Server 2022 or newer).
    - Starts Emby Server as a Service.
    - Launches Emby Server WebUI in the default system browser.
    - Outputs credentials of created service account.

    The script can also be used to update Emby Server when new versions are available.

    .PARAMETER InstallationPath
    Path to install Emby-Server to.

    .PARAMETER ServiceName
    Service Name for Emby-Server.

    .PARAMETER Install
    Installs Emby-Server.

    .PARAMETER ServiceAccount
    User to run Emby-Server as.

    .PARAMETER ServiceAccountSecret
    > Password for the account to run Emby-Server As, Do not supply if using -CreateServiceAccount

    .PARAMETER CreateServiceAccount
    Superceeds ServiceAccountSecret. Creates a new service account and generates password which will be logged to the console.

    .PARAMETER CertificateContactName
    Name to use with certiify the web client as contact on certificate

    .PARAMETER CertificateContactEmail
    Email to use with certify the web as contact on certificate

    .PARAMETER ExternalHostName
    Hostname to configure reverse proxy to listen on

    .PARAMETER RestartIfNeeded
    Tells the script to restart the system if needed when installing IIS.

    .PARAMETER Beta
    Configures with latest beta version instead of latest stable.

    .PARAMETER Update
    Updates Emby-Server.

    .EXAMPLE
    .\Manage-EmbyServer.ps1 -Install -CreateServiceAccount -CertificateContactName "Admin" -CertificateContactEmal "admin@example.com" -ExternalHostName "media.example.com"

    Installs latest version and creates a service account with the default options. IIS is confgured and certify the web client is used to create and maintain SSL certificate

    .EXAMPLE
    .\Manage-EmbyServer.ps1 -Install -InstallationPath "F:\EmbyServer" -CreateServiceAccount -CertificateContactName "Admin" -CertificateContactEmal "admin@example.com" -ExternalHostName "media.example.com"

    Installs latest version and creates a service account with custom path and default options.  IIS is confgured and certify the web client is used to create and maintain SSL certificate

    .EXAMPLE
    .\Manage-EmbyServer.ps1 -Install -Beta -InstallationPath "F:\EmbyServer" -CreateServiceAccount -CertificateContactName "Admin" -CertificateContactEmal "admin@example.com" -ExternalHostName "media.example.com"

    Installs latest beta version and creates a service account with custom path and default options.  IIS is confgured and certify the web client is used to create and maintain SSL certificate

    .EXAMPLE
    .\Manage-EmbyServer.ps1 -Update

    Updates Emby Server to the latest version using the default options

    .EXAMPLE
    .\Manage-EmbyServer.ps1 -Update -InstallationPath "F:\EmbyServer"

    Updates Emby Server to the latest version using a custom path

    .EXAMPLE
    .\Manage-EmbyServer.ps1 -Update -InstallationPath "F:\EmbyServer" -ServiceName "My Server"

    Updates Emby Server to the latest version using a custom path and service name

    .EXAMPLE
    .\Manage-EmbyServer.ps1 -Update -Beta

    Updates Emby Server to latest beta version using the default options

    .EXAMPLE
    .\Manage-EmbyServer.ps1 -Update  -Beta -InstallationPath "F:\EmbyServer"

    Updates Emby Server to the latest beta version using a custom path

    .NOTES
    AUTHOR Ben Felton (itjedi42)
    VERSION 0.1
    GUID 849a4b01-d245-4a14-b682-b5c0fdaf0b09
    RELEASENOTES
    This script comes with no warranties. It should work properly, but use at your own risk.
#>

<### Requires ###>

#Requires -RunAsAdministrator


<### Parameters ###>
[CmdletBinding()]
param (
    [Parameter(ParameterSetName = "Setup")]
    [Parameter(ParameterSetName = "SetupWithAccount")]
    [Parameter(ParameterSetName = "Update")]
    [string]
    $InstallationPath = "C:\EmbyServer",

    [Parameter(ParameterSetName = "Setup")]
    [Parameter(ParameterSetName = "SetupWithAccount")]
    [Parameter(ParameterSetName = "Update")]
    [string]
    $ServiceName = "EmbyServer",

    [Parameter(ParameterSetName = "Setup", Mandatory)]
    [Parameter(ParameterSetName = "SetupWithAccount", Mandatory)]
    [switch]
    $Install,

    [Parameter(ParameterSetName = "Setup")]
    [Parameter(ParameterSetName = "SetupWithAccount")]
    [string]
    $ServiceAccount = "SvcEmbyServer",

    [Parameter(ParameterSetName = "Setup", Mandatory)]
    [string]
    $ServiceAccountSecret,

    [Parameter(ParameterSetName = "SetupWithAccount", Mandatory)]
    [switch]
    $CreateServiceAccount,

    [Parameter(ParameterSetName = "Setup", Mandatory)]
    [Parameter(ParameterSetName = "SetupWithAccount", Mandatory)]
    [string]
    $ExternalHostName,

    [Parameter(ParameterSetName = "Setup", Mandatory)]
    [Parameter(ParameterSetName = "SetupWithAccount", Mandatory)]
    [string]
    $CertificateContactName,

    [Parameter(ParameterSetName = "Setup", Mandatory)]
    [Parameter(ParameterSetName = "SetupWithAccount", Mandatory)]
    [string]
    $CertificateContactEmail,

    [Parameter(ParameterSetName = "Update")]
    [switch]
    $Update,

    [Parameter(ParameterSetName = "Setup")]
    [Parameter(ParameterSetName = "SetupWithAccount")]
    [switch]
    $RestartIfNeeded = $false,

    [Parameter(ParameterSetName = "Setup")]
    [Parameter(ParameterSetName = "SetupWithAccount")]
    [switch]
    $CreateModificationsDirectory,

    [Parameter(ParameterSetName = "Setup")]
    [Parameter(ParameterSetName = "SetupWithAccount")]
    [Parameter(ParameterSetName = "Update")]
    [switch]
    $Beta
)

<### Functions ###>
Function Format-String
{
    <#
    .SYNOPSIS
    Format and colorize text for output using 24bit color pallet

    .DESCRIPTION
    Format and colorize text for output via Jenkins and Windows 10 or newer terminals, using 24bit color pallet

    .PARAMETER String
    String to format

    .PARAMETER Color
    Color to set

    .PARAMETER Bold
    Format as Bold

    .PARAMETER Italic
    Format as Italic

    .PARAMETER Underline
    Format as Underlined

    .PARAMETER StrikeThrough
    Format as Struckthrough

    .PARAMETER Invert
    Format with inverted colors

    .PARAMETER NoReset
    Prevents reseting ANSI codes at the end of the string. For use with Write-Progress.

    .EXAMPLE
    Write-Host (Format-String -Color Red "This is red text")


    Formats string "This is red text" to display in red.

    .EXAMPLE
    Write-Host (Format-String -StrikeThrough "This text is struck through")


    Formats string "This text is struck through" with a line through it.

    .EXAMPLE
    Write-Host (Format-String -Color Sand -Invert "This is inverted text")


    Formats string "This is inverted text" with sand color to have foreground and background colors inverted.

    .EXAMPLE
    Write-Host "$(Format-String -Color White -Italic "Skittles") taste the $(Format-String -Color Red R)$(Format-String -Color DarkOrange a)$(Format-String -Color Gold i)$(Format-String -Color Green n)$(Format-String -Color Blue b)$(Format-String -Color DarkBlue o)$(Format-String -Color Purple w)"


    Formats the work "Skittles" in italics and each letter of the word Rainbow as a color in the rainbow.

    .EXAMPLE
    $max = 100
    For ($i = 0; $i -lt $max; $i++)
    {
        Write-Progress -Id 1 -Activity "$(Format-String -Color SkyBlue "Processing")$(Format-String -Color DarkOrange -NoReset)" -Status "Running" -PercentComplete ([Math]::Round((($i / $max) * 100), 2))
        Start-Sleep -Milliseconds 100
    }


    Formats the activity string to be sky blue and the progress bar and status to be dark orange.

    .EXAMPLE
    Format-String -Color Blue -NoReset
    Write-Host "This text is blue"
    Format-String


    Sets all console output to be blue until reset then writes "This text is blue" and then resets the console back to default.

    .NOTES
    Written By: Ben (itjedi42)
	Copyright 2024

    https://en.wikipedia.org/wiki/ANSI_escape_code#graphics
    #>

    Param(
        [parameter(ValueFromPipeline)]
        [string]
        $String,

        [parameter()]
        [ValidateSet($null, 'Almond', 'Aqua', 'Beige', 'Black', 'Blue', 'BlueViolet', 'Brown', 'Coral', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGray', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOlive', 'DarkOrange', 'DarkRed', 'DarkViolet', 'DarkYellow', 'Gold', 'GoldenRod', 'Gray', 'Green', 'Honeydew', 'HotPink', 'IceBlue', 'Indigo', 'Khaki', 'Lavender', 'Lemon', 'LightBlue', 'LightGray', 'LightGreen', 'LightPink', 'LightPurple', 'LimeGreen', 'Magenta', 'Maroon', 'Mint', 'Navy', 'Olive', 'Orange', 'Orchid', 'PaleGreen', 'Peach', 'Peru', 'Pink', 'Plum', 'Purple', 'Red', 'RedOrange', 'RosyBrown', 'RoyalBlue', 'Salmon', 'Sand', 'SeaGreen', 'Sienna', 'Silk', 'SkyBlue', 'Slate', 'SteelBlue', 'Tan', 'Teal', 'Turquoise', 'Violet', 'Wheat', 'White', 'Yellow', 'YellowGreen')]
        [string]
        $Color,

        [parameter()]
        [switch]
        $Bold,

        [parameter()]
        [switch]
        $Italic,

        [parameter()]
        [switch]
        $Underline,

        [parameter()]
        [switch]
        $StrikeThrough,

        [parameter()]
        [switch]
        $Invert,

        [parameter()]
        [switch]
        $NoReset
    )

    Begin
    {
        $Escape = [char]27;
        $Reset = "$Escape[0m"
        $Ansi24BitTemplate = "$Escape[38;2;{0};{1};{2}m"
        $Ansi24BitColor = [PSCustomObject]@{
            Almond      = (255, 235, 205)
            Aqua        = (127, 255, 212)
            Beige       = (245, 245, 220)
            Black       = (0, 0, 0)
            Blue        = (30, 144, 255)
            BlueViolet  = (138, 43, 226)
            Brown       = (139, 69, 19)
            Coral       = (255, 127, 80)
            Cyan        = (0, 255, 255)
            DarkBlue    = (0, 0, 255)
            DarkCyan    = (0, 139, 139)
            DarkGray    = (105, 105, 105)
            DarkGreen   = (0, 128, 0)
            DarkKhaki   = (189, 183, 107)
            DarkMagenta = (139, 0, 139)
            DarkOlive   = (85, 107, 47)
            DarkOrange  = (255, 140, 0)
            DarkRed     = (139, 0, 0)
            DarkViolet  = (148, 0, 211)
            DarkYellow  = (184, 134, 11)
            Gold        = (255, 215, 0)
            GoldenRod   = (218, 165, 32)
            Gray        = (128, 128, 128)
            Green       = (50, 205, 50)
            Honeydew    = (240, 255, 240)
            HotPink     = (255, 105, 180)
            IceBlue     = (173, 216, 230)
            Indigo      = (75, 0, 130)
            Khaki       = (240, 230, 140)
            Lavender    = (230, 230, 250)
            Lemon       = (255, 250, 205)
            LightBlue   = (135, 206, 250)
            LightGray   = (192, 192, 192)
            LightGreen  = (144, 238, 144)
            LightPink   = (255, 182, 193)
            LightPurple = (147, 112, 219)
            LimeGreen   = (0, 255, 0)
            Magenta     = (255, 0, 255)
            Maroon      = (128, 0, 0)
            Mint        = (0, 250, 154)
            Navy        = (0, 0, 128)
            Olive       = (128, 128, 0)
            Orange      = (255, 165, 0)
            Orchid      = (218, 112, 214)
            PaleGreen   = (152, 251, 152)
            Peach       = (255, 218, 185)
            Peru        = (205, 133, 63)
            Pink        = (255, 20, 147)
            Plum        = (221, 160, 221)
            Purple      = (128, 0, 128)
            Red         = (255, 0, 0)
            RedOrange   = (255, 69, 0)
            RosyBrown   = (188, 143, 143)
            RoyalBlue   = (65, 105, 225)
            Salmon      = (250, 128, 114)
            Sand        = (244, 164, 96)
            SeaGreen    = (60, 179, 113)
            Sienna      = (160, 82, 45)
            Silk        = (255, 248, 220)
            SkyBlue     = (0, 191, 255)
            Slate       = (112, 128, 144)
            SteelBlue   = (176, 196, 222)
            Tan         = (210, 180, 140)
            Teal        = (0, 128, 128)
            Turquoise   = (64, 224, 208)
            Violet      = (238, 130, 238)
            Wheat       = (245, 222, 179)
            White       = (255, 255, 255)
            Yellow      = (255, 255, 0)
            YellowGreen = (154, 205, 50)
        }
    }

    Process
    {
        $Effects = $null
        If ($Bold.IsPresent)
        {
            $Effects += "$Escape[1m"
        }

        If ($Italic.IsPresent)
        {
            $Effects += "$Escape[3m"
        }

        If ($Underline.IsPresent)
        {
            $Effects += "$Escape[4m"
        }

        If ($StrikeThrough.IsPresent)
        {
            $Effects += "$Escape[9m"
        }

        If ($Invert.IsPresent)
        {
            $Effects += "$Escape[7m"
        }

        If ($Color)
        {
            $Result = "$($Effects)$($Ansi24BitTemplate -f (($Ansi24BitColor.$Color)|ForEach-Object {$_.toString().padleft(3,"0")}))$($String)"
        }
        Else
        {
            $Result = "$($Effects)$($String)"
        }

        If (-Not($NoReset.IsPresent))
        {
            $Result += $Reset
        }

        Return $Result
    }
}

Function New-Password
{
    <#
    .SYNOPSIS
    Generates a new Password or API token

    .DESCRIPTION
    Randomly generates a new Password or API token

    .PARAMETER Length
    Length of Password to generate

    .PARAMETER IncludeNumbers
    Include numbers in password

    .PARAMETER IncludeSpecialCharacters
    Include special characters in password

    .PARAMETER Segments
    Number of segments to break password into.

    .PARAMETER APIToken
    Generate API Token instead of password

    .EXAMPLE
    New-Password -Length 12 -IncludeNumbers

    .EXAMPLE
    New-Password -Length 12 -IncludeNumbers -IncludeSpecialCharacters

    .EXAMPLE
    New-Password -Length 16 -IncludeNumbers -Segments 3

    .EXAMPLE
    New-Password -APIToken -Length 64

    .NOTES
    Written By: Ben (itjedi42)
	Copyright 2024

    #>

    [CmdletBinding()]
    param (
        [Parameter(ParameterSetName = "Password")]
        [ValidateSet(4, 8, 12, 16, 24, 32, 48, 64, 72, 96, 128)]
        [Parameter(ParameterSetName = "Token")]
        [int32]
        $Length = 12,

        [Parameter(ParameterSetName = "Password")]
        [switch]
        $IncludeNumbers,

        [Parameter(ParameterSetName = "Password")]
        [switch]
        $IncludeSpecialCharacters,

        [Parameter(ParameterSetName = "Password")]
        [ValidateRange(1, 4)]
        [int32]
        $Segments = 1,

        [Parameter(ParameterSetName = "Token")]
        [switch]
        $APIToken
    )

    $LowerCase = 'a,b,c,d,e,f,g,h,i,j,k,m,n,p,q,r,t,u,v,w,x,y,z'
    $UpperCase = 'A,B,C,D,E,F,G,H,I,J,K,M,N,P,Q,R,T,U,V,W,X,Y,Z'

    If ($APIToken.IsPresent)
    {
        $PasswordLength = [Math]::Round($Length / 2.66667, 0)
    }
    Else
    {
        $PasswordLength = $Length
    }

    $CharacterArray = @()
    $CharacterArray += $LowerCase.Split(',') | Get-Random -Count ([Math]::Floor($PasswordLength / 2))
    $CharacterArray += $UpperCase.Split(',') | Get-Random -Count ([Math]::Floor($PasswordLength / 2))

    If ($IncludeNumbers.IsPresent)
    {
        $Numbers = 2..9
        $CharacterArray += $Numbers | Get-Random -Count ([Math]::Floor($PasswordLength / 2))
    }

    If ($IncludeSpecialCharacters.IsPresent)
    {
        $SpecialCharacters = '!,@,#,$,&,?'
        $CharacterArray += $SpecialCharacters.Split(',') | Get-Random -Count ([Math]::Floor($PasswordLength / 4))
    }

    $PasswordSegments = @()
    $i = 1
    While ($i -le $Segments)
    {
        If (($Segments / $PasswordLength) % 2 -eq 0)
        {
            If ($i % 2 -eq 0)
            {
                $PasswordSegments += ($CharacterArray | Get-Random -Count (($PasswordLength / $Segments))) -join ""
            }
            Else
            {
                $PasswordSegments += ($CharacterArray | Get-Random -Count (($PasswordLength / $Segments) - 1)) -join ""
            }
        }
        Else
        {
            If ($i % 2 -eq 0)
            {
                $PasswordSegments += ($CharacterArray | Get-Random -Count (($PasswordLength / $Segments) - 1)) -join ""
            }
            Else
            {
                $PasswordSegments += ($CharacterArray | Get-Random -Count (($PasswordLength / $Segments))) -join ""
            }
        }

        $i++
    }

    If ($APIToken.IsPresent)
    {
        $Password = ConvertTo-Base64 -String "$($PasswordSegments)"
    }
    Else
    {
        $Password = $PasswordSegments -join "-"
    }

    Return $Password
}

Function Get-LatestEmbyVersion
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [switch]
        $Beta
    )

    $ProgressPreference = "SilentlyContinue"
    $url = "https://github.com/MediaBrowser/Emby.Releases/releases.atom"
    [xml]$Content = Invoke-WebRequest -UseBasicParsing -Uri $url
    If ($Content)
    {
        $Feed = $Content.feed.entry
        $i = 0
        While ($i -lt 3)
        {
            If (($Feed[$i].title -like '*-beta') -or ($Feed[$i].title -like '*-Beta') -and $Beta.IsPresent)
            {
                $Version = $Feed[$i].title
                $Version = $Version.Replace("-beta", "").Replace("-Beta", "")
                $i = 3
            }
            ElseIf (($Feed[$i].title -notlike '*-beta') -and ($Feed[$i].title -notlike '*-Beta') -and !($Beta.IsPresent))
            {
                $Version = $Feed[$i].title
                $Version = $Version
                $i = 3
            }
            Else
            {
                $i++
            }
        }
        Return $Version
    }
    Else
    {
        Return  "Error"
    }
}

Function Update-EmbyServer
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Version,

        [Parameter(Mandatory)]
        [string]
        $EmbyRoot,

        [Parameter(Mandatory)]
        [string]
        $ServiceName
    )

    $ProgressPreference = "SilentlyContinue"
    Invoke-WebRequest -UseBasicParsing -Uri ("https://github.com/MediaBrowser/Emby.Releases/releases/download/$($Version)/embyserver-win-x64-$($Version).7z") -OutFile "$($EmbyRoot)\Emby-$($Version).7z"
    $7z = "$($env:ProgramFiles)\7-Zip\7z.exe"
    Stop-Service $ServiceName
    &$7z x "$($EmbyRoot)\Emby-$($Version).7z" -o"$($EmbyRoot)\" -r -aoa
    If (Test-Path -Path "$($EmbyRoot)\modifications\")
    {
        Copy-Item -Path "$($EmbyRoot)\modifications\*" -Destination "$($EmbyRoot)\" -Recurse -Force
    }

    Start-Service $ServiceName
}

Function Install-EmbyServer
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Version,

        [Parameter(Mandatory)]
        [string]
        $EmbyRoot
    )

    $ProgressPreference = "SilentlyContinue"
    Invoke-WebRequest -UseBasicParsing -Uri ("https://github.com/MediaBrowser/Emby.Releases/releases/download/$($Version)/embyserver-win-x64-$($Version).7z") -OutFile "$($EmbyRoot)\Emby-$($Version).7z"
    $7z = "$($env:ProgramFiles)\7-Zip\7z.exe"
    &$7z x "$($EmbyRoot)\Emby-$($Version).7z" -o"$($EmbyRoot)\" -r -aoa
}

Function Install-7Zip
{
    <#
    .SYNOPSIS
    Installs 7-Zip
    #>

    $7zip = "https://www.7-zip.org/"
    Try
    {
        Write-Host "$($InformationSlug) Downloading aand installing $(Format-String -Color DarkOrange '7-Zip')"
        $Response = Invoke-WebRequest -Uri $7zip
        $Link = "$($7zip)$(($Response.links | Where-Object {($_.innerHTML -eq 'Download') -and ($_.href -match '-x64')}).href)"
        $File = "$($Env:TEMP)\$(($Link.Split('/'))[-1])"
        Invoke-WebRequest -Uri $Link -OutFile $File
        $Status = Start-Process -FilePath $File -PassThru -Wait -ArgumentList "/S", "/D=`"C:\Program Files\7-Zip`""
        If ($Status.ExitCode -eq 0)
        {
            Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange '7-Zip') has been successfully installed."
        }
    }
    Catch
    {
        Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange '7-Zip') installation failed."
        Write-Error "An error occurred: $_"

    }
}

Function Install-NonSuckyServiceManager
{
    <#
    .SYNOPSIS
    Installs NSSM
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $InstallationPath
    )

    $url = "https://nssm.cc/download"

    Try
    {
        # Download the HTML content of the page
        $htmlContent = Invoke-WebRequest -Uri $url -UseBasicParsing

        # Parse the links and find the one that matches the latest stable version
        $stableLink = $htmlContent.Links | Where-Object { $_.href -match 'nssm-\d\.\d{2}.zip' -and $_.innerText -notmatch "unstable" } | Sort-Object href -Descending | Select-Object -First 1

        If ($null -eq $stableLink)
        {
            Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'Non Sucky Service Manager') Could not find the latest stable version link."
        }

        $downloadUrl = "https://nssm.cc" + $stableLink.href

        $fileName = [System.IO.Path]::GetFileName($stableLink.href)
        $destinationPath = Join-Path -Path $InstallationPath -ChildPath $fileName

        Write-Host "$($InformationSlug) Downloading NSSM from $downloadUrl to $InstallationPath"
        Invoke-WebRequest -Uri $downloadUrl -OutFile $destinationPath

        Expand-Archive -Path $destinationPath -DestinationPath $InstallationPath -Force

        $tempPath = Join-Path -Path $InstallationPath -ChildPath "$(($fileName.Substring(0,$($fileName.Length -4))))"
        $nssmExePath = Join-Path -Path $tempPath -ChildPath "win64\nssm.exe"

        Copy-Item -Path $nssmExePath -Destination $InstallationPath -Force

        Remove-Item -Path $destinationPath -Force
        Remove-Item -Path $tempPath -Recurse -Force

        Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'Non Sucky Service Manager') downloaded."
    }
    Catch
    {
        Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'Non Sucky Service Manager') installation failed."
        Write-Error "An error occurred: $_"
    }
}

Function Install-InternetInformationServices
{
    <#
    .SYNOPSIS
    Installs IIS with WebSockets, Request Filtering, Logging Tools, and Management Tools features.
    #>

    [CmdletBinding()]
    param (
        [Parameter()]
        [switch]
        $RestartIfNeeded = $false
    )

    Try
    {
        # Install IIS and the specified features
        Install-WindowsFeature -Name Web-Server, Web-WebSockets, Web-Filtering, Web-Log-Libraries, Web-Mgmt-Tools -IncludeManagementTools -Restart:$RestartIfNeeded

        # Verify that IIS and the features were installed
        $iisFeature = Get-WindowsFeature -Name Web-Server
        $webSocketsFeature = Get-WindowsFeature -Name Web-WebSockets
        $requestFilteringFeature = Get-WindowsFeature -Name Web-Filtering
        $loggingToolsFeature = Get-WindowsFeature -Name Web-Log-Libraries
        $managementToolsFeature = Get-WindowsFeature -Name Web-Mgmt-Tools

        If ($iisFeature.Installed -and $webSocketsFeature.Installed -and $requestFilteringFeature.Installed -and $loggingToolsFeature.Installed -and $managementToolsFeature.Installed)
        {
            Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'Internet Information Services with WebSockets, Request Filtering, Logging Tools, and Management Tools') has been successfully installed."
        }
        Else
        {
            Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'Internet Information Services or one of its features') installation failed."
        }
    }
    Catch
    {
        Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'Internet Information Services') installation failed."
        Write-Error "An error occurred: $_"
    }
}

Function Install-IISRewriteModule
{
    <#
    .SYNOPSIS
    Installs IIS Rewrite Module
    #>

    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $DownloadDirectory = "$($env:TEMP)\Downloads"
    )

    $rewriteModuleUrl = "https://download.microsoft.com/download/D/A/0/DA0DB2F1-7157-4550-8C2C-567A51B1B68A/rewrite_amd64_en-US.msi"

    $fileName = [System.IO.Path]::GetFileName($rewriteModuleUrl)
    $destinationPath = Join-Path -Path $DownloadDirectory -ChildPath $fileName

    Try
    {
        Write-Host "$($InformationSlug) Downloading IIS Rewrite Module from $rewriteModuleUrl to $destinationPath"
        Invoke-WebRequest -Uri $rewriteModuleUrl -OutFile $destinationPath

        Write-Host "$($InformationSlug) Installing IIS Rewrite Module..."
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i", "`"$destinationPath`"", "/quiet", "/norestart" -Wait | Out-Null

        $isRewriteInstalled = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name = 'IIS URL Rewrite Module 2'" -ErrorAction SilentlyContinue
        If ($isRewriteInstalled)
        {
            Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'IIS URL Rewrite Module 2') has been successfully installed."
        }
        Else
        {
            Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'IIS URL Rewrite Module 2') installation failed."
        }
    }
    Catch
    {
        Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'IIS URL Rewrite Module 2') installation failed."
        Write-Error "An error occurred: $_"
    }
}

Function Install-IISAARModule
{
    <#
    .SYNOPSIS
    Installs IIS Application Request Routing (ARR) Module
    #>

    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $DownloadDirectory = "$($env:TEMP)\Downloads"
    )

    $aarModuleUrl = "https://download.microsoft.com/download/C/A/F/CAF6DE19-8EE1-4B1E-B2E2-F4A8E69A92EC/requestRouter_amd64_en-US.msi"

    $fileName = [System.IO.Path]::GetFileName($aarModuleUrl)
    $destinationPath = Join-Path -Path $DownloadDirectory -ChildPath $fileName

    Try
    {
        Write-Host "$($InformationSlug) Downloading IIS ARR Module from $aarModuleUrl to $destinationPath"
        Invoke-WebRequest -Uri $aarModuleUrl -OutFile $destinationPath

        Write-Host "$($InformationSlug) Installing IIS ARR Module..."
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i", "`"$destinationPath`"", "/quiet", "/norestart" -Wait | Out-Null

        $isAARInstalled = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name = 'IIS Application Request Routing 3'" -ErrorAction SilentlyContinue
        If ($isAARInstalled)
        {
            Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'IIS Application Request Routing 3') has been successfully installed."
        }
        Else
        {
            Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'IIS Application Request Routing 3') installation failed."
        }
    }
    Catch
    {
        Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'IIS Application Request Routing 3') installation failed."
        Write-Error "An error occurred: $_"
    }
}

Function Install-VcRuntime
{
    <#
    .SYNOPSIS
    Installs the Visual C++ 2015-2022 Redistributable Runtime
    #>

    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $DownloadDirectory = "$($env:TEMP)\Downloads"
    )

    $vcRuntimeUrl = "https://aka.ms/vs/17/release/vc_redist.x64.exe"

    $fileName = [System.IO.Path]::GetFileName($vcRuntimeUrl)
    $destinationPath = Join-Path -Path $DownloadDirectory -ChildPath $fileName

    Try
    {
        Write-Host "$($InformationSlug) Downloading Visual C++ 2015-2022 Redistributable from $vcRuntimeUrl to $destinationPath"
        Invoke-WebRequest -Uri $vcRuntimeUrl -OutFile $destinationPath

        Write-Host "$($InformationSlug) Installing Visual C++ 2015-2022 Redistributable..."
        Start-Process -FilePath $destinationPath -ArgumentList "/install", "/quiet", "/norestart" -Wait | Out-Null

        # Verify installation by checking the registry for the installed version
        $vcInstalled = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64" -ErrorAction SilentlyContinue

        If ($vcInstalled)
        {
            Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'Visual C++ 2015-2022 Redistributable') has been successfully installed."
        }
        Else
        {
            Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'Visual C++ 2015-2022 Redistributable') installation failed."
        }
    }
    Catch
    {
        Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'Visual C++ 2015-2022 Redistributable') installation failed."
        Write-Error "An error occurred: $_"
    }
}

Function Install-CertifyTheWebClient
{
    <#
    .SYNOPSIS
    Installs Certify The Web Client
    #>

    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $DownloadDirectory = "$($env:TEMP)\Downloads"
    )

    $certifyClientUrl = "https://github.com/webprofusion/certify/releases/download/v5.5.5/CertifySetup-x64.msi"

    $fileName = [System.IO.Path]::GetFileName($certifyClientUrl)
    $destinationPath = Join-Path -Path $DownloadDirectory -ChildPath $fileName

    Try
    {
        Write-Host "$($InformationSlug) Downloading Certify The Web Client from $certifyClientUrl to $destinationPath"
        Invoke-WebRequest -Uri $certifyClientUrl -OutFile $destinationPath

        Write-Host "$($InformationSlug) Installing Certify The Web Client..."
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i", "`"$destinationPath`"", "/quiet", "/norestart" -Wait | Out-Null

        $isCertifyInstalled = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name = 'Certify The Web'" -ErrorAction SilentlyContinue
        If ($isCertifyInstalled)
        {
            Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'Certify The Web') has been successfully installed."
        }
        Else
        {
            Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'Certify The Web') installation failed."
        }
    }
    Catch
    {
        Write-Host "$($ErrorSlug) $(Format-String -Color DarkOrange 'Certify The Web') installation failed."
        Write-Error "An error occurred: $_"
    }
}



<### Script ###>
$SuccessSlug = "$(Format-String -Color Green '[Success]')"
$WarningSlug = "$(Format-String -Color Gold '[Warning]')"
$ErrorSlug = "$(Format-String -Color Red '[Error]')"
$InformationSlug = "$(Format-String -Color SkyBlue '[Information]')"

Write-Host "$(Format-String -Color SkyBlue '############################')"
Write-Host "$(Format-String -Color SkyBlue '#    Emby Server Manager   #')"
Write-Host "$(Format-String -Color SkyBlue '############################')"
Write-Host " "

# Check if installation path exists
Write-Host "$($InformationSlug) Checking if $(Format-String -Color DarkOrange $InstallationPath) exists."
If (Test-Path -Path $InstallationPath -ErrorAction SilentlyContinue)
{
    Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange $InstallationPath) exists."
}
Else
{
    Write-Host "$($WarningSlug) Installation path does not exist, attempting to create"
    Try
    {
        New-Item -Path $InstallationPath -Force -ItemType Directory -ErrorAction SilentlyContinue
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Unable to create folder."
        Write-Error "An error occurred: $_"

    }
}

# Handle Install
If ($Install.IsPresent)
{
    Write-Host "$($InformationSlug) Starting Emby Server Setup."
    Write-Host "$($InformationSlug) Checking Prerequisites."

    # Check if 7-Zip is installed
    Write-Host "$($InformationSlug) Checking if $(Format-String -Color DarkOrange '7-Zip') is installed."
    If (Test-Path -Path "C:\Program Files\7-Zip\7z.exe")
    {
        Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange '7-Zip') is installed."
    }
    Else
    {
        Write-Host "$($WarningSlug) $(Format-String -Color DarkOrange '7-Zip') is missing. Downloading and installing..."
        Install-7Zip
    }

    # Check if NSSM exists in installation path
    Write-Host "$($InformationSlug) Checking if $(Format-String -Color DarkOrange 'Non Sucky Service Manager') is in $InstallationPath"
    If (Test-Path -Path "$InstallationPath\nssm.exe" -ErrorAction SilentlyContinue)
    {
        Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'Non Sucky Service Manager') is present in $InstallationPath."
    }
    Else
    {
        Write-Host "$($WarningSlug) $(Format-String -Color DarkOrange 'Non Sucky Service Manager') is missing. Downloading and installing..."
        Install-NonSuckyServiceManager -DownloadDirectory $InstallationPath
    }

    # Check if IIS installed
    Write-Host "$($InformationSlug) Checking if $(Format-String -Color DarkOrange 'Internet Informataion Services') is installed."
    $iisFeature = Get-WindowsFeature -Name Web-Server
    If ($iisFeature -and $iisFeature.Installed)
    {
        Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'Internet Informataion Services') is already installed on this system."
    }
    Else
    {
        Write-Host "$($WarningSlug) $(Format-String -Color DarkOrange 'Internet Informataion Services') is not installed. Installing IIS..."
        Install-InternetInformationServices -RestartIfNeeded:$RestartIfNeeded
    }

    # Check if IIS Rewrite Module is already installed
    Write-Host "$($InformationSlug) Checking if $(Format-String -Color DarkOrange 'IIS URL Rewrite Module 2') is installed."
    $isRewriteInstalled = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name = 'IIS URL Rewrite Module 2'" -ErrorAction SilentlyContinue
    If ($isRewriteInstalled)
    {
        Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'IIS URL Rewrite Module 2') is already installed on this system."
    }
    Else
    {
        Write-Host "$($WarningSlug) $(Format-String -Color DarkOrange 'IIS URL Rewrite Module 2') is not installed. Downloading and installing..."
        Install-IISRewriteModule
    }

    # Check if IIS ARR Module is already installed
    Write-Host "$($InformationSlug) Checking if $(Format-String -Color DarkOrange 'IIS Application Request Routing 3') is installed."
    $isAARInstalled = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name = 'IIS Application Request Routing 3'" -ErrorAction SilentlyContinue

    If ($isAARInstalled)
    {
        Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'IIS Application Request Routing 3') is already installed on this system."
    }
    Else
    {
        Write-Host "$($WarningSlug) $(Format-String -Color DarkOrange 'IIS Application Request Routing 3') is not installed. Downloading and installing..."
        Install-IISAARModule -DownloadDirectory $DownloadDirectory
    }

    # Check if Visual C++ 2015-2022 Redistributable Runtime is already installed
    Write-Host "$($InformationSlug) Checking if $(Format-String -Color DarkOrange 'Visual C++ 2015-2022 Redistributable Runtim') is installed."
    $isVcRuntimeInstalled = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64" -ErrorAction SilentlyContinue

    If ($isVcRuntimeInstalled)
    {
        Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'Visual C++ 2015-2022 Redistributable') is already installed on this system."
    }
    Else
    {
        Write-Host "$($WarningSlug) $(Format-String -Color DarkOrange 'Visual C++ 2015-2022 Redistributable') is not installed. Downloading and installing..."
        Install-VcRuntime
    }

    # Check if Certify The Web Client is already installed
    Write-Host "$($InformationSlug) Checking if $(Format-String -Color DarkOrange 'Certify The Web') is installed."
    $isCertifyInstalled = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name = 'Certify The Web'" -ErrorAction SilentlyContinue

    If ($isCertifyInstalled)
    {
        Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange 'Certify The Web') is already installed on this system."
    }
    Else
    {
        Write-Host "$($WarningSlug) $(Format-String -Color DarkOrange 'Certify The Web') is not installed. Downloading and installing..."
        Install-CertifyTheWebClient
    }

    # Install Emby Server
    If ($Beta.IsPresent)
    {
        $Available = (Get-LatestEmbyVersion -Beta)
        Write-Host "$($InformationSlug) Installing Emby Server version $(Format-String -Color DarkOrange $Available) (beta)"
    }
    Else
    {
        $Available = (Get-LatestEmbyVersion)
        Write-Host "$($InformationSlug) Installing Emby Server version $(Format-String -Color DarkOrange $Available)"
    }
    Try
    {
        Install-EmbyServer -EmbyRoot $InstallationPath -Version $Available
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Emby Server installation failed."
        Write-Error "An error occurred: $_"
    }

    # Create Service Account
    If ($CreateServiceAccount.IsPresent)
    {
        Write-Host "$($InformationSlug) Creating Service Account"
        Try
        {
            $ServiceAccountSecret = $(New-Password -Length 16 -IncludeNumbers -IncludeSpecialCharacters -Segments 3)
            New-LocalUser -Name $ServiceAccount -Password (ConvertTo-SecureString $ServiceAccountSecre -AsPlainText -Force) -PasswordNeverExpires:$true -UserMayNotChangePassword:$true
            icacls $InstallationPath /setowner "$ServiceAccount" /t /c
            icacls $InstallationPath /grant "$($ServiceAccount):(OI)(CI)(F)" /t /c
        }
        Catch
        {
            Write-Host "$($ErrorSlug) User creation failed."
            Write-Error "An error occurred: $_"
        }
    }

    # Configure NSSM service for Emby-Server
    Write-Host "$($InformationSlug) Creating Emby Server Service"
    Try
    {
        $nssmPath = Join-Path -Path $InstallationPath -ChildPath "nssm.exe"

        $executablePath = Join-Path -Path $InstallationPath -ChildPath "system\EmbyServer.exe"
        $startupDirectory = Join-Path -Path $InstallationPath -ChildPath "system"

        & $nssmPath install $ServiceName $executablePath -arg "-service" -start "auto" -startupdir $startupDirectory
        & $nssmPath set $ServiceName Description "Emby Media Server"
        & $nssmPath set $ServiceName Start "SERVICE_DELAYED_AUTO_START"
        & $nssmPath set $ServiceName AppExit Default "stop"
        & $nssmPath set $ServiceName ObjectName $ServiceAccount $ServiceAccountSecret
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Server service installation failed."
        Write-Error "An error occurred: $_"
    }

    Write-Host "$($InformationSlug) Creating firewall rule."
    Try
    {
        New-NetFirewallRule -DisplayName "Web Ports - TCP" -Direction Inbound -Protocol TCP -LocalPort 80, 443 -Action Allow -Profile Any
        New-NetFirewallRule -DisplayName "Web Ports - UDP" -Direction Inbound -Protocol UDP -LocalPort 443 -Action Allow -Profile Any
        Write-Host "$($SuccessSlug) Firewall rule created."
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Firewall rule creation failed."
        Write-Error "An error occurred: $_"
    }

    # Configure IIS
    Write-Host "$($InformationSlug) Configuring Internet Information Services"
    Import-Module WebAdministration
    Stop-Website -Name "Default Web Site"

    # Configure AAR
    Try
    {
        Write-Host "$($InformationSlug) Configuring Web Farm"
        New-WebFarmsFarm -Name "WebServer"
        New-WebFarmsServer -FarmName "WebServer" -Address 'localhost' -State Started
        Set-WebFarmsLoadBalancerSettings -FarmName "WebServer" -RequestTimeout 90
        Set-WebFarmsProxySettings -FarmName "WebServer" -UseServerNameIndication $true -AddXForwardedForHeader $true
        Write-Host "$($SuccessSlug) Configured Web Farm"
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring Web Farm failed."
        Write-Error "An error occurred: $_"
    }

    # Configure Server Variables
    Try
    {
        Write-Host "$($InformationSlug) Configuring Server Variables"
        $serverVariables = @(
            "HTTP_ACCEPT_ENCODING",
            "HTTP_X_ORIGINAL_ACCEPT_ENCODING",
            "HTTP_X_FORWARDED_FOR",
            "HTTP_X_REAL_IP",
            "HTTP_REMOTE_ADDR",
            "HTTP_SEC_WEBSOCKET_EXTENSIONS",
            "HTTP_HOST"
        )

        ForEach ($variable in $serverVariables)
        {
            $exists = Get-WebConfigurationProperty -Filter "/system.webServer/proxy" -PSPath "MACHINE/WEBROOT/APPHOST" -Name "reverseRewriteHostInResponseHeaders" | Select-String -Pattern $variable
            If (-Not $exists)
            {
                Add-WebConfigurationProperty -pspath 'MACHINE/WEBROOT/APPHOST' -filter "system.webServer/proxy" -name "reverseRewriteHostInResponseHeaders" -value @{name = $variable }
                Write-Host "$($SuccessSlug) $(Format-String -Color DarkOrange $variable) has been added."
            }
            Else
            {
                Write-Host "$($InformationSlug) $(Format-String -Color DarkOrange $variable) already exists."
            }
        }

        Write-Host "$($SuccessSlug) Server variables have been configured."
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring server variables failed."
        Write-Error "An error occurred: $_"
    }

    # Create Reverse Proxy Site
    Try
    {
        Write-Host "$($InformationSlug) Configuring Reverse Proxy Site"
        $siteName = "Emby Server Reverse Proxy"
        $reverseProxyPath = Join-Path -Path $InstallationPath -ChildPath "reverseproxy"
        $port = 80
        If (-Not (Test-Path -Path $reverseProxyPath))
        {
            New-Item -Path $reverseProxyPath -ItemType Directory
            icacls $reverseProxyPath /grant "IIS_IUSRS:(OI)(CI)(M)" /t /c
            Write-Host "$($SuccessSlug) Created directory: $(Format-String -Color DarkOrange $reverseProxyPath)"
        }
        Else
        {
            Write-Host "$($InformationSlug) Directory already exists: $(Format-String -Color DarkOrange $reverseProxyPath)"
        }

        If (-Not [string]::IsNullOrEmpty($ExternalHostName))
        {
            $bindingInformation = "*:$($port):$($ExternalHostName)"
            Write-Host "$($InformationSlug) Binding site $(Format-String -Color DarkOrange $siteName) to hostname $(Format-String -Color DarkOrange $ExternalHostName) on port $(Format-String -Color DarkOrange $port)."
        }
        Else
        {
            $bindingInformation = "*:$($port):"
            Write-Host "$($InformationSlug) Binding site $(Format-String -Color DarkOrange $siteName) to all IP addresses on port $(Format-String -Color DarkOrange $port)."
        }
        If (-Not (Get-Website | Where-Object { $_.Name -eq $siteName }))
        {
            New-Website -Name $siteName -PhysicalPath $reverseProxyPath -BindingInformation $bindingInformation -Force
            Write-Host "$($SuccessSlug) Site $(Format-String -Color DarkOrange $siteName) created with path $(Format-String -Color DarkOrange $reverseProxyPath)."
        }
        Else
        {
            Write-Host "$($InformationSlug) Site $(Format-String -Color DarkOrange $siteName) already exists."
        }
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring Reverse Proxy site failed."
        Write-Error "An error occurred: $_"
    }

    # Disable Caching
    Try
    {
        Write-Host "$($InformationSlug) Configuring Caching"
        Set-WebConfigurationProperty -Filter "system.webServer/caching" -PSPath "IIS:\Sites\$siteName" -Name "enabled" -Value "false"
        Set-WebConfigurationProperty -Filter "system.webServer/caching" -PSPath "IIS:\Sites\$siteName" -Name "enableKernelCache" -Value "false"
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring caching failed."
        Write-Error "An error occurred: $_"
    }

    # Configure Request Filtering
    Try
    {
        Write-Host "$($InformationSlug) Configuring Request Filtering"
        Set-WebConfigurationProperty -Filter "system.webServer/security/requestFiltering/requestLimits" -PSPath "IIS:\Sites\$siteName" -Name "maxUrl" -Value "65534"
        Set-WebConfigurationProperty -Filter "system.webServer/security/requestFiltering/requestLimits" -PSPath "IIS:\Sites\$siteName" -Name "maxQueryString" -Value "65534"
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring request filtering failed."
        Write-Error "An error occurred: $_"
    }

    # Create IIS Headers
    Try
    {
        Write-Host "$($InformationSlug) Configuring Reverse Proxy Headers"
        Set-WebConfigurationProperty -Filter "system.web/httpRuntime" -PSPath "IIS:\Sites\$siteName" -Name "enableVersionHeader" -Value "false"
        Set-WebConfigurationProperty -Filter "system.webServer/security/requestFiltering" -PSPath "IIS:\Sites\$siteName" -Name "removeServerHeader" -Value "true"

        $headers = @(
            @{Name = "X-Powered-By"; Remove = $true },
            @{Name = "X-Frame-Options"; Value = "SAMEORIGIN" },
            @{Name = "X-Xss-Protection"; Value = "1; mode=block" },
            @{Name = "X-Content-Type-Options"; Value = "nosniff" },
            @{Name = "Referrer-Policy"; Value = "same-origin" },
            @{Name = "Feature-Policy"; Value = "sync-xhr 'self'" },
            @{Name = "Permissions-Policy"; Value = "accelerometer=(self), ambient-light-sensor=(self), autoplay=(self), battery=(), camera=(self), fullscreen=(self), geolocation=(self), gyroscope=(self), microphone=(), midi=(self), payment=(), picture-in-picture=(self), screen-wake-lock=(self), sync-xhr=(self), usb=(), web-share=(self), clipboard-read=(self), clipboard-write=(self);" },
            @{Name = "Cache-Control"; Value = "no-cache" },
            @{Name = "Cross-Origin-Resource-Policy"; Value = "cross-origin" }
        )

        ForEach ($header in $headers)
        {
            If ($header.Remove)
            {
                Remove-WebConfigurationProperty -Filter "system.webServer/httpProtocol/customHeaders" -PSPath "IIS:\Sites\$siteName" -Name "." -AtElement @{name = $header.Name } -ErrorAction SilentlyContinue
                Write-Host "$($InformationSlug) Removed header $(Format-String -Color DarkOrange $($header.Name))."
            }
            Else
            {
                Set-WebConfigurationProperty -Filter "system.webServer/httpProtocol/customHeaders" -PSPath "IIS:\Sites\$siteName" -Name "." -Value @{name = $header.Name; value = $header.Value }
                Write-Host "$($InformationSlug) Added/Updated header $(Format-String -Color DarkOrange $($header.Name)) with value $(Format-String -Color DarkOrange $($header.Value))."
            }
        }
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring reverse proxy headers failed."
        Write-Error "An error occurred: $_"
    }

    # Create Rewrite Rules
    Write-Host "$($InformationSlug) Configuring Rewrite Rules"
    $rewriteSection = "system.webServer/rewrite/rules"
    $outboundRulesPath = "system.webServer/rewrite/outboundRules"
    Try
    {
        Clear-WebConfiguration -Filter $rewriteSection -PSPath "IIS:\Sites\$siteName"
        Add-WebConfiguration -Filter $rewriteSection -PSPath "IIS:\Sites\$siteName" -Value @{
            name           = "Redirect to HTTPS"
            enabled        = "true"
            stopProcessing = "true"
            patternSyntax  = "Wildcard"
        } | Out-Null
        Add-WebConfigurationProperty -Filter "$rewriteSection/rule[@name='Redirect to HTTPS']" -PSPath "IIS:\Sites\$siteName" -Name "match" -Value @{url = "*"; negate = "false" }
        Add-WebConfigurationProperty -Filter "$rewriteSection/rule[@name='Redirect to HTTPS']" -PSPath "IIS:\Sites\$siteName" -Name "conditions" -Value @{logicalGrouping = "MatchAny" }
        Add-WebConfigurationProperty -Filter "$rewriteSection/rule[@name='Redirect to HTTPS']/conditions" -PSPath "IIS:\Sites\$siteName" -Name "." -Value @{input = "{HTTPS}"; pattern = "off" }
        Add-WebConfigurationProperty -Filter "$rewriteSection/rule[@name='Redirect to HTTPS']" -PSPath "IIS:\Sites\$siteName" -Name "action" -Value @{
            type         = "Redirect"
            url          = "https://{HTTP_HOST}{REQUEST_URI}"
            redirectType = "Found"
        }

        Write-Host "$($SuccessSlug) Rewrite rule $(Format-String -Color DarkOrange 'Redirect to HTTPS') has been added."
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring $(Format-String -Color DarkOrange 'Redirect to HTTPS') rewrite failed."
        Write-Error "An error occurred: $_"
    }

    Try
    {
        Add-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter $outboundRulesPath -name "." -value @{name = "Add Strict-Transport-Security when HTTPS"; enabled = "true" } | Out-Null
        Set-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter "$outboundRulesPath/rule[@name='Add Strict-Transport-Security when HTTPS']" -name "match" -value @{
            serverVariable = "RESPONSE_Strict_Transport_Security"
            pattern        = ".*"
        } | Out-Null
        Add-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter "$outboundRulesPath/rule[@name='Add Strict-Transport-Security when HTTPS']" -name "conditions" -value @{logicalGrouping = "MatchAll" } | Out-Null
        Add-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter "$outboundRulesPath/rule[@name='Add Strict-Transport-Security when HTTPS']/conditions" -name "." -value @{
            input   = "{HTTPS}"
            pattern = "on"
        } | Out-Null
        Write-Host "$($SuccessSlug) Rewrite rule $(Format-String -Color DarkOrange 'Strict Transport Security') has been added."
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring $(Format-String -Color DarkOrange 'Strict Transport Security') rewrite failed."
        Write-Error "An error occurred: $_"
    }

    Try
    {
        Add-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter "$outboundRulesPath/preConditions" -name "." -value @{
            name = "NeedsRestoringAcceptEncoding"
        } | Out-Null

        Add-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter "$outboundRulesPath/preConditions/preCondition[@name='NeedsRestoringAcceptEncoding']" -name "." -value @{
            input   = "{HTTP_X_ORIGINAL_ACCEPT_ENCODING}"
            pattern = ".+"
        } | Out-Null

        Add-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter $outboundRulesPath -name "." -value @{
            name         = "Restore-AcceptEncoding"
            enabled      = "true"
            preCondition = "NeedsRestoringAcceptEncoding"
        } | Out-Null

        Set-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter "$outboundRulesPath/rule[@name='Restore-AcceptEncoding']" -name "match" -value @{
            serverVariable = "HTTP_ACCEPT_ENCODING"
            pattern        = "^(.*)"
        } | Out-Null

        Set-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter "$outboundRulesPath/rule[@name='Restore-AcceptEncoding']" -name "action" -value @{
            type  = "Rewrite"
            value = "{HTTP_X_ORIGINAL_ACCEPT_ENCODING}"
        } | Out-Null
        Write-Host "$($SuccessSlug) Rewrite rule $(Format-String -Color DarkOrange 'Restore-AcceptEncoding') has been added."
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring $(Format-String -Color DarkOrange 'Restore-AcceptEncoding') rewrite failed."
        Write-Error "An error occurred: $_"
    }

    Try
    {
        Add-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter $outboundRulesPath -name "." -value @{
            name         = "Proxy to Emby"
            enabled      = "true"
            preCondition = "ResponseIsHTML"
        } | Out-Null

        Add-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter "$outboundRulesPath/preConditions" -name "." -value @{
            name = "ResponseIsHTML"
        } | Out-Null

        Add-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter "$outboundRulesPath/preConditions/preCondition[@name='ResponseIsHTML']" -name "." -value @{
            input   = "{RESPONSE_CONTENT_TYPE}"
            pattern = "^application/json|text/(.+)"
        } | Out-Null

        Set-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter "$outboundRulesPath/rule[@name='Proxy to Emby']" -name "match" -value @{
            filterByTags = "A, Area, Base, Form, Frame, Head, IFrame, Img, Input, Link, Script"
            pattern      = "^http(s)?://localhost:8096/(.*)"
        } | Out-Null

        Set-WebConfigurationProperty -pspath "IIS:\Sites\$siteName" -filter "$outboundRulesPath/rule[@name='Proxy to Emby']" -name "action" -value @{
            type  = "Rewrite"
            value = "http{R:1}://$($ExternalHostName)/{R:2}"
        } | Out-Null
        Write-Host "$($SuccessSlug) Rewrite rule $(Format-String -Color DarkOrange 'Proxy to Emby') has been added."
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring $(Format-String -Color DarkOrange 'Proxy to Emby') rewrite failed."
        Write-Error "An error occurred: $_"
    }

    # Create SSL Configuration
    Try
    {
        $certifyCLI = "C:\Program Files\CertifyTheWeb\CertifyCLI.exe"
        $certificateFriendlyName = "$siteName Certificate"
        & $certifyCLI contacts add --name $CertificateContactName --email $CertificateContactEmail
        & $certifyCLI managedcerts new --name $certificateFriendlyName --primarydomain $ExternalHostName --siteid $siteName --contact $CertificateContactEmail --auto 1
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring $(Format-String -Color DarkOrange 'Certify The Web client') failed."
        Write-Error "An error occurred: $_"
    }

    # Configure TLS
    Try
    {
        Write-Host "$($InformationSlug) Configuring TLS"
        $bindings = Get-WebBinding -Name $siteName
        $osVersion = [System.Environment]::OSVersion.Version

        # Check if the OS is Windows Server 2022 (Build 20348) or newer
        If ($osVersion.Major -ge 10 -and $osVersion.Build -ge 20348)
        {
            Write-Host "$($InformationSlug) Server 2022 or newer detected. Enabling TLS 1.3 and QUIC..."

            # Set Alt-Svc Header
            Set-WebConfigurationProperty -Filter "system.webServer/httpProtocol/customHeaders" -PSPath "IIS:\Sites\$siteName" -Name "." -Value @{name = 'alt-svc'; value = 'h3=":443"' }

            # Enable TLS 1.3
            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3" -Force
            New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Server" -Force
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.3\Server" -Name "Enabled" -Value 1
            Write-Host "$($InformationSlug) TLS 1.3 has been enabled."

            # Enable QUIC
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name "EnableQUIC" -Value 1 -Force
            Write-Host "$($InformationSlug) QUIC has been enabled."
        }
        Else
        {
            Write-Host "$($WarningSlug) This server does not support TLS 1.3 and QUIC as it is not Windows Server 2022 or newer."
        }

        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\HTTP\Parameters" -Name "EnableOcspStapling" -Value 0 -Force

        ForEach ($binding in $bindings)
        {
            # Check if the binding has the DisableLegacyTLS flag
            $disableLegacyTLS = $binding.BindingInformation -match "disablelegacytls=True"

            If (-Not $disableLegacyTLS)
            {
                # Update the binding to disable legacy TLS (TLS 1.0 and TLS 1.1)
                $bindingInfo = $binding.BindingInformation.Replace("*:", "disablelegacytls=True:*:")
                Set-WebBinding -Name $siteName -BindingInformation $bindingInfo -PropertyName "bindingInformation"
                Write-Host "$($SuccessSlug)  Disabled legacy TLS for binding: $($binding.BindingInformation) on site '$siteName'."
            }
            Else
            {
                Write-Host "$($InformationSlug) Legacy TLS already disabled for binding: $($binding.BindingInformation) on site '$siteName'."
            }
        }
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Configuring TLS failed."
        Write-Error "An error occurred: $_"
    }

    Try
    {   Write-Host "$($InformationSlug) Starting Emby Server"
        Start-Service -Name $ServiceName
        Write-Host "$($SuccessSlug) Emby Server is running."
    }
    Catch
    {
        Write-Host "$($ErrorSlug) Failed to start Emby Server service."
        Write-Error "An error occurred: $_"
    }

    Write-Host "Launching Emby in default browser."
    Start-Process "https://$($ExternalHostName)"

    If ($CreateServiceAccount.IsPresent)
    {
        Write-Host "$($InformationSlug) Service Account Information. `r`n`r`n`t`t$(Format-String -Italic -Color GoldenRod '*** SAVE FOR LATER ***')"
        Write-Host "`t`tUsername: $($ServiceAccount)"
        Write-Host "`t`tPassword: $($ServiceAccountSecret)"
    }
}

# Handle Updating
If ($Update.IsPresent)
{
    Write-Host "$($InformationSlug) Starting Emby Server Update."

    $Installed = (Get-Item -Path "$($EmbyRoot)\system\EmbyServer.exe").VersionInfo.FileVersion
    If ($Beta.IsPresent)
    {
        $Available = (Get-LatestEmbyVersion -Beta)
    }
    Else
    {
        $Available = (Get-LatestEmbyVersion)
    }
    If ($Available -ne "Error")
    {
        If ($Available -gt $Installed)
        {
            Update-EmbyServer -Version $Available -EmbyRoot $EmbyRoot -ServiceName $ServiceName
        }
    }
}



<### Signature ###>

# SIG # Begin signature block
# MIIgBgYJKoZIhvcNAQcCoIIf9zCCH/MCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBmhs5EQ2Qr4Gd1
# 9xuGbFImRdEMnBhF9uF25VtoUSwjsaCCGiUwggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwggauMIIElqADAgECAhAHNje3JFR82Ees/ShmKl5bMA0GCSqG
# SIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0yMjAzMjMwMDAwMDBaFw0zNzAzMjIyMzU5NTlaMGMx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMy
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcg
# Q0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDGhjUGSbPBPXJJUVXH
# JQPE8pE3qZdRodbSg9GeTKJtoLDMg/la9hGhRBVCX6SI82j6ffOciQt/nR+eDzMf
# UBMLJnOWbfhXqAJ9/UO0hNoR8XOxs+4rgISKIhjf69o9xBd/qxkrPkLcZ47qUT3w
# 1lbU5ygt69OxtXXnHwZljZQp09nsad/ZkIdGAHvbREGJ3HxqV3rwN3mfXazL6IRk
# tFLydkf3YYMZ3V+0VAshaG43IbtArF+y3kp9zvU5EmfvDqVjbOSmxR3NNg1c1eYb
# qMFkdECnwHLFuk4fsbVYTXn+149zk6wsOeKlSNbwsDETqVcplicu9Yemj052FVUm
# cJgmf6AaRyBD40NjgHt1biclkJg6OBGz9vae5jtb7IHeIhTZgirHkr+g3uM+onP6
# 5x9abJTyUpURK1h0QCirc0PO30qhHGs4xSnzyqqWc0Jon7ZGs506o9UD4L/wojzK
# QtwYSH8UNM/STKvvmz3+DrhkKvp1KCRB7UK/BZxmSVJQ9FHzNklNiyDSLFc1eSuo
# 80VgvCONWPfcYd6T/jnA+bIwpUzX6ZhKWD7TA4j+s4/TXkt2ElGTyYwMO1uKIqjB
# Jgj5FBASA31fI7tk42PgpuE+9sJ0sj8eCXbsq11GdeJgo1gJASgADoRU7s7pXche
# MBK9Rp6103a50g5rmQzSM7TNsQIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB
# /wIBADAdBgNVHQ4EFgQUuhbZbU2FL3MpdpovdYxqII+eyG8wHwYDVR0jBBgwFoAU
# 7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoG
# CCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDig
# NqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9v
# dEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZI
# hvcNAQELBQADggIBAH1ZjsCTtm+YqUQiAX5m1tghQuGwGC4QTRPPMFPOvxj7x1Bd
# 4ksp+3CKDaopafxpwc8dB+k+YMjYC+VcW9dth/qEICU0MWfNthKWb8RQTGIdDAiC
# qBa9qVbPFXONASIlzpVpP0d3+3J0FNf/q0+KLHqrhc1DX+1gtqpPkWaeLJ7giqzl
# /Yy8ZCaHbJK9nXzQcAp876i8dU+6WvepELJd6f8oVInw1YpxdmXazPByoyP6wCeC
# RK6ZJxurJB4mwbfeKuv2nrF5mYGjVoarCkXJ38SNoOeY+/umnXKvxMfBwWpx2cYT
# gAnEtp/Nh4cku0+jSbl3ZpHxcpzpSwJSpzd+k1OsOx0ISQ+UzTl63f8lY5knLD0/
# a6fxZsNBzU+2QJshIUDQtxMkzdwdeDrknq3lNHGS1yZr5Dhzq6YBT70/O3itTK37
# xJV77QpfMzmHQXh6OOmc4d0j/R0o08f56PGYX/sr2H7yRp11LB4nLCbbbxV7HhmL
# NriT1ObyF5lZynDwN7+YAN8gFk8n+2BnFqFmut1VwDophrCYoCvtlUG3OtUVmDG0
# YgkPCr2B2RP+v6TR81fZvAT6gt4y3wSJ8ADNXcL50CN/AAvkdgIm2fBldkKmKYcJ
# RyvmfxqkhQ/8mJb2VVQrH4D6wPIOK+XW+6kvRBVK5xMOHds3OBqhK/bt1nz8MIIG
# wjCCBKqgAwIBAgIQBUSv85SdCDmmv9s/X+VhFjANBgkqhkiG9w0BAQsFADBjMQsw
# CQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRp
# Z2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENB
# MB4XDTIzMDcxNDAwMDAwMFoXDTM0MTAxMzIzNTk1OVowSDELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMSAwHgYDVQQDExdEaWdpQ2VydCBUaW1l
# c3RhbXAgMjAyMzCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKNTRYcd
# g45brD5UsyPgz5/X5dLnXaEOCdwvSKOXejsqnGfcYhVYwamTEafNqrJq3RApih5i
# Y2nTWJw1cb86l+uUUI8cIOrHmjsvlmbjaedp/lvD1isgHMGXlLSlUIHyz8sHpjBo
# yoNC2vx/CSSUpIIa2mq62DvKXd4ZGIX7ReoNYWyd/nFexAaaPPDFLnkPG2ZS48jW
# Pl/aQ9OE9dDH9kgtXkV1lnX+3RChG4PBuOZSlbVH13gpOWvgeFmX40QrStWVzu8I
# F+qCZE3/I+PKhu60pCFkcOvV5aDaY7Mu6QXuqvYk9R28mxyyt1/f8O52fTGZZUdV
# nUokL6wrl76f5P17cz4y7lI0+9S769SgLDSb495uZBkHNwGRDxy1Uc2qTGaDiGhi
# u7xBG3gZbeTZD+BYQfvYsSzhUa+0rRUGFOpiCBPTaR58ZE2dD9/O0V6MqqtQFcmz
# yrzXxDtoRKOlO0L9c33u3Qr/eTQQfqZcClhMAD6FaXXHg2TWdc2PEnZWpST618Rr
# IbroHzSYLzrqawGw9/sqhux7UjipmAmhcbJsca8+uG+W1eEQE/5hRwqM/vC2x9XH
# 3mwk8L9CgsqgcT2ckpMEtGlwJw1Pt7U20clfCKRwo+wK8REuZODLIivK8SgTIUlR
# fgZm0zu++uuRONhRB8qUt+JQofM604qDy0B7AgMBAAGjggGLMIIBhzAOBgNVHQ8B
# Af8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAg
# BgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwHwYDVR0jBBgwFoAUuhbZ
# bU2FL3MpdpovdYxqII+eyG8wHQYDVR0OBBYEFKW27xPn783QZKHVVqllMaPe1eNJ
# MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcmwwgZAG
# CCsGAQUFBwEBBIGDMIGAMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wWAYIKwYBBQUHMAKGTGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcnQw
# DQYJKoZIhvcNAQELBQADggIBAIEa1t6gqbWYF7xwjU+KPGic2CX/yyzkzepdIpLs
# jCICqbjPgKjZ5+PF7SaCinEvGN1Ott5s1+FgnCvt7T1IjrhrunxdvcJhN2hJd6Pr
# kKoS1yeF844ektrCQDifXcigLiV4JZ0qBXqEKZi2V3mP2yZWK7Dzp703DNiYdk9W
# uVLCtp04qYHnbUFcjGnRuSvExnvPnPp44pMadqJpddNQ5EQSviANnqlE0PjlSXcI
# WiHFtM+YlRpUurm8wWkZus8W8oM3NG6wQSbd3lqXTzON1I13fXVFoaVYJmoDRd7Z
# ULVQjK9WvUzF4UbFKNOt50MAcN7MmJ4ZiQPq1JE3701S88lgIcRWR+3aEUuMMsOI
# 5ljitts++V+wQtaP4xeR0arAVeOGv6wnLEHQmjNKqDbUuXKWfpd5OEhfysLcPTLf
# ddY2Z1qJ+Panx+VPNTwAvb6cKmx5AdzaROY63jg7B145WPR8czFVoIARyxQMfq68
# /qTreWWqaNYiyjvrmoI1VygWy2nyMpqy0tg6uLFGhmu6F/3Ed2wVbK6rr3M66ElG
# t9V/zLY4wNjsHPW2obhDLN9OTH0eaHDAdwrUAuBcYLso/zjlUlrWrBciI0707NMX
# +1Br/wd3H3GXREHJuEbTbDJ8WC9nR2XlG3O2mflrLAZG70Ee8PBf4NvZrZCARK+A
# EEGKMIIHGDCCBgCgAwIBAgITQQAAAHEfmpXWr8phQwAAAAAAcTANBgkqhkiG9w0B
# AQsFADBKMRIwEAYKCZImiZPyLGQBGRYCbWUxFzAVBgoJkiaJk/IsZAEZFgdmZWx0
# b25zMRswGQYDVQQDExJGZWx0b25zLk1lIFJvb3QgQ0EwHhcNMjMwMzEyMDM0NDQx
# WhcNMjgwMzEwMDM0NDQxWjCBijESMBAGCgmSJomT8ixkARkWAm1lMRcwFQYKCZIm
# iZPyLGQBGRYHZmVsdG9uczEWMBQGA1UECwwNRmVsdG9uc19Vc2VyczEPMA0GA1UE
# CxMGQWR1bHRzMRMwEQYDVQQDEwpCZW4gRmVsdG9uMR0wGwYJKoZIhvcNAQkBFg5i
# ZW5AZmVsdG9ucy5tZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALTP
# mkrvUGAoCS+R5bqMHfJ9lwUzGGHeNMmVfYTg7SlvpraVkFDEwnyWiCJGNbek8K0i
# FZGXrJLN+sW6PYgpIbTUqQLMRV3z0S3ALnLZJaLNLH/kpbKfcOYUPYbxzmKxlLXS
# lq62PhV8Vr7kDj0x5JwXhml/houCx+TE8aCiXPVLOAUIn1LVfYO+23W3z7+7N+p8
# LR4AeiHP4E9EKThpy7Es9qfJDTUt1jUfr395I/zU5qt2ENTlSwxHq5tD8N5EcIoP
# bsLI18jJDMa015F5j5dPHVwyJNueu2TRguyi5sQYTcrvo74sen2HLQM3YvrZ2cAN
# n5WvvMk9WAIka2BPoN0CAwEAAaOCA7QwggOwMDwGCSsGAQQBgjcVBwQvMC0GJSsG
# AQQBgjcVCKa/eYThyH+ErY8zhLy9VIaVhQETgaPFRobczFsCAWQCAQUwEwYDVR0l
# BAwwCgYIKwYBBQUHAwMwDgYDVR0PAQH/BAQDAgeAMBsGCSsGAQQBgjcVCgQOMAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFJFFCFuQm7EAfQS2APByaKHLrIV3MB8GA1Ud
# IwQYMBaAFOrvjalYX+9hs7DK1ECfGHK+bhYPMIIBFQYDVR0fBIIBDDCCAQgwggEE
# oIIBAKCB/YaBu2xkYXA6Ly8vQ049RmVsdG9ucy5NZSUyMFJvb3QlMjBDQSxDTj1T
# RVJWRVIsQ049Q0RQLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZp
# Y2VzLENOPUNvbmZpZ3VyYXRpb24sREM9ZmVsdG9ucyxEQz1tZT9jZXJ0aWZpY2F0
# ZVJldm9jYXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9u
# UG9pbnSGPWh0dHA6Ly9jZXJ0aWZpY2F0ZXMuZmVsdG9ucy5tZS9jcmwvRmVsdG9u
# cy5NZSUyMFJvb3QlMjBDQS5jcmwwggFXBggrBgEFBQcBAQSCAUkwggFFMIG0Bggr
# BgEFBQcwAoaBp2xkYXA6Ly8vQ049RmVsdG9ucy5NZSUyMFJvb3QlMjBDQSxDTj1B
# SUEsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29u
# ZmlndXJhdGlvbixEQz1mZWx0b25zLERDPW1lP2NBQ2VydGlmaWNhdGU/YmFzZT9v
# YmplY3RDbGFzcz1jZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MFsGCCsGAQUFBzAChk9o
# dHRwOi8vY2VydGlmaWNhdGVzLmZlbHRvbnMubWUvY3JsL1NFUlZFUi5mZWx0b25z
# Lm1lX0ZlbHRvbnMuTWUlMjBSb290JTIwQ0EuY3J0MC8GCCsGAQUFBzABhiNodHRw
# Oi8vY2VydGlmaWNhdGVzLmZlbHRvbnMubWUvb2NzcDApBgNVHREEIjAgoB4GCisG
# AQQBgjcUAgOgEAwOYmVuQGZlbHRvbnMubWUwTwYJKwYBBAGCNxkCBEIwQKA+Bgor
# BgEEAYI3GQIBoDAELlMtMS01LTIxLTMwODQ3NjU0NzYtMTgxNzAwMTYzOC0yNTkz
# Mzg3MDUzLTExMDQwDQYJKoZIhvcNAQELBQADggEBAClbGEHdaNY3w44glmPdnk+4
# 2tdDaY/ldylXqxjAXk3lMauoKPBgaeAIhAC+qbGpmirSApV1AvDIN2D2sUCqGBgP
# 32xi/Kux15kMmIKAVcbKCCHK5JWrpTiq9q+hjLoAwhapqFOb8Pfy0I+IkOVES6yH
# UF/kp88gDZTkAfcKBXFfb18NZIWastLawhVOnyJ9R8wCKlgAmxGeNZj+DBUsAmfd
# Eg/RynU5t8auCrfXX3ZBy8gc5z9Tn9WBQuF8Rto792y+PYPBk1DkLwANjWACD/fM
# FqJZO0X6gXgh+RwTcMt488wntfZdHRXd3/R9eJMTCiZEHLABMj9anMk+JS5dJrox
# ggU3MIIFMwIBATBhMEoxEjAQBgoJkiaJk/IsZAEZFgJtZTEXMBUGCgmSJomT8ixk
# ARkWB2ZlbHRvbnMxGzAZBgNVBAMTEkZlbHRvbnMuTWUgUm9vdCBDQQITQQAAAHEf
# mpXWr8phQwAAAAAAcTANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBdyDy/Ki8oguW/OGrT
# VPIYs0RMzsfDzcoeHgSkERrwTDANBgkqhkiG9w0BAQEFAASCAQBUI/znatepYHkL
# lZzKdCxQ+VxNVcVtb3vUjs7WP0Ix4YRWUlrd8XUhrBR/Kk+vRVpMdzY390mJ3+FT
# d8owub8G1rAvYKwmU1xR1s8+uGYn+LCUL60ZLvT3usg6cgDBdFuo0gDeXr4S6J7q
# /u8CX2u9Vw1IL1vUWdak+qBI2nj1atRJ0BeAWbEtPcTYgutA414zAmfvrXEyrD6f
# ppqlLkKq8nWwjDDtMu2TOBDC/+L4YSTdrX8yY4jHete+9mYcN93hsGRLLGhQuWw+
# 6VldmUO/xvr7kJXXeYjqfSnFlWmCsbCogUMBjIfGLKopL4VoT29Qjifq01G3cxDO
# y40ZuS5BoYIDIDCCAxwGCSqGSIb3DQEJBjGCAw0wggMJAgEBMHcwYzELMAkGA1UE
# BhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2Vy
# dCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQQIQBUSv
# 85SdCDmmv9s/X+VhFjANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJAzELBgkq
# hkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI0MDgxODA4MDIzMlowLwYJKoZIhvcN
# AQkEMSIEIPQwA4NlIbqEEztX+sw6et/qCw4GrD0MBNEViXxHYyjjMA0GCSqGSIb3
# DQEBAQUABIICAFls5UWCkZZnxEq6dAeFFVCSD/ZfJDOLoh9t10j91f+RpI9rZRQ1
# D162/jeQVAXwNRmfOhcSFxK4dQzwyfSbYDs7Ygh21FURsJ08V2sFsEjH3+mKgG6q
# q5xNyUDlnHbvafWRosCAg978zadpCHCOcCmpo8zRP/Ri6PYGBBjzW5jHaC/umcDl
# pM0fm48wGAQF7Ccetz+c8/2R9mHUc/gb/QR5NJ3t7sCFRYzkcjj1g8E7FZpIdH93
# 0N/86LCED2PXT4YxDMUjMchqEL2JdhicPOXCPSpeN0U3bRm4RAo81VkcoW4sRjlH
# 35XTw7S+mCiVH9P4RvyRIBDBoGSXB85Ul91MoReKa2GUcvhojd1kPXN21egV3N/n
# UFZzohC7mdwWY6MT2Z9v01odpUQZBrbD4pDecxgt+aJKQ96Cviw7naw8HMhN/NZR
# mfH07amg3ZIlDZxRt3NVj7FvRuw8mNY1y+BtzCzrLNtTQnLL+LdVOiIMUKhCLFnn
# 8/3a6FA/X86Kd54Lm9YgudwFI9mnSAarfxtIOS6yAkL6EXdnznepei5tGrBgIPb2
# qytgRDu0k6oolY+utcA/VYxvfUPFFNFN8fsFZRchS4I4P2i6iJKZCk+7narOuTfd
# 6qMjoZ44UNgc+XQphP2JGP4Vw6br+LUl6/RLz8lWRe9VNvoTkPGajhbH
# SIG # End signature block
