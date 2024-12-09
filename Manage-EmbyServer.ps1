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
    VERSION 0.1.2
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

    $rewriteModuleUrl = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi"

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

    $aarModuleUrl = "https://download.microsoft.com/download/E/9/8/E9849D6A-020E-47E4-9FD0-A023E99B54EB/requestRouter_amd64.msi"

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

    $certifyClientUrl = "https://certifytheweb.s3.amazonaws.com/downloads/archive/CertifyTheWebSetup_V6.1.2.exe"

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
        Install-NonSuckyServiceManager -InstallationPath $InstallationPath
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
# MIIjcwYJKoZIhvcNAQcCoIIjZDCCI2ACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAlRSLdNUm/vlc1
# YtfFeVnkaAKAC14paUOiizt+eUB7iaCCHZIwggNvMIICV6ADAgECAhA//eMbXdmu
# n0bPkAzCxYPqMA0GCSqGSIb3DQEBCwUAMEoxEjAQBgoJkiaJk/IsZAEZFgJtZTEX
# MBUGCgmSJomT8ixkARkWB2ZlbHRvbnMxGzAZBgNVBAMTEkZlbHRvbnMuTWUgUm9v
# dCBDQTAeFw0yMDA1MTgxNjEwMjNaFw0zMDA1MTgxNjIwMjNaMEoxEjAQBgoJkiaJ
# k/IsZAEZFgJtZTEXMBUGCgmSJomT8ixkARkWB2ZlbHRvbnMxGzAZBgNVBAMTEkZl
# bHRvbnMuTWUgUm9vdCBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# ANQ/CKYUxj9lg9SFVaIyWP3db5dPFW+ZuDnJWwnE48siXe9DYOhLbzN9t6S84/fc
# n3GOZOH5aeENjHHljoEtTYfPG4epKxDMit7VhMamzvsISc0Tm9Fjsq6sfNkX3tOR
# 2Pcmkg/NxJCpVd2LKKE2hASxi16f9dRRsiufY8nWjEa8tHrE7qXRdIs5h6+3gbi1
# vfbcWqT19Qdj9WGoveU25sJWnH+SiJIuzhNRrhx2sLxr9dcEiQ0axjc12pysLni9
# zu1BUBvslIhoGqNx2P0kbgHbGlHEKfZaQGUG/I6hC4dDbwFCPdyPK96ofmjoBnrm
# 19ZUQlQwCZK9u6P2VysxGsECAwEAAaNRME8wCwYDVR0PBAQDAgGGMA8GA1UdEwEB
# /wQFMAMBAf8wHQYDVR0OBBYEFOrvjalYX+9hs7DK1ECfGHK+bhYPMBAGCSsGAQQB
# gjcVAQQDAgEAMA0GCSqGSIb3DQEBCwUAA4IBAQB+QVhBym+Bpey+yRG/Lw/Uotsz
# OU0Bn5MZ4v8ea7ArpTDKB3Xek4Qal9QY7c9QBmHbK/U6E/ZgW123UO89htIJBnk6
# G3l0Fn2cPGNYOqAyovWRXNF30vrP6cMC+4efFKo7Xw80auL0Fz8vKI048EE5MFqf
# psHvmP1XdZtYrwahKICuvfP1Ee29/+wKpAngQslNX6z4wknu7mWwV99/1Drdn2c6
# niPozDwO96evpr8oelemhci/4z8bQfcQQ/sVBJCcmyg4bVom0A3lSEFttPYb37jR
# R3gZka984nWYkhn82sOBUg5853l3ESQKPbR4tQtd2pW6lQT8F0ElAjjTSed5MIIF
# jTCCBHWgAwIBAgIQDpsYjvnQLefv21DiCEAYWjANBgkqhkiG9w0BAQwFADBlMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3Qg
# Q0EwHhcNMjIwODAxMDAwMDAwWhcNMzExMTA5MjM1OTU5WjBiMQswCQYDVQQGEwJV
# UzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQu
# Y29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwggIiMA0GCSqG
# SIb3DQEBAQUAA4ICDwAwggIKAoICAQC/5pBzaN675F1KPDAiMGkz7MKnJS7JIT3y
# ithZwuEppz1Yq3aaza57G4QNxDAf8xukOBbrVsaXbR2rsnnyyhHS5F/WBTxSD1If
# xp4VpX6+n6lXFllVcq9ok3DCsrp1mWpzMpTREEQQLt+C8weE5nQ7bXHiLQwb7iDV
# ySAdYyktzuxeTsiT+CFhmzTrBcZe7FsavOvJz82sNEBfsXpm7nfISKhmV1efVFiO
# DCu3T6cw2Vbuyntd463JT17lNecxy9qTXtyOj4DatpGYQJB5w3jHtrHEtWoYOAMQ
# jdjUN6QuBX2I9YI+EJFwq1WCQTLX2wRzKm6RAXwhTNS8rhsDdV14Ztk6MUSaM0C/
# CNdaSaTC5qmgZ92kJ7yhTzm1EVgX9yRcRo9k98FpiHaYdj1ZXUJ2h4mXaXpI8OCi
# EhtmmnTK3kse5w5jrubU75KSOp493ADkRSWJtppEGSt+wJS00mFt6zPZxd9LBADM
# fRyVw4/3IbKyEbe7f/LVjHAsQWCqsWMYRJUadmJ+9oCw++hkpjPRiQfhvbfmQ6QY
# uKZ3AeEPlAwhHbJUKSWJbOUOUlFHdL4mrLZBdd56rF+NP8m800ERElvlEFDrMcXK
# chYiCd98THU/Y+whX8QgUWtvsauGi0/C1kVfnSD8oR7FwI+isX4KJpn15GkvmB0t
# 9dmpsh3lGwIDAQABo4IBOjCCATYwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU
# 7NfjgtJxXWRM3y5nP+e6mK4cD08wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6ch
# nfNtyA8wDgYDVR0PAQH/BAQDAgGGMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0
# MEUGA1UdHwQ+MDwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwEQYDVR0gBAowCDAGBgRVHSAAMA0GCSqG
# SIb3DQEBDAUAA4IBAQBwoL9DXFXnOF+go3QbPbYW1/e/Vwe9mqyhhyzshV6pGrsi
# +IcaaVQi7aSId229GhT0E0p6Ly23OO/0/4C5+KH38nLeJLxSA8hO0Cre+i1Wz/n0
# 96wwepqLsl7Uz9FDRJtDIeuWcqFItJnLnU+nBgMTdydE1Od/6Fmo8L8vC6bp8jQ8
# 7PcDx4eo0kxAGTVGamlUsLihVo7spNU96LHc/RzY9HdaXFSMb++hUD38dglohJ9v
# ytsgjTVgHAIDyyCwrFigDkBjxZgiwbJZ9VVrzyerbHbObyMt9H5xaiNrIv8SuFQt
# J37YOtnwtoeW/VvRXKwYw02fc7cBqZ9Xql4o4rmUMIIGrjCCBJagAwIBAgIQBzY3
# tyRUfNhHrP0oZipeWzANBgkqhkiG9w0BAQsFADBiMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEw
# HwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwHhcNMjIwMzIzMDAwMDAw
# WhcNMzcwMzIyMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBT
# SEEyNTYgVGltZVN0YW1waW5nIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAxoY1BkmzwT1ySVFVxyUDxPKRN6mXUaHW0oPRnkyibaCwzIP5WvYRoUQV
# Ql+kiPNo+n3znIkLf50fng8zH1ATCyZzlm34V6gCff1DtITaEfFzsbPuK4CEiiIY
# 3+vaPcQXf6sZKz5C3GeO6lE98NZW1OcoLevTsbV15x8GZY2UKdPZ7Gnf2ZCHRgB7
# 20RBidx8ald68Dd5n12sy+iEZLRS8nZH92GDGd1ftFQLIWhuNyG7QKxfst5Kfc71
# ORJn7w6lY2zkpsUdzTYNXNXmG6jBZHRAp8ByxbpOH7G1WE15/tePc5OsLDnipUjW
# 8LAxE6lXKZYnLvWHpo9OdhVVJnCYJn+gGkcgQ+NDY4B7dW4nJZCYOjgRs/b2nuY7
# W+yB3iIU2YIqx5K/oN7jPqJz+ucfWmyU8lKVEStYdEAoq3NDzt9KoRxrOMUp88qq
# lnNCaJ+2RrOdOqPVA+C/8KI8ykLcGEh/FDTP0kyr75s9/g64ZCr6dSgkQe1CvwWc
# ZklSUPRR8zZJTYsg0ixXNXkrqPNFYLwjjVj33GHek/45wPmyMKVM1+mYSlg+0wOI
# /rOP015LdhJRk8mMDDtbiiKowSYI+RQQEgN9XyO7ZONj4KbhPvbCdLI/Hgl27Ktd
# RnXiYKNYCQEoAA6EVO7O6V3IXjASvUaetdN2udIOa5kM0jO0zbECAwEAAaOCAV0w
# ggFZMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFLoW2W1NhS9zKXaaL3WM
# aiCPnshvMB8GA1UdIwQYMBaAFOzX44LScV1kTN8uZz/nupiuHA9PMA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB3BggrBgEFBQcBAQRrMGkwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RH
# NC5jcnQwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29t
# L0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIw
# CwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQB9WY7Ak7ZvmKlEIgF+ZtbY
# IULhsBguEE0TzzBTzr8Y+8dQXeJLKftwig2qKWn8acHPHQfpPmDI2AvlXFvXbYf6
# hCAlNDFnzbYSlm/EUExiHQwIgqgWvalWzxVzjQEiJc6VaT9Hd/tydBTX/6tPiix6
# q4XNQ1/tYLaqT5Fmniye4Iqs5f2MvGQmh2ySvZ180HAKfO+ovHVPulr3qRCyXen/
# KFSJ8NWKcXZl2szwcqMj+sAngkSumScbqyQeJsG33irr9p6xeZmBo1aGqwpFyd/E
# jaDnmPv7pp1yr8THwcFqcdnGE4AJxLafzYeHJLtPo0m5d2aR8XKc6UsCUqc3fpNT
# rDsdCEkPlM05et3/JWOZJyw9P2un8WbDQc1PtkCbISFA0LcTJM3cHXg65J6t5TRx
# ktcma+Q4c6umAU+9Pzt4rUyt+8SVe+0KXzM5h0F4ejjpnOHdI/0dKNPH+ejxmF/7
# K9h+8kaddSweJywm228Vex4Ziza4k9Tm8heZWcpw8De/mADfIBZPJ/tgZxahZrrd
# VcA6KYawmKAr7ZVBtzrVFZgxtGIJDwq9gdkT/r+k0fNX2bwE+oLeMt8EifAAzV3C
# +dAjfwAL5HYCJtnwZXZCpimHCUcr5n8apIUP/JiW9lVUKx+A+sDyDivl1vupL0QV
# SucTDh3bNzgaoSv27dZ8/DCCBrwwggSkoAMCAQICEAuuZrxaun+Vh8b56QTjMwQw
# DQYJKoZIhvcNAQELBQAwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hB
# MjU2IFRpbWVTdGFtcGluZyBDQTAeFw0yNDA5MjYwMDAwMDBaFw0zNTExMjUyMzU5
# NTlaMEIxCzAJBgNVBAYTAlVTMREwDwYDVQQKEwhEaWdpQ2VydDEgMB4GA1UEAxMX
# RGlnaUNlcnQgVGltZXN0YW1wIDIwMjQwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAw
# ggIKAoICAQC+anOf9pUhq5Ywultt5lmjtej9kR8YxIg7apnjpcH9CjAgQxK+CMR0
# Rne/i+utMeV5bUlYYSuuM4vQngvQepVHVzNLO9RDnEXvPghCaft0djvKKO+hDu6O
# bS7rJcXa/UKvNminKQPTv/1+kBPgHGlP28mgmoCw/xi6FG9+Un1h4eN6zh926SxM
# e6We2r1Z6VFZj75MU/HNmtsgtFjKfITLutLWUdAoWle+jYZ49+wxGE1/UXjWfISD
# mHuI5e/6+NfQrxGFSKx+rDdNMsePW6FLrphfYtk/FLihp/feun0eV+pIF496OVh4
# R1TvjQYpAztJpVIfdNsEvxHofBf1BWkadc+Up0Th8EifkEEWdX4rA/FE1Q0rqViT
# bLVZIqi6viEk3RIySho1XyHLIAOJfXG5PEppc3XYeBH7xa6VTZ3rOHNeiYnY+V4j
# 1XbJ+Z9dI8ZhqcaDHOoj5KGg4YuiYx3eYm33aebsyF6eD9MF5IDbPgjvwmnAalNE
# eJPvIeoGJXaeBQjIK13SlnzODdLtuThALhGtyconcVuPI8AaiCaiJnfdzUcb3dWn
# qUnjXkRFwLtsVAxFvGqsxUA2Jq/WTjbnNjIUzIs3ITVC6VBKAOlb2u29Vwgfta8b
# 2ypi6n2PzP0nVepsFk8nlcuWfyZLzBaZ0MucEdeBiXL+nUOGhCjl+QIDAQABo4IB
# izCCAYcwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAww
# CgYIKwYBBQUHAwgwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMB8G
# A1UdIwQYMBaAFLoW2W1NhS9zKXaaL3WMaiCPnshvMB0GA1UdDgQWBBSfVywDdw4o
# FZBmpWNe7k+SH3agWzBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRSU0E0MDk2U0hBMjU2VGltZVN0YW1w
# aW5nQ0EuY3JsMIGQBggrBgEFBQcBAQSBgzCBgDAkBggrBgEFBQcwAYYYaHR0cDov
# L29jc3AuZGlnaWNlcnQuY29tMFgGCCsGAQUFBzAChkxodHRwOi8vY2FjZXJ0cy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRSU0E0MDk2U0hBMjU2VGltZVN0
# YW1waW5nQ0EuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQA9rR4fdplb4ziEEkfZQ5H2
# EdubTggd0ShPz9Pce4FLJl6reNKLkZd5Y/vEIqFWKt4oKcKz7wZmXa5VgW9B76k9
# NJxUl4JlKwyjUkKhk3aYx7D8vi2mpU1tKlY71AYXB8wTLrQeh83pXnWwwsxc1Mt+
# FWqz57yFq6laICtKjPICYYf/qgxACHTvypGHrC8k1TqCeHk6u4I/VBQC9VK7iSpU
# 5wlWjNlHlFFv/M93748YTeoXU/fFa9hWJQkuzG2+B7+bMDvmgF8VlJt1qQcl7YFU
# MYgZU1WM6nyw23vT6QSgwX5Pq2m0xQ2V6FJHu8z4LXe/371k5QrN9FQBhLLISZi2
# yemW0P8ZZfx4zvSWzVXpAb9k4Hpvpi6bUe8iK6WonUSV6yPlMwerwJZP/Gtbu3CK
# ldMnn+LmmRTkTXpFIEB06nXZrDwhCGED+8RsWQSIXZpuG4WLFQOhtloDRWGoCwwc
# 6ZpPddOFkM2LlTbMcqFSzm4cd0boGhBq7vkqI1uHRz6Fq1IX7TaRQuR+0BGOzISk
# cqwXu7nMpFu3mgrlgbAW+BzikRVQ3K2YHcGkiKjA4gi4OA/kz1YCsdhIBHXqBzR0
# /Zd2QwQ/l4Gxftt/8wY3grcc/nS//TVkej9nmUYu83BDtccHHXKibMs/yXHhDXNk
# oPIdynhVAku7aRZOwqw6pDCCBxgwggYAoAMCAQICE0EAAABxH5qV1q/KYUMAAAAA
# AHEwDQYJKoZIhvcNAQELBQAwSjESMBAGCgmSJomT8ixkARkWAm1lMRcwFQYKCZIm
# iZPyLGQBGRYHZmVsdG9uczEbMBkGA1UEAxMSRmVsdG9ucy5NZSBSb290IENBMB4X
# DTIzMDMxMjAzNDQ0MVoXDTI4MDMxMDAzNDQ0MVowgYoxEjAQBgoJkiaJk/IsZAEZ
# FgJtZTEXMBUGCgmSJomT8ixkARkWB2ZlbHRvbnMxFjAUBgNVBAsMDUZlbHRvbnNf
# VXNlcnMxDzANBgNVBAsTBkFkdWx0czETMBEGA1UEAxMKQmVuIEZlbHRvbjEdMBsG
# CSqGSIb3DQEJARYOYmVuQGZlbHRvbnMubWUwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQC0z5pK71BgKAkvkeW6jB3yfZcFMxhh3jTJlX2E4O0pb6a2lZBQ
# xMJ8logiRjW3pPCtIhWRl6ySzfrFuj2IKSG01KkCzEVd89EtwC5y2SWizSx/5KWy
# n3DmFD2G8c5isZS10pautj4VfFa+5A49MeScF4Zpf4aLgsfkxPGgolz1SzgFCJ9S
# 1X2Dvtt1t8+/uzfqfC0eAHohz+BPRCk4acuxLPanyQ01LdY1H69/eSP81OardhDU
# 5UsMR6ubQ/DeRHCKD27CyNfIyQzGtNeReY+XTx1cMiTbnrtk0YLsoubEGE3K76O+
# LHp9hy0DN2L62dnADZ+Vr7zJPVgCJGtgT6DdAgMBAAGjggO0MIIDsDA8BgkrBgEE
# AYI3FQcELzAtBiUrBgEEAYI3FQimv3mE4ch/hK2PM4S8vVSGlYUBE4GjxUaG3Mxb
# AgFkAgEFMBMGA1UdJQQMMAoGCCsGAQUFBwMDMA4GA1UdDwEB/wQEAwIHgDAbBgkr
# BgEEAYI3FQoEDjAMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBSRRQhbkJuxAH0EtgDw
# cmihy6yFdzAfBgNVHSMEGDAWgBTq742pWF/vYbOwytRAnxhyvm4WDzCCARUGA1Ud
# HwSCAQwwggEIMIIBBKCCAQCggf2GgbtsZGFwOi8vL0NOPUZlbHRvbnMuTWUlMjBS
# b290JTIwQ0EsQ049U0VSVkVSLENOPUNEUCxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2
# aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPWZlbHRvbnMsREM9
# bWU/Y2VydGlmaWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNlP29iamVjdENsYXNzPWNS
# TERpc3RyaWJ1dGlvblBvaW50hj1odHRwOi8vY2VydGlmaWNhdGVzLmZlbHRvbnMu
# bWUvY3JsL0ZlbHRvbnMuTWUlMjBSb290JTIwQ0EuY3JsMIIBVwYIKwYBBQUHAQEE
# ggFJMIIBRTCBtAYIKwYBBQUHMAKGgadsZGFwOi8vL0NOPUZlbHRvbnMuTWUlMjBS
# b290JTIwQ0EsQ049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNl
# cnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9ZmVsdG9ucyxEQz1tZT9jQUNlcnRp
# ZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2VydGlmaWNhdGlvbkF1dGhvcml0eTBb
# BggrBgEFBQcwAoZPaHR0cDovL2NlcnRpZmljYXRlcy5mZWx0b25zLm1lL2NybC9T
# RVJWRVIuZmVsdG9ucy5tZV9GZWx0b25zLk1lJTIwUm9vdCUyMENBLmNydDAvBggr
# BgEFBQcwAYYjaHR0cDovL2NlcnRpZmljYXRlcy5mZWx0b25zLm1lL29jc3AwKQYD
# VR0RBCIwIKAeBgorBgEEAYI3FAIDoBAMDmJlbkBmZWx0b25zLm1lME8GCSsGAQQB
# gjcZAgRCMECgPgYKKwYBBAGCNxkCAaAwBC5TLTEtNS0yMS0zMDg0NzY1NDc2LTE4
# MTcwMDE2MzgtMjU5MzM4NzA1My0xMTA0MA0GCSqGSIb3DQEBCwUAA4IBAQApWxhB
# 3WjWN8OOIJZj3Z5PuNrXQ2mP5XcpV6sYwF5N5TGrqCjwYGngCIQAvqmxqZoq0gKV
# dQLwyDdg9rFAqhgYD99sYvyrsdeZDJiCgFXGygghyuSVq6U4qvavoYy6AMIWqahT
# m/D38tCPiJDlREush1Bf5KfPIA2U5AH3CgVxX29fDWSFmrLS2sIVTp8ifUfMAipY
# AJsRnjWY/gwVLAJn3RIP0cp1ObfGrgq31192QcvIHOc/U5/VgULhfEbaO/dsvj2D
# wZNQ5C8ADY1gAg/3zBaiWTtF+oF4IfkcE3DLePPMJ7X2XR0V3d/0fXiTEwomRByw
# ATI/WpzJPiUuXSa6MYIFNzCCBTMCAQEwYTBKMRIwEAYKCZImiZPyLGQBGRYCbWUx
# FzAVBgoJkiaJk/IsZAEZFgdmZWx0b25zMRswGQYDVQQDExJGZWx0b25zLk1lIFJv
# b3QgQ0ECE0EAAABxH5qV1q/KYUMAAAAAAHEwDQYJYIZIAWUDBAIBBQCggYQwGAYK
# KwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIB
# BDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg
# A+OfO2Vg4hJiWGIiDGGkUT9kcN4Q1zlJpoJUCPi32cowDQYJKoZIhvcNAQEBBQAE
# ggEAMaC/ShqpcNyxfwvrtrcijxo6522AsZrHmYX6/tfqzTJsGc+QcI/qoTCAp79U
# xh+neBKEzuv4y+rMlkKmbPyhbFhzISTEh+qDJsJqnBHIBzJnxcOM8iFur7Pqm1Wx
# dc0Er7cwzrWnAp/jvfohLjdmYFkbYXBGCQ4J9wlVcYDbXZ40EBstENz0idMYD6kd
# Uslbf6ybsIo2PCqoj0rsNS7rCEjYT+FiWgGKxqCc7uBsYnPgoR21/avFVytrljG3
# wWthlCEecDXpCwkdNg5caH/E/F7oxDS5xGgC4fjNKiXzJW9PlWCoahLLLwg3xCwP
# /YOGgIVjKQnQ262gV1zXgOYbxKGCAyAwggMcBgkqhkiG9w0BCQYxggMNMIIDCQIB
# ATB3MGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkG
# A1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3Rh
# bXBpbmcgQ0ECEAuuZrxaun+Vh8b56QTjMwQwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNDEyMDkwMzI5
# NTBaMC8GCSqGSIb3DQEJBDEiBCCKR2xGVftzuuHnN0pd5aX094EsdGJ2nvoPUfo8
# 3OreNzANBgkqhkiG9w0BAQEFAASCAgCqptN5EOgKJ5I0r9FGGiAx3CxFkLVDibVb
# ukk6ILBXhgNJYLmeTkN62MmvYsQb9UHza/9pJDRqG8M4DE6acyYBbMGoyMaJq+AU
# zAb23nxf4XN+bT+U7M1R3qH7lxhvp5ndPf/ukc7PozHLeU4L8YjRuaamxUSzq8fs
# AB94ArlDi/2ucJMsGRO5R9Da7aKRrFL9bd/K2qBFs4wuW9zBTC64bRUfHFVtUXkB
# MdEZkYWaICnkpy9UI0UJ6Lrk/GVGWAeJyoJh7Cnfoa2eowiKXkm00LDgElTp0QfT
# ti4KAcCbtnM3lhopDKTH6FMKZ2LzVZX1SMAoRe3wo7ZLfye4vCXN933SDYleXdoT
# Ebk1LUiMAncELmEpdxOl2LdsYfYysRypsnWy7jCD9LnDLUT7voXbHFWN5fFMm+8K
# CAegk9PAtfA+DAChmZz4xuTVmJ6rxUXXYNmAXgdKsMmebSJi/RC76BT6xbfW7fZz
# 9SsWue9loeETtE5f1L6Wb3G7QDkC8jf6XfTENF5HdMs505PqUfupxKRFGHF+aeM3
# ydRYWidMqV3xmASGqFQd0k0h/0JFrPBteK/bii45FetCwSFKFnkJtkO6+xjYSyR+
# dDDG0BDKr+0+I2EGpieB+q2bk3eMKmiOE301ybWt8a2jj06EMlforVfNrUWu923U
# RsKKMtEznQ==
# SIG # End signature block
