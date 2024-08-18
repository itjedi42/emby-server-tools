# Emby Tools   

A collection of PowerShell tools for use with [Emby Server](https://emby.media/)    

<br>   

<br>   

<br>   

## Manage-EmbyServer   

![null](https://img.shields.io/badge/PowerShell-_5.1+-0062AD?logo=powershell)   

   

![null](https://img.shields.io/badge/RunAs-Administrator-red)   

### Synopsis   
Emby Server Management Script   

[View File](./Manage-EmbyServer.ps1)   
<br>   
### Description      

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



  <br>   

### Parameters (Install)   
#### **-Install**   
> ![null](https://img.shields.io/badge/Type-SwitchParameter-blue?) ![null](https://img.shields.io/badge/Mandatory-True-red?)   
> ![null](https://img.shields.io/badge/DefaultValue-False-grey?color=5547a8)   
> Installs Emby-Server.   

#### **-InstallationPath**   
> ![null](https://img.shields.io/badge/Type-String-blue?) ![null](https://img.shields.io/badge/Mandatory-False-green?)   
> ![null](https://img.shields.io/badge/DefaultValue-C:\EmbyServer-grey?color=5547a8)   
> Path to install Emby-Server to.   

#### **-ServiceName**   
> ![null](https://img.shields.io/badge/Type-String-blue?) ![null](https://img.shields.io/badge/Mandatory-False-green?)   
> ![null](https://img.shields.io/badge/DefaultValue-EmbyServer-grey?color=5547a8)   
> Service Name for Emby-Server.   

#### **-ServiceAccount**      
> ![null](https://img.shields.io/badge/Type-String-blue?) ![null](https://img.shields.io/badge/Mandatory-False-green?)   
> ![null](https://img.shields.io/badge/DefaultValue-SvcEmbyServer-grey?color=5547a8)   
> User to run Emby-Server as.   

#### **-ServiceAccountSecret**   
> ![null](https://img.shields.io/badge/Type-String-blue?) ![null](https://img.shields.io/badge/Mandatory-True_If_CreateServiceAccount_not_specified-red?)   
> Password for the account to run Emby-Server As, Do not supply if using -CreateServiceAccount   

#### **-CreateServiceAccount**   
> ![null](https://img.shields.io/badge/Type-SwitchParameter-blue?) ![null](https://img.shields.io/badge/Mandatory-True_If_ServiceAccountSecret_not_specified-red?)   
> Superceeds ServiceAccountSecret. Creates a new service account and generates password which will be logged to the console.   

#### **-ExternalHostName**   
> ![null](https://img.shields.io/badge/Type-String-blue?) ![null](https://img.shields.io/badge/Mandatory-True-red?)   
> Hostname to configure reverse proxy to listen on.  DO NOT INCLUDE `HTTPS://`!      

#### **-CertificateContactName**   
> ![null](https://img.shields.io/badge/Type-String-blue?) ![null](https://img.shields.io/badge/Mandatory-True-red?)   
>  Name to use with certiify the web client as contact on certificate   

#### **-CertificateContactEmail**   
> ![null](https://img.shields.io/badge/Type-String-blue?) ![null](https://img.shields.io/badge/Mandatory-True-red?)   
> Email to use with certify the web as contact on certificate   

#### **-RestartIfNeeded**   
> ![null](https://img.shields.io/badge/Type-SwitchParameter-blue?) ![null](https://img.shields.io/badge/Mandatory-False-green?)   
> ![null](https://img.shields.io/badge/DefaultValue-False-grey?color=5547a8)   
> Tells the script to restart the system if needed when installing IIS. Note, script will need to be restarted after reboot if specified.   

#### **-Beta**
> ![null](https://img.shields.io/badge/Type-SwitchParameter-blue?) ![null](https://img.shields.io/badge/Mandatory-False-green?)   
> ![null](https://img.shields.io/badge/DefaultValue-False-grey?color=5547a8)   
> Configures with latest beta version instead of latest stable.   

<br>   
   
### Parameters (Update)      
#### **-Update**   
> ![null](https://img.shields.io/badge/Type-SwitchParameter-blue?) ![null](https://img.shields.io/badge/Mandatory-True-red?)   
> ![null](https://img.shields.io/badge/DefaultValue-False-grey?color=5547a8)   
> Updates Emby-Server.   

#### **-InstallationPath**   
> ![null](https://img.shields.io/badge/Type-String-blue?) ![null](https://img.shields.io/badge/Mandatory-False-green?)   
> ![null](https://img.shields.io/badge/DefaultValue-C:\EmbyServer-grey?color=5547a8)   
> Path to install Emby-Server to.   

#### **-ServiceName**   
> ![null](https://img.shields.io/badge/Type-String-blue?) ![null](https://img.shields.io/badge/Mandatory-False-green?)   
> ![null](https://img.shields.io/badge/DefaultValue-EmbyServer-grey?color=5547a8)   
> Service Name for Emby-Server.   

#### **-Beta**
> ![null](https://img.shields.io/badge/Type-SwitchParameter-blue?) ![null](https://img.shields.io/badge/Mandatory-False-green?)   
> ![null](https://img.shields.io/badge/DefaultValue-False-grey?color=5547a8)   
> Configures with latest beta version instead of latest stable.   

<br>   

### Examples   
#### Example 1   
```powershell
.\Manage-EmbyServer.ps1 -Install -CreateServiceAccount -CertificateContactName "Admin" -CertificateContactEmal "admin@example.com" -ExternalHostName "media.example.com"
```   
Installs latest version and creates a service account with the default options. IIS is confgured and certify the web client is used to create and maintain SSL certificate   
#### Example 2   
```powershell
.\Manage-EmbyServer.ps1 -Install -InstallationPath "F:\EmbyServer" -CreateServiceAccount -CertificateContactName "Admin" -CertificateContactEmal "admin@example.com" -ExternalHostName "media.example.com"
```   
Installs latest version and creates a service account with custom path and default options.  IIS is confgured and certify the web client is used to create and maintain SSL certificate   
#### Example 3   
```powershell
.\Manage-EmbyServer.ps1 -Install -Beta -InstallationPath "F:\EmbyServer" -CreateServiceAccount -CertificateContactName "Admin" -CertificateContactEmal "admin@example.com" -ExternalHostName "media.example.com"
```   
Installs latest beta version and creates a service account with custom path and default options.  IIS is confgured and certify the web client is used to create and maintain SSL certificate   
#### Example 4   
```powershell
.\Manage-EmbyServer.ps1 -Update
```   
Updates Emby Server to the latest version using the default options   
#### Example 5   
```powershell
.\Manage-EmbyServer.ps1 -Update -InstallationPath "F:\EmbyServer"
```   
Updates Emby Server to the latest version using a custom path   
#### Example 6   
```powershell
.\Manage-EmbyServer.ps1 -Update -InstallationPath "F:\EmbyServer" -ServiceName "My Server"
```   
Updates Emby Server to the latest version using a custom path and service name   
#### Example 7   
```powershell
.\Manage-EmbyServer.ps1 -Update -Beta
```   
Updates Emby Server to latest beta version using the default options   
#### Example 8   
```powershell
.\Manage-EmbyServer.ps1 -Update  -Beta -InstallationPath "F:\EmbyServer"
```   
Updates Emby Server to the latest beta version using a custom path   

<br>   


### Files   
- Manage-EmbyServer.ps1   


<br>   


### Notes   
AUTHOR Ben (itjedi42)   
VERSION 0.1   
GUID 849a4b01-d245-4a14-b682-b5c0fdaf0b09   
This script comes with no warranties. It should work properly, but use at your own risk.   
