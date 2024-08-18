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
> Hostname to configure reverse proxy to listen on

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
AUTHOR Ben (TheITJedi)
VERSION 0.1
GUID 849a4b01-d245-4a14-b682-b5c0fdaf0b09
This script comes with no warranties. It should work properly, but use at your own risk.
