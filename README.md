# Microsoft Office Offline Installer

This repository contains scripts and instructions for creating and installing an offline package of Microsoft Office (Volume License).

## Preparation

1. Go to the official [Office Configuration Tool](https://config.office.com/deploymentsettings).
2. Create your custom configuration and export the **configuration.xml** file.
3. Download and install the [Office Deployment Tool (ODT)](https://www.microsoft.com/en-us/download/details.aspx?id=49117).
4. Replace the default `configuration.xml` inside the ODT folder with the one you exported.

## Installation

Open Command Prompt or PowerShell in the ODT folder and run:

### For PowerShell
```powershell
.\setup.exe /download configuration.xml
```
### For CMD
```cmd
setup.exe /download configuration.xml
```

## Running the Installation Script

After receiving the offline installer:

1. Copy the provided **`INSTALL.cmd`** file into the same folder where `setup.exe` and `configuration.xml` are located.  
2. Right-click **`INSTALL.cmd`** and choose **Run as Administrator**.  
3. The script will start installation and activation automatically.

## Licensing setup

In order for the script to be able to take licenses from your KMS, you need to enter the IP in line 39
