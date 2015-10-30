# Microsoft Health
PowerShell Module for Microsoft Health Cloud API

10/29/2015: - Initial version
10/30/2015: - Updated README.md file

# Install
1. Download the MicrosoftHealth-masterzip file.
2. Unblock the zip file (right click and select unblock).
3. Extract zip contents MicrosoftHealth-master.zip file.
4. Rename MicrosoftHealth-master folder to MicrosoftHealth.
5. Copy renamed MicrosoftHealth folder with contents to C:\Users\[username]\Documents\WindowsPowerShell\Modules folder
6. Run import-module MicrosoftHealth cmdlet in PowerShell after following below steps.

# Microsoft Health Register your app
In order to connect to the Microsoft Health Cloud APIs, you will need a Microsoft Account with a registered application. 
Remember that each app registered with Microsoft Account Developer Center is associated with the Microsoft account used for login to https://account.live.com/developers/applications. 
We recommend that you use a developer account instead of a personal account. 
To learn more about developer accounts, please visit https://msdn.microsoft.com/enus/library/windows/apps/hh868184.aspx 
To sign up for a Microsoft account, please visit http://account.microsoft.com.  
Please make sure your Microsoft account is associated with a valid email address so we can keep you up-to-date on our latest status and releases. 
To register your application in the Microsoft Account Developer Center, visit https://account.live.com/developers/applications. 
This will provide the client id and client secret that can be used within your application to authorize against Microsoft Health Cloud APIs.  

1. Go to https://account.live.com/developers/applications. 
2. Login with your Microsoft account.
3. Create application.
4. Enter a name for your awesome app (Microsoft Health PowerShell).
5. Select Yes for Mobile or desktop client app.
6. Save settings.
7. Retrieve and write down your client id and client secret needed for using the PowerShell MicrosoftHealth Module.
8. Open Authentication.config.xml_example file in your downloaded MicrosoftHealth Module folder.
9. Enter your personal MicrosoftHealth client id and secret in the Authentication.config.xml_example file.
10. Save Authentication.config.xml_example file as Authentication.config.xml in MicrosoftHealth Module folder.

#Usage
After importing the MicrosoftHealth Module run Get-Command -Module MicrosoftHealth for the available commands in the module.
```PowerShell
PS C:\windows\system32> get-command -module Microsofthealth

CommandType     Name                                               Version    Source                                                                   
-----------     ----                                               -------    ------                                                                   
Function        Get-MicrosoftHealthActivity                        1.0        microsofthealth                                                          
Function        Get-MicrosoftHealthDevice                          1.0        microsofthealth                                                          
Function        Get-MicrosoftHealthProfile                         1.0        microsofthealth                                                          
Function        Get-MicrosoftHealthSummary                         1.0        microsofthealth                                                          
Function        Get-oAuth2AccessToken                              1.0        microsofthealth
```

Use Get-Help command to explore the command.
```PowerShell
PS C:\windows\system32> Get-Help Get-MicrosoftHealthProfile

NAME
    Get-MicrosoftHealthProfile
    
SYNOPSIS
    Gets the UserProfile.
    
    
SYNTAX
    Get-MicrosoftHealthProfile [<CommonParameters>]
    
    
DESCRIPTION
    The UserProfile object contains the general profile of the person using Microsoft Band.
    

RELATED LINKS

REMARKS
    To see the examples, type: "get-help Get-MicrosoftHealthProfile -examples".
    For more information, type: "get-help Get-MicrosoftHealthProfile -detailed".
    For technical information, type: "get-help Get-MicrosoftHealthProfile -full".
```