#requires -Version 3
$script:AuthenticationSettingsPath = "$PSScriptRoot\Authentication.config.xml"
$script:AuthenticationTokenSettingsPath = "$PSScriptRoot\AuthenticationToken.config.xml"

#clear AccessToken,ValidThru variables when loading module
Remove-Variable -Name AccessToken, ValidThru -ErrorAction SilentlyContinue

<#	
        ===========================================================================
        Created on:   	10/10/2015 12:00 PM
        Created by:   	Stefan Stranger
        Filename:     	MicrosoftHealth.psm1
        -------------------------------------------------------------------------
        Module Name: MicrosoftHealth
        Description: This Microsoft Health PowerShell module was built to give a
        Microsoft Band user the ability to interact with Microsoft Health data via Powershell.

        Before importing this module, you must create your own Healt application.
        To register your application in the Microsoft Account Developer Center, 
        visit https://account.live.com/developers/applications.
        Once you do so, I recommend copying/pasting your
        Client ID and App URL to the
        parameters under the Get-OAuthAuthorization function.

        More info: http://developer.microsoftband.com/Content/docs/MS%20Health%20API%20Getting%20Started.pdf
        ===========================================================================
#>

#####################################################################################
# Helper function for MicrosoftHealth Module
# Description: 
# To start the sign-in process within your application or web service, you need to
# use a web browser or web browser control to load a URL request for the Access Token
#####################################################################################
function Get-oAuth2AccessToken 
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)] $AuthorizeUri = 'https://login.live.com/oauth20_authorize.srf',
        [Parameter(Mandatory = $true)] [string] $ClientId,
        [Parameter(Mandatory = $true)] [string] $Secret,
        [Parameter(Mandatory = $false)] [string] $RedirectUri = 'https://login.live.com/oauth20_desktop.srf'
    )

    #region - Authorization code grant flow...
    Add-Type -AssemblyName System.Windows.Forms
    $OnDocumentCompleted = {
        if($web.Url.AbsoluteUri -match 'code=([^&]*)') 
        {
            $script:AuthCode = $Matches[1]
            $form.Close()
        }
        elseif($web.Url.AbsoluteUri -match 'error=') 
        {
            $form.Close()
        }
    }
 
    
    $web = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{
        Width  = 400
        Height = 750
    }

    $web.Add_DocumentCompleted($OnDocumentCompleted)

    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{
        Width      = 400
        Height     = 750
        Autoscroll = $true
    }

    $form.Add_Shown({
            $form.Activate()
    })

    $form.Controls.Add($web)


    # Request Authorization Code
    #$Scope = @('mshealth.ReadProfile', 'mshealth.ReadActivityHistory', 'mshealth.ReadDevices', 'mshealth.ReadActivityLocation')  
    $Scope = @('mshealth.ReadProfile', 'mshealth.ReadActivityHistory', 'mshealth.ReadDevices', 'mshealth.ReadActivityLocation', 'offline_access')  
    $web.Navigate("$AuthorizeUri`?client_id=$ClientId&scope=$Scope&response_type=code&redirect_uri=$RedirectUri")
    $null = $form.ShowDialog()

    # Request AccessToken
    $Response = Invoke-RestMethod -Uri 'https://login.live.com/oauth20_token.srf' -Method Post `
    -ContentType 'application/x-www-form-urlencoded' `
    -Body "client_id=$ClientId&redirect_uri=$RedirectUri&client_secret=$Secret&code=$AuthCode&grant_type=authorization_code"
    $global:AccessToken = $Response.access_token
    $global:ValidThru = (Get-Date).AddSeconds([int]$Response.expires_in)
    $global:RefreshToken = $Response.refresh_token
    Write-Debug -Message ('Access token is: {0}' -f $AccessToken)
    #endregion

    # Write AccessCode and RefreshToken to file for future usage
    #@($($accesstoken | select @{L='AccessToken';E={$_}}),$($refreshtoken | select @{L='Refreshtoken';E={$_}})) | export-clixml -Path $AuthenticationTokenSettingsPath 
    Set-AuthenticationToken -AccessToken $AccessToken -RefreshToken $RefreshToken
    Write-Debug -Message ('Refresh token is: {0}' -f $RefreshToken)
}

#####################################################################################
# Helper function for MicrosoftHealth Module
# Description: 
# When acurrent access_token has expired, if it expires, 
# run the following request to redeem the refresh token for a new access token
#####################################################################################
function Get-oAuth2RefreshToken
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [string] $ClientId,
        [Parameter(Mandatory = $true)] [string] $Secret,
        [Parameter(Mandatory = $false)] [string] $RedirectUri = 'https://login.live.com/oauth20_desktop.srf'
    )

    $settings = Load-AuthenticationSettings
    $RefreshToken = $settings.RefreshToken
    $Response = Invoke-RestMethod -Uri 'https://login.live.com/oauth20_token.srf' -Method Post `
    -ContentType 'application/x-www-form-urlencoded' `
    -Body "client_id=$ClientId&redirect_uri=$RedirectUri&client_secret=$Secret&refresh_token=$RefreshToken&grant_type=refresh_token"
    $global:AccessToken = $Response.access_token
    $global:ValidThru = (Get-Date).AddSeconds([int]$Response.expires_in)
    $global:RefreshToken = $Response.refresh_token
    Write-Debug -Message ('Access token is: {0}' -f $AccessToken)
}


###################################################################################
# Helper function for MicrosoftHealth Module
# Description: Builds the Access Header for the Microsoft Health Cloud REST API url
###################################################################################
function Build-AccessHeader
{
    param ($AccessToken)
 
    @{
        'Authorization' = 'Bearer ' + $AccessToken
    }
}


###################################################################################
# Helper function for MicrosoftHealth Module
# Description: Helper function is called from other MicrosoftHealth Function to
#              retrieve Microsoft Health Data for different End point of REST API
###################################################################################
function Get-MicrosoftHealthData
{
    param (
        $RequestUrl
    )

    $settings = Load-AuthenticationSettings
    #Check for AccessToken variable and if there is no refreshtoken stored in settings file
    if (!($AccessToken) -and (!($settings.RefreshToken)))
    {
        Get-oAuth2AccessToken -ClientId $settings.ClientId -Secret $settings.Secret
    }
    elseif ($ValidThru -lt (Get-Date)) 
    {
        Write-Verbose 'AccessToken has expired'
        #Get-oAuth2AccessToken -ClientId $settings.ClientId -Secret $settings.Secret
        Get-oAuth2RefreshToken -ClientId $settings.ClientId -Secret $settings.Secret
    }

    $headers = Build-AccessHeader -AccessToken $AccessToken
    Write-Verbose $RequestUrl
    $result = Invoke-RestMethod -Uri $RequestUrl -Method GET -Headers $headers -ContentType 'application/json'
    return $result
}

# .EXTERNALHELP MicrosoftHealth.psm1-help.xml
Function Get-MicrosoftHealthProfile 
{
    [CmdletBinding()]
    [OutputType('System.Management.Automation.PSCustomObject')]
    [Alias('ghp')]
    param()

    process {
        try
        {
            Get-MicrosoftHealthData -RequestUrl 'https://api.microsofthealth.net/v1/me/Profile'
        }
        catch [System.Net.WebException]
        {
            'The remote server returned an error: (400) Bad Request.'
            'Authenticate first. Run Get-oAuth2AccessToken function'            
        }
    }
}

# .EXTERNALHELP MicrosoftHealth.psm1-help.xml
Function Get-MicrosoftHealthDevice 
{
    [CmdletBinding()]
    [OutputType('System.Management.Automation.PSCustomObject')]
    [Alias('ghd')]
    param()

    process {   
        try
        {
            $result = Get-MicrosoftHealthData -RequestUrl 'https://api.microsofthealth.net/v1/me/Devices'
            $result.deviceProfiles
        }
        catch [System.Net.WebException]
        {
            'The remote server returned an error: (400) Bad Request.'
            'Authenticate first. Run Get-oAuth2AccessToken function'            
        }     

    }
}

# .EXTERNALHELP MicrosoftHealth.psm1-help.xml
Function Get-MicrosoftHealthActivity 
{
    [CmdletBinding()]
    [OutputType('System.Management.Automation.PSCustomObject')]
    [Alias('gha')]
    param(
        [ValidateSet('Run', 'Bike', 'Freeplay','GuidedWorkout','Golf','Sleep')]
        [Parameter(Mandatory = $true)]
        [string]$activity,
        [Parameter(Mandatory = $false)]
        [Validatepattern('^(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)\d\d$')]
        [string]$StartTime,
        [Parameter(Mandatory = $false)]
        [Validatepattern('^(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)\d\d$')]
        [string]$EndTime,
        [Parameter(Mandatory = $false)]
        [Switch]$Details,
        [Parameter(Mandatory = $false)]
        [Switch]$MapPoints,
        [Parameter(Mandatory = $false)]
    [Switch]$MinuteSummaries)


    process {
        #Get-MicrosoftHealthData -RequestUrl 'https://api.microsofthealth.net/v1/me/Activities?activityTypes=run&activityIncludes=Details,MapPoints,MinuteSummaries'
        $params = [pscustomobject]([ordered]@{}+$PSBoundParameters)
        #Check if Details, MapPoint or MinuteSummeries params are being used. 
        if ($params.Details) 
        {
            $HttpRequestUrl = "https://api.microsofthealth.net/v1/me/Activities?activityTypes=$activity&activityIncludes"
        }
        elseif ($params.MapPoints) 
        {
            $HttpRequestUrl = "https://api.microsofthealth.net/v1/me/Activities?activityTypes=$activity&activityIncludes"
        }
        elseif ($params.MinuteSummaries) 
        {
            $HttpRequestUrl = "https://api.microsofthealth.net/v1/me/Activities?activityTypes=$activity&activityIncludes"
        }
        else
        {
            $HttpRequestUrl = "https://api.microsofthealth.net/v1/me/Activities?activityTypes=$activity"
        }

        switch ($params)
        {
            {
                ($_.Details)
            } 
            {
                $HttpRequestUrl = $HttpRequestUrl + '=Details'
            } #fix issue when this is not selected first.
            {
                ($_.MapPoints)
            } 
            {
                $HttpRequestUrl = $HttpRequestUrl + ',MapPoints'
            }
            {
                ($_.MinuteSummaries)
            } 
            {
                $HttpRequestUrl = $HttpRequestUrl + ',MinuteSummaries'
            }
            {
                -not([String]::IsNullOrWhiteSpace($_.StartTime))
            } 
            {
                $HttpRequestUrl = $HttpRequestUrl + "&startTime=$(([datetime]($_.StartTime)).toString('o'))"
            }
            {
                -not([String]::IsNullOrWhiteSpace($_.EndTime))
            } 
            {
                $HttpRequestUrl = $HttpRequestUrl + "&endTime=$(([datetime]($_.Endtime)).toString('o'))"
            }
            {
                -not([String]::IsNullOrWhiteSpace($_.MaxPageSize))
            } 
            {
                $HttpRequestUrl = $HttpRequestUrl + "&maxPageSize=$($_.maxPageSize)"
            }
            
            Default 
            {
                $HttpRequestUrl = "https://api.microsofthealth.net/v1/me/Activities?activityTypes=$activity"
            }
        }
        $result = Get-MicrosoftHealthData -RequestUrl $HttpRequestUrl
        $result.($($activity+'Activities'))
    }
}

# .EXTERNALHELP MicrosoftHealth.psm1-help.xml
Function Get-MicrosoftHealthSummary 
{
    [CmdletBinding()]
    [OutputType('System.Management.Automation.PSCustomObject')]
    [Alias('ghs')]
    param(
        
        [ValidateSet('Daily', 'Hourly')]
        [Parameter(Mandatory = $true,
        Position = 0)]
        [string]$Period,
        [Parameter(Mandatory = $false)]
        [Validatepattern('^(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)\d\d$')]
        [string]$StartTime,
        [Parameter(Mandatory = $false)]
        [Validatepattern('^(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)\d\d$')]
        [string]$EndTime,
        [Parameter(Mandatory = $false)]
        [int]$maxPageSize
    )

    process {
        $params = [pscustomobject]([ordered]@{}+$PSBoundParameters)
        $HttpRequestUrl = "https://api.microsofthealth.net/v1/me/Summaries/$Period"+'?'
        switch ($params)
        {
            {
                -not([String]::IsNullOrWhiteSpace($_.StartTime))
            } 
            {
                $HttpRequestUrl = $HttpRequestUrl + "startTime=$(([datetime]($_.StartTime)).toString('o'))"
            }
            {
                -not([String]::IsNullOrWhiteSpace($_.EndTime))
            } 
            {
                $HttpRequestUrl = $HttpRequestUrl + "&endTime=$(([datetime]($_.Endtime)).toString('o'))"
            }
            {
                -not([String]::IsNullOrWhiteSpace($_.MaxPageSize))
            } 
            {
                #Check if MaxPageSize is first param.
                Write-Verbose $HttpRequestUrl
                if (!($HttpRequestUrl -match '\?$')) #if httprequesturl does not end on ? use & sign
                {
                    $HttpRequestUrl = $HttpRequestUrl + "&maxPageSize=$($_.maxPageSize)"
                }
                elseif ($HttpRequestUrl -match '\?' -and (!($HttpRequestUrl -match '$?'))) #check if ? is used somewhere in httprequest     
                {
                    $HttpRequestUrl = $HttpRequestUrl + "&maxPageSize=$($_.maxPageSize)"
                }
                else
                {
                    $HttpRequestUrl = $HttpRequestUrl + "maxPageSize=$($_.maxPageSize)"
                }
            }
            Default 
            {
                $HttpRequestUrl = "https://api.microsofthealth.net/v1/me/Summaries/$Period"
            }
        }
        Write-Verbose $HttpRequestUrl
        $result = Get-MicrosoftHealthData -RequestUrl $HttpRequestUrl
        $result.summaries
    }
}


#region Authentication
function Get-AuthenticationSettingsPath 
{
    $script:AuthenticationSettingsPath
}

function Load-AuthenticationSettings 
{
    $path = Get-AuthenticationSettingsPath
    Import-Clixml -Path $path
}


function Set-AuthenticationToken
{
    param (
        $AccessToken,
        $RefreshToken
    )
    
    $settings = Load-AuthenticationSettings

    $AuthObject = New-Object -TypeName psObject -Property @{
        ClientId     = $settings.ClientId
        Secret       = $settings.Secret
        AccessToken  = $AccessToken
        RefreshToken = $RefreshToken
    }

    #Store AccessToken and RefreshToken in configuration file.
    $path = Get-AuthenticationSettingsPath
    Export-Clixml -Path $path -InputObject $AuthObject
}
#endregion

#region Unused Functions
function New-AuthenticationSettings
{
    param (
        $ClientId,
        $AccessToken
    )

    New-Object -TypeName psObject -Property @{
        ClientId    = $ClientId
        AccessToken = $AccessToken
    }
}

function Save-AuthenticationSettings 
{
    param (
        $AuthenticationSettings
    )

    $path = Get-AuthenticationSettingsPath
    Export-Clixml -Path $path -InputObject $AuthenticationSettings
}
#endregion
