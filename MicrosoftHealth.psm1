#requires -Version 3
$script:AuthenticationSettingsPath = "$PSScriptRoot\Authentication.config.xml"

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
        Height = 500
    }

    $web.Add_DocumentCompleted($OnDocumentCompleted)

    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{
        Width  = 400
        Height = 500
    }

    $form.Add_Shown({
            $form.Activate()
    })

    $form.Controls.Add($web)


    # Request Authorization Code
    $Scope = @('mshealth.ReadProfile', 'mshealth.ReadActivityHistory', 'mshealth.ReadDevices', 'mshealth.ReadActivityLocation')  
    $web.Navigate("$AuthorizeUri`?client_id=$ClientId&scope=$Scope&response_type=code&redirect_uri=$RedirectUri")
    $null = $form.ShowDialog()

    # Request AccessToken
    $Response = Invoke-RestMethod -Uri 'https://login.live.com/oauth20_token.srf' -Method Post `
    -ContentType 'application/x-www-form-urlencoded' `
    -Body "client_id=$ClientId&redirect_uri=$RedirectUri&client_secret=$Secret&code=$AuthCode&grant_type=authorization_code"
    $global:AccessToken = $Response.access_token
    $global:ValidThru = (Get-Date).AddSeconds([int]$Response.expires_in)
    $RefreshToken = $Response.refresh_token
    Write-Debug -Message ('Access token is: {0}' -f $AccessToken)
    #endregion
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
    #Check for AccessToken variable
    if (!($AccessToken))
    {
        Get-oAuth2AccessToken -ClientId $settings.ClientId -Secret $settings.Secret
    }
    elseif ($ValidThru -lt (Get-Date)) 
    {
        'AccessToken has expired'
        Get-oAuth2AccessToken -ClientId $settings.ClientId -Secret $settings.Secret
    }

    $headers = Build-AccessHeader -AccessToken $AccessToken
    Write-Verbose $RequestUrl
    $result = Invoke-RestMethod -Uri $RequestUrl -Method GET -Headers $headers -ContentType 'application/json'
    return $result
}

Function Get-MicrosoftHealthProfile 
{
    <#
            .Synopsis
            Gets the UserProfile.
            .DESCRIPTION
            The UserProfile object contains the general profile of the person using Microsoft Band. 
            .EXAMPLE
            Get-MicrosoftHealthProfile

            firstName      : John
            middleName     : null
            lastName       : Doe
            birthdate      : 1970-01-01T00:00:00.000+00:00
            postalCode     : 
            gender         : Male
            height         : 1680
            weight         : 72000
            lastUpdateTime : 2015-10-28T19:43:27.123+00:00
    #>
    [CmdletBinding()]
    [OutputType('System.Management.Automation.PSCustomObject')]
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

Function Get-MicrosoftHealthDevice 
{
    <#
            .Synopsis
            Gets Microsoft Band Device information
            .DESCRIPTION
            The Device object represents a device that collects and sends user data to the Microsoft Health service
            .EXAMPLE
            Get-MicrosoftHealthDevice
    #>
    [CmdletBinding()]
    [OutputType('System.Management.Automation.PSCustomObject')]
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

Function Get-MicrosoftHealthActivity 
{
<#
.Synopsis
   Gets Microsoft Band Activity
.DESCRIPTION
   The Activity object represents activities a user has completed using the tiles on the Microsoft Band.
   The following activities can be retrieved:
   - Run
   - Bike
   - Free Play (Workout)
   - Guided Workout
   - Golf
   - Sleep 
.EXAMPLE
   Get-MicrosoftHealthActivity -Activity Run
.EXAMPLE
   Get-MicrosoftHealthActivity -Activity Bike
.EXAMPLE
   Get-MicrosoftHealthActivity -activity Run -Details -MapPoints -MinuteSummaries
.EXAMPLE
   Get-MicrosoftHealthActivity -activity Run -Details -StartTime 10-01-2015 -EndTime 10-28-2015
   Returns all Run Activities with Detailed information between dates 10-01-2015 and 10-28-2015
   
#>
    [CmdletBinding()]
    [OutputType('System.Management.Automation.PSCustomObject')]
    param(
        [ValidateSet("Run", "Bike", "Freeplay","GuidedWorkout","Golf","Sleep")]
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
        if ($params.Details) {$HttpRequestUrl = "https://api.microsofthealth.net/v1/me/Activities?activityTypes=$activity&activityIncludes"}
        elseif ($params.MapPoints) {$HttpRequestUrl = "https://api.microsofthealth.net/v1/me/Activities?activityTypes=$activity&activityIncludes"}
        elseif ($params.MinuteSummaries) {$HttpRequestUrl = "https://api.microsofthealth.net/v1/me/Activities?activityTypes=$activity&activityIncludes"}
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
        $result.($($Activity+"Activities"))
    }
}

Function Get-MicrosoftHealthSummary 
{
<#
.Synopsis
   Gets summarized data
.DESCRIPTION
    Provides a sum-up of user data on an hourly or daily basis. 
    This data is divided into several sub-groups: 
    - Steps The total number of steps taken in the time period 
    - Calories Burned The total calories burned in the time period 
    - Heart Rate The average, peak, and lowest heart rate in the time period 
    - Distance 
.EXAMPLE
   Get-MicrosoftHealthSummary -Period Daily
   Gets a summary of Daily data on a Daily basis
.EXAMPLE
   Get-MicrosoftHealthSummary -Period Hourly
   Gets a summary of Daily data on a Hourly basis
.EXAMPLE
   Get-MicrosoftHealthSummary -Period Hourly -StartTime 10-01-2015 -EndTime 10-10-2015 -maxPageSize 1
   Gets a summary overview of the Hourly data from 1st of October 2015 till 10th of October 2015 with page size of 1 
#>
    [CmdletBinding()]
    [OutputType('System.Management.Automation.PSCustomObject')]
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
        $HttpRequestUrl = "https://api.microsofthealth.net/v1/me/Summaries/$Period"+"?"
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
                write-verbose $HttpRequestUrl
                if (!($HttpRequestUrl -match '$?')) #if httprequesturl does not end on ? use & sign
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

function Load-AuthenticationSettings 
{
    $path = Get-AuthenticationSettingsPath
    Import-Clixml -Path $path
}
#endregion