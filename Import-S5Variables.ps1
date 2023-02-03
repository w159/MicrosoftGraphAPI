<##

Use this function ahead of all scripts to import the required variables

#>

function Import-S5Variables {
   
    # Directory Path Section
    $script:CurrentLocation            = ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath('.\')) + "\Auto Assess and Document"
    $script:CSV_Exports                = "$CurrentLocation\CSV_Exports"
    
    # File Path Section
    $script:PNGLogoPath                = "$CurrentLocation\S5-SmallLogo.png"
    #$script:AppCsvPath                 =  "$CSV_Exports\APP_Variables.csv"

    # CSV Imports Section
    #$script:APP_Variables              = Import-Csv -Path $AppCsvPath

    # 365 Auth Variables Section
    $script:AppId                      = $S5App.ClientID
    $script:AppSecret                  = $S5App.ClientSecret
    $script:TenantId                   = $S5App.TenantId
    $script:resourceAppIdUri           = 'https://api-us.securitycenter.microsoft.com'
    $script:oAuthUri                   = "https://login.microsoftonline.com/$TenantId/oauth2/token"

    # Settings Variables Section
    $script:ProgressPreference         = 'SilentlyContinue'
    ${script:[Net.ServicePointManager]::SecurityProtocol} = [Net.SecurityProtocolType]::Tls12
    $script:MaximumFunctionCount       = "10000"

    # Graph URL's
    $script:GRAPH_BaseURL              = "https://graph.microsoft.com"
    $script:graphApiVersion            = "v1.0"
    $script:User_resource              = "users"
    $script:AuditLogsSignIns           = "auditLogs/signIns"

    #$resourceAppIdURI = "https://graph.microsoft.com"
    #$authority = "https://login.microsoftonline.com/$Tenant"

}


$RequiredModules = @(

  'AzureAD',
  'AzureADPreview',
  'Az.Resources',
  'Microsoft.Graph',
  'PSWriteWord',
  'MSAL.PS',
  'PSWriteHTML',
  'PSWriteOffice'

)


foreach ($RequiredModule in $RequiredModules) {
  if (-not (Get-Module $RequiredModule -ListAvailable)) {
    "Installing $RequiredModule now, please wait...."
    Install-Module $RequiredModule -AllowClobber -Force
    "Importing $RequiredModule now, please wait...."
    Import-Module $RequiredModule
  }
}

