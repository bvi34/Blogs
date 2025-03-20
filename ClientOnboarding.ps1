#################-Prerequisites-#############################################
# PowerShell Modules: ActiveDirectory, SQLServer, ReportingServicesTools, IISAdministration, GroupPolicy
# Files: 1Password CLI tool, 1Password Desktop App
###########################################################################

# System Variables
$onePasswordPath = "C:\Program Files\1Password CLI\"
Set-Location $onePasswordPath
import-module SqlServer
import-module ActiveDirectory
Import-Module ReportingServicesTools -UseWindowsPowerShell
Import-Module GroupPolicy

# Server Addresses (Genericized)
$SSRSServer = "SSRS_SERVER_IP"
$ADServer = "AD_SERVER_IP"
$SQLServer = "SQL_SERVER_IP"
$FileServer = "FILE_SERVER_IP"
$WebServer = "WEB_SERVER_IP"
$DNSServer = "DNS_SERVER_IP"
$SpectrumServer = "SPECTRUM_SERVER_IP"
$Confirmation = "N"

# User Inputs
while($Products -ne "4"){
    while($Confirmation -ne "Y"){
        $ClientCode = Read-Host -Prompt "Please Input the Client Code (Example: ABC123)"
        $ClientName = Read-Host "Please Input the Client Name (Example: Acme Corp)"
        $Products = Read-Host "Select implementation option (1 for TypeA, 2 for TypeB, 3 for TypeC, 4 to exit)"
        $Retry = Read-Host "Is this a retry of a past implementation? (y/n)"
        Write-Host "$ClientCode - $ClientName - Implementation Option: $Products"
        if($Retry -eq "y"){
            Write-Host("This is a retry of an implementation.")
        }
        $Confirmation = Read-Host "Was that correct? (y/n)"
    }
}

######################################### OnePassword Section ####################################################################################################
if($Retry -eq "y"){
    $Do = Read-Host "Is 1Password needed? (y/n)"
}
if($Do -eq "y"){
    $Vault = "$ClientCode - $ClientName"
    if(!(./op vault get $Vault)){
        ./op vault create $Vault
        ./op vault group grant --vault $Vault --group Implementation --permissions allow_viewing,allow_editing
        ./op vault group grant --vault $Vault --group Administrators --permissions allow_viewing,allow_editing
    }
    
    ./op item create --category login --vault $Vault --title "$ClientCode Service Account" --generate-password='letters,digits,symbols,15' username[text]="$ClientCode-Service"
}

######################################### Active Directory Section ####################################################################################################
if($Retry -eq "y"){
    $Do = Read-Host "Is AD needed? (y/n)"
}
if($Do -eq "y"){
    $OUName = "OU=$ClientCode,DC=example,DC=com"
    $SecurityGroup = "$ClientCode Users"
    $User = "$ClientCode User"
    $UPNUser = "$ClientCode@example.com"
    
    # Create Organizational Unit
    New-ADOrganizationalUnit -Name "$ClientCode" -Path "DC=example,DC=com" -Server $ADServer
    # Create security group
    New-ADGroup -Name $SecurityGroup -GroupScope Global -Path $OUName -Server $ADServer
    # Create user
    $SecuredPassword = ConvertTo-SecureString -String $(./op item get "$ClientCode Service Account" --fields password) -AsPlainText -Force
    New-ADUser -Name $User -SamAccountName "$ClientCode" -UserPrincipalName $UPNUser -AccountPassword $SecuredPassword -Path $OUName -Server $ADServer -Enabled $true
    Add-ADGroupMember -Identity $SecurityGroup -Members $User -Server $ADServer
}

######################################### SQL Server Section ####################################################################################################
if($Retry -eq "y"){
    $Do = Read-Host "Is SQL needed? (y/n)"
}
if($Do -eq "y"){
    $DatabaseName = "DB_$ClientCode"
    $ServerInstance = "$SQLServer"
    
    # Create SQL User
    $Username = "$ClientCode_SQLUser"
    $Password = ConvertTo-SecureString -String $(./op item get "$ClientCode Service Account" --fields password) -AsPlainText -Force
    $SQLUser = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $Password
    
    $SQLQuery = "
    CREATE LOGIN [$Username] WITH PASSWORD = '$(./op item get "$ClientCode Service Account" --fields password)', CHECK_POLICY=OFF;
    USE $DatabaseName;
    CREATE USER [$Username] FOR LOGIN [$Username] WITH DEFAULT_SCHEMA=[dbo];
    EXEC sp_addrolemember 'db_owner', '$Username';"
    
    Invoke-Sqlcmd -ServerInstance $ServerInstance -Query $SQLQuery
}

######################################### SSRS Configuration ####################################################################################################
if($Retry -eq "y"){
    $Do = Read-Host "Is SSRS needed? (y/n)"
}
if($Do -eq "y"){
    $SSRSuri = "http://SSRS_SERVER_IP/ReportServer/ReportService2010.asmx?wsdl"
    New-rsFolder -RsFolder "/" -FolderName "$ClientCode Reports" -Credential $Credential -ReportServerUri $SSRSuri
    New-rsFolder -RsFolder "/$ClientCode Reports" -FolderName "DataSource" -Credential $Credential -ReportServerUri $SSRSuri
    New-rsFolder -RsFolder "/$ClientCode Reports" -FolderName "UserCreatedRpts" -Credential $Credential -ReportServerUri $SSRSuri
    
    $SQLUser = "$ClientCode_SQLUser"
    $PWord = ConvertTo-SecureString -String $(./op item get "$ClientCode Service Account" --fields password) -AsPlainText -Force
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SQLUser, $PWord
    
    New-RsDataSource -RsFolder "/$ClientCode Reports/DataSource" -Name "MainDataSource" -Extension "SQL" -ConnectionString "Data Source = SQL_SERVER_IP; Initial Catalog=DB_$ClientCode" -DatasourceCredentials $Credential -CredentialRetrieval "Store" -Credential $Credential -ReportServerUri $SSRSuri
}

######################################### Conclusion ####################################################################################################
Write-Host "Setup complete for client: $ClientCode - $ClientName"
