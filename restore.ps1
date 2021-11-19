# Name - Kole Armstrong, Student ID - #001380821

# Version 2.0
# Creation Date - 10/24/2021



Param (
 [string]$OUName = "finance",
  [string]$ADUsersCSVPath = "$PSScriptRoot\financePersonnel.csv",
   [string]$SQLDataCSVPath = "$PSScriptRoot\NewClientData.csv",
    [string]$OUPath = "DC=consultingfirm,DC=com",
     [string]$Database = "ClientDB",
      [string]$Servername = ".\SQLEXPRESS"
      )
 
 #AD creation and import user data
 try {
NEW-ADOrganizationalUnit -Name Finance -ProtectedFromAccidentalDeletion $false
  $NewAD = Import-CSV $ADUsersCSVPath
    $path = "OU=finance,DC=consultingfirm,DC=com"
        foreach ($ADUser in $NewAD)
       
        {
            $First = $ADUser.first_name
                $Last = $ADUser.last_name
                    $DisplayName = $First + " " + $Last
                        $Name = $DisplayName
                        if($Name.Length -gt 20){
                            $Name = $DisplayName.SubString(0,20)
                            }
                              $Postal = $ADUser.PostalCode
  $Office = $ADUser.OfficePhone
                                  $Mobile = $ADUser.MobilePhone
                                   
                                   New-ADUser  -GivenName $First -Surname $Last -DisplayName $DisplayName -Name $Name -PostalCode $Postal -OfficePhone $Office -MobilePhone $Mobile -Path $path
                             }

}    



catch [System.OutOfMemoryException] {
    Write-Host "System out of memory exception has occurred."
}

try{
#SQL DB creation, table creation and column creation.
Import-Module -Name SqlServer -DisableNameChecking -Force
  $servername = ".\SQLEXPRESS"
    $srv = New-Object Microsoft.sqlServer.Management.Smo.Server -argumentlist $servername
      $databasename = "ClientDB"
        $db = New-Object Microsoft.sqlServer.Management.Smo.Database -argumentlist $servername, $databasename
          $db.Create()


         $CreateTable = @"
                      Use ClientDB
                      CREATE TABLE Client_A_Contacts
                     
        (
                      first_name varchar(400) NOT NULL,
                      last_name varchar(400) NOT NULL,
                      city varchar(400) NOT NULL,
                      county varchar(400) NOT NULL,
                      zip int NOT NULL,
                      officePhone varchar(400) NOT NULL,
                      mobilePhone varchar(400) NOT NULL
                      )
                     
"@
                     


                                                           #EXECUTION - SQL Data import
                                                           Invoke-Sqlcmd -ServerInstance $Servername -Database $Database -Query $CreateTable
                                                           $Table = "Client_A_Contacts"
                                                            Import-csv $SQLDataCSVPath | ForEach-Object {
                                                             Invoke-sqlcmd -Database ClientDB -ServerInstance $Servername -Query "insert into $Table
                                                             (first_name, last_name, city, county, zip, OfficePhone, MobilePhone) Values ('$($_.first_name)','$($_.last_name)','$($_.city)','$($_.county)','$($_.zip)','$($_.OfficePhone)','$($_.MobilePhone)')" }
   
}                                                            
catch [System.OutOfMemoryException] {
    Write-Host "System out of memory exception has occurred."
}