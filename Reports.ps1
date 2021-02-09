#Requirements

# 1) Access Software
# 2) Access Database engine
# 3) Powershell Active Directory Module


## FUNCTIONS ###

Function Invoke-ADOCommand($Db, $Command){
    $conn = New-Object -ComObject ADODB.Connection
    $recordset = New-Object -ComObject ADODB.Recordset
    $conn.Open("Provider=Microsoft.Ace.OLEDB.12.0;Data Source=$Db")
    $conn.Execute($command)
    $conn.Close()
}


### Create a new Table ### Why? To keep track of new and old servers.
$Db = "D:\Reports\Database\Reports.accdb"
$table = "Servers" + (get-date).GetDateTimeFormats()[10].Replace(",","").Replace(" ","")
$Fields = "ID Text, ServerName Text, Created Text, SerialNumber Text, OperatingSystem Text, InstallDate Text, IP Text, LastLogon Text, LastBootUpTime Text, ObjectLocation Text"
$command = "CREATE TABLE $table ($fields)"

Invoke-ADOCommand -db $Db -command $command

#
# delete all the existing records
#

<#
Step 1 : Connection to the Database
$data = "D:\Reports\Database\Reports.accdb"
$conn = New-Object -ComObject ADODB.Connection
$recordset = New-Object -ComObject ADODB.Recordset
$conn.Open("Provider=Microsoft.Ace.OLEDB.12.0;Data Source=$data")


Step 2: Delete All records
$query = "DELETE FROM Servers"
$cursor = 3
$lock = 3
$recordset.open("$query",$conn, $cursor,$lock)
$conn.Close()
 #>

#
# Updating the table with new servers
#


<# Get all online servers #>
$servers = (Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -Properties operatingsystem).Name | sort

<# Creating Server objects #>
$newObject = $servers | %{
    
    $eachServer = $_
    $Computer = Get-ADComputer -Identity $eachServer -Properties *
    
    if(Test-Connection -ComputerName $_ -Count 1 -ErrorAction SilentlyContinue)
    {
        try
        {
            
            $os = Get-WmiObject -ComputerName $eachServer -Class win32_operatingsystem -ErrorAction Stop
            $serial = (Get-WmiObject -ComputerName $eachServer -Class win32_bios -Property * -ErrorAction stop).SerialNumber
        }

        catch
        {
        New-Object -TypeName PSObject -Property @{
                ServerName = $Computer.Name;
                Created = $Computer.Created;
                SerialNumber = "";
                OperatingSystem = $Computer.OperatingSystem;
                InstallDate = "";
                IP = $Computer.IPv4Address;
                LastLogon = [datetime]::FromFileTime($Computer.lastLogon);
                LastBootUpTime = "";
                ObjectLocation = $Computer.DistinguishedName;
            }
            $eachServer = $null;        
        }
        finally
        {
        if($eachServer -ne $null){

            New-Object -TypeName PSObject -Property @{
                ServerName = $Computer.Name;
                Created = $Computer.Created;
                SerialNumber = $serial;
                OperatingSystem = $Computer.OperatingSystem;
                InstallDate = $os.ConverttoDateTime($os.InstallDate);
                IP = $Computer.IPv4Address;
                LastLogon = [datetime]::FromFileTime($Computer.lastLogon);
                LastBootUpTime = $os.ConverttoDateTime($os.lastbootuptime);
                ObjectLocation = $Computer.DistinguishedName;
            }
            }
        }
     }
     else{
         New-Object -TypeName PSObject -Property @{
                    ServerName = $Computer.Name;
                    Created = $Computer.Created;
                    SerialNumber = "";
                    OperatingSystem = $Computer.OperatingSystem;
                    InstallDate = "";
                    IP = $Computer.IPv4Address;
                    LastLogon = [datetime]::FromFileTime($Computer.lastLogon);
                    LastBootUpTime = "";
                    ObjectLocation = $Computer.DistinguishedName;
                }
            }

}

<#--Connection to the Database--#>
$data = "D:\Reports\Database\Reports.accdb"
$conn = New-Object -ComObject ADODB.Connection
$recordset = New-Object -ComObject ADODB.Recordset
$conn.Open("Provider=Microsoft.Ace.OLEDB.12.0;Data Source=$data")

<#--Selecting the table--#>
$query = "Select * from $table"
$cursor = 3
$lock = 3
$recordset.open("$query",$conn, $cursor,$lock)

<#--Updating the records--#>
for($j=0;$j -lt $newObject.Count;$j++)
{
    $index = $j + 1
    $i = $index - 1
    $InstallDate = $newObject[$i].InstallDate
    $LastBootUpTime = $newObject[$i].LastBootUpTime
    $Name = $newObject[$i].ServerName
    $Created = $newObject[$i].Created
    $OperatingSystem = $newObject[$i].OperatingSystem
    $SerialNumber = $newObject[$i].SerialNumber
    $LastLogon = $newObject[$i].LastLogon
    $IP = $newObject[$i].IP
    $Ol = $newObject[$i].ObjectLocation

    $recordset.AddNew()
    $recordset.Fields.Item("ID").value = $index
    $recordset.Fields.Item("InstallDate").value = "$InstallDate"
    $recordset.Fields.Item("LastBootUpTime").value = "$LastBootUpTime"
    $recordset.Fields.Item("ServerName").value = "$Name"
    $recordset.Fields.Item("Created").value = "$Created"
    $recordset.Fields.Item("OperatingSystem").value = "$OperatingSystem"
    $recordset.Fields.Item("SerialNumber").value = "$SerialNumber"
    $recordset.Fields.Item("LastLogon").value = "$LastLogon"
    $recordset.Fields.Item("IP").value = "$IP"
    $recordset.Fields.Item("ObjectLocation").Value = "$Ol"
    $recordset.update()
}

<# -- Close Database --- #>
$recordset.Close()
$conn.Close()