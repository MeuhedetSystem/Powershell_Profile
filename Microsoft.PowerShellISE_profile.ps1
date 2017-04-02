#add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
If (!(Get-module "activedirectory")) { Import-Module "activedirectory" }
$ExSession = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri ‘http://cas01.meuhedet.org/powershell?serializationLevel=Full'
Import-PSSession $ExSession

Set-Location D:\script
#Set-ExecutionPolicy -ExecutionPolicy Unrestricted
New-Alias -Name comm -Value .\comm.txt

#Last modified by offir.k 17.11.2016
function UID ($input1){
    if(!($input1)){
        $input1 = read-host "Please enter a username"
    }
    $input1 = $input1.Trim()
        $SearchID = "*"+$input1 ; $User = Get-ADUser -Filter {uid -like $SearchID} -Properties UID,OfficePhone,mobile| Select-Object samaccountname,UID,name,OfficePhone,mobile
        if($User){
             Write-Host "`nUID: " $User.uid "`nUser:" $User.samaccountname "`nName:" $User.name "`nmobile:" $User.mobile "`nOfficePhone:" $User.OfficePhone
             break
        }
        $sam = Get-Mailbox -Identity $input1 | Select-Object samaccountname
        $uid = Get-ADUser $sam.samaccountname -Properties UID,OfficePhone,mobile | select UID,name,OfficePhone,mobile
        Write-Host "`nUID: " $uid.uid "`nUser:" $sam.samaccountname "`nName:" $uid.name "`nmobile:" $uid.mobile "`nOfficePhone:" $uid.OfficePhone
}

function boot ($comp){
    $LastBootUpTime = Get-WmiObject Win32_OperatingSystem -Comp $comp | Select -Exp LastBootUpTime
    [System.Management.ManagementDateTimeConverter]::ToDateTime($LastBootUpTime)
}

function setATR15($user, $value){
    Set-ADUser $user -clear extensionattribute15
    Set-ADUser $user -Add @{"extensionattribute15"="$value"}
}


function ComputerNameByIP {
    param(
        $IPAddress = $null
    )
    BEGIN {
    }
    PROCESS {
        if ($IPAddress -and $_) {
            throw 'Please use either pipeline or input parameter'
            break
        } elseif ($IPAddress) {
            ([System.Net.Dns]::GetHostbyAddress($IPAddress)).HostName
        } elseif ($_) {
            [System.Net.Dns]::GetHostbyAddress($_).HostName
        } else {
            $IPAddress = Read-Host "Please supply the IP Address"
            [System.Net.Dns]::GetHostbyAddress($IPAddress).HostName
        }
    }
    END {
    }
}

<#
.SYNOPSIS
    DropDown UI - let user to choice element from list.

.DESCRIPTION
    function require array with opstions and return the choice or $flase.

.EXAMPLE
    Declare array
        $Array = "value1","value2","value3";
    Run function with param
        $result = Select-GuiDropDown -data $Array

.ToDo

.Last modified
    11/03/15 - grisha
#>

function Select-GuiDropDown ([array]$data, [string]$title){
    #gen clean form
    function GenWindow([string]$title, [int]$Width, [int]$Height){
        $form = New-Object System.Windows.Forms.Form
        $form.Text = $title
        $form.Width = $Width
        $form.Height = $Height
        $form.AutoSize = $true
        $form.StartPosition = "CenterScreen"
        return $form
    }

    #gen comnobox
    function GenDropDown([array]$data, [int]$x, [int]$y){
        $DropDown = New-Object System.Windows.Forms.ComboBox
        $DropDown.DataSource = @($data)
        $DropDown.Location  = New-Object System.Drawing.Point($x,$y)
	$DropDown.TabIndex = 0
        return $DropDown
    }

    #gen button
    function GenButton($window='', [string]$text, [int]$x, [int]$y, [scriptblock]$action=''){
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $text
        $button.Location = New-Object System.Drawing.Point($x,$y)
        if($window){$window.Controls.Add($button)}
        if($action){$button.Add_Click($action )}
        return $button
    }

    #valid return
    function return-combo($status){
        if ($status -eq 'OK'){
            $script:choice = $combo.SelectedItem.ToString()
            $window.close()
        }else{
            $script:choice = $false
            $window.close()
        }
    }

    $window = GenWindow -title $title -Width 200 -Height 120
    $combo = GenDropDown -data $data -x 40 -y 20
    $OkButton = GenButton -text 'OK' -x 10 -y 60 -action { return-combo -status 'OK' }
    $OkButton.TabIndex = "1"
    $CloseButton = GenButton -text 'Cancel' -x 100 -y 60 -action { return-combo -status 'cancel' }
    $CloseButton.TabIndex = "2"

    $window.Controls.Add($CloseButton)
    $window.Controls.Add($OkButton)
    $window.Controls.Add($combo)
    [void]$window.ShowDialog()
    return $choice
}
#Created by offir.k
#Last modified by offir.k 2.4.17

function Kill_Composit($SystemTeam = @($zahi,$offir,$eliav,$shlomi,$ron,$diana,$shahar)){
#region Users
    $zahi = New-Object PSObject -property @{
                                Name = "zahi"
                                Computer = "212"
                                }
    $offir = New-Object PSObject -property @{
                                Name = "offir"
                                Computer = "231"
                                }
    $eliav = New-Object PSObject -property @{
                                Name = "eliav"
                                Computer = "123"
                                }
    $shlomi = New-Object PSObject -property @{
                                Name = "shlomi"
                                Computer = "53"
                                }
    $ron = New-Object PSObject -property @{
                                Name = "ron"
                                Computer = "57"
                                }
    $diana = New-Object PSObject -property @{
                                Name = "diana"
                                Computer = "216"
                                }
    $shahar = New-Object PSObject -property @{
                                Name = "shahar"
                                Computer = "129"
                                }
#endregion
    $SysteamUsers = @($zahi,$offir,$eliav,$shlomi,$ron,$diana,$shahar)
    foreach($User in $SystemTeam){
        $User = $SysteamUsers -match $User
        taskkill /s ("10.0.1."+($User.Computer)) /im compositagentextender.exe /t /f
        taskkill /s ("10.0.1."+($User.Computer)) /im cimphoneagent.exe /t /f
    }
}

Function Convert-BytesToSize
{
<#
.SYNOPSIS
Converts any integer size given to a user friendly size.
.DESCRIPTION


Converts any integer size given to a user friendly size.

.PARAMETER size


Used to convert into a more readable format.
Required Parameter

.EXAMPLE


ConvertSize -size 134217728
Converts size to show 128MB

#>


#Requires -version 2.0


[CmdletBinding()]
Param
(
[parameter(Mandatory=$False,Position=0)][int64]$Size

)


#Decide what is the type of size
Switch ($Size)
{
{$Size -gt 1PB}
{
Write-Verbose “Convert to PB”
$NewSize = “$([math]::Round(($Size / 1PB),2))PB”
Break
}
{$Size -gt 1TB}
{
Write-Verbose “Convert to TB”
$NewSize = “$([math]::Round(($Size / 1TB),2))TB”
Break
}
{$Size -gt 1GB}
{
Write-Verbose “Convert to GB”
$NewSize = “$([math]::Round(($Size / 1GB),2))GB”
Break
}
{$Size -gt 1MB}
{
Write-Verbose “Convert to MB”
$NewSize = “$([math]::Round(($Size / 1MB),2))MB”
Break
}
{$Size -gt 1KB}
{
Write-Verbose “Convert to KB”
$NewSize = “$([math]::Round(($Size / 1KB),2))KB”
Break
}
Default
{
Write-Verbose “Convert to Bytes”
$NewSize = “$([math]::Round($Size,2))Bytes”
Break
}
}
Return $NewSize

}


function TestPort
{
    Param(
        [parameter(ParameterSetName='ComputerName', Position=0)]
        [string]
        $ComputerName,

        [parameter(ParameterSetName='IP', Position=0)]
        [System.Net.IPAddress]
        $IPAddress,

        [parameter(Mandatory=$true , Position=1)]
        [int]
        $Port,

        [parameter(Mandatory=$true, Position=2)]
        [ValidateSet("TCP", "UDP")]
        [string]
        $Protocol
        )

    $RemoteServer = If ([string]::IsNullOrEmpty($ComputerName)) {$IPAddress} Else {$ComputerName};

    If ($Protocol -eq 'TCP')
    {
        $test = New-Object System.Net.Sockets.TcpClient;
        Try
        {
            Write-Host "Connecting to "$RemoteServer":"$Port" (TCP)..";
            $test.Connect($RemoteServer, $Port);
            Write-Host "Connection successful";
        }
        Catch
        {
            Write-Host "Connection failed";
        }
        Finally
        {
            $test.Dispose();
        }
    }

    If ($Protocol -eq 'UDP')
    {
        $test = New-Object System.Net.Sockets.UdpClient;
        Try
        {
            Write-Host "Connecting to "$RemoteServer":"$Port" (UDP)..";
            $test.Connect($RemoteServer, $Port);
            Write-Host "Connection successful";
        }
        Catch
        {
            Write-Host "Connection failed";
        }
        Finally
        {
            $test.Dispose();
        }
    }
}
function Skype{
	$credential = Get-Credential "xoffir.k@meuhedet.org"
	$sessionskype = New-PSSession -ConnectionUri https://sfbpool.meuhedet.co.il/OcsPowershell -Credential ($credential)
	echo "Loading module..."
	Import-PSSession -Session $sessionskype
} 

#Last modified 7.2.17 by Offir.k
function NurseGroups{   
    function Nurses ($User){ 
        $NurseArray = "GRP_USR_HederTzevetGeneticScreening",`
                      "GRP_USR_HederTzevetNurses",`
                      "GRP_USR_HederTzevetPtzaim",`
                      "GRP_USR_HederTzevetStoma";

        $UsersGroups = Get-ADPrincipalGroupMembership -Identity $User |Select-Object name
        foreach($NurseGroup in $NurseArray){
            if($UsersGroups.name -contains $NurseGroup){
                Echo "User is already in $NurseGroup"
            }
            else{
                Add-ADGroupMember $NurseGroup $user
                Echo "User was added to $NurseGroup"
            }
        }
    }
    $Search = Select-GuiDropDown -data ("UID","Samaccountname","Hebrew name") -title "NursesGroups"
    $PARAMETER = Read-Host "Enter Parameter"
    if($Search -eq "UID"){
        $UID = "*"+$PARAMETER
        $UIDsearch = Get-ADUser -Filter {UID -like $UID } -Properties *|Select-Object  name, samaccountname
        $UIDsearch
        if(!($UIDsearch)){
            Write-Host "UID not found. You better check that." -ForegroundColor Red
            Break
        }
        $Question = Read-Host "do you wish to continue?(y/n)"
        if($Question -eq "y"){
            Nurses $UIDsearch.samaccountname
        }
    }
    if($Search -eq "Samaccountname"){
        Nurses $PARAMETER
    }
    if($Search -eq "Hebrew name"){
        $Name = $PARAMETER+"*"
        $Namesearch = Get-ADUser -Filter {name -like $Name } -Properties *|Select-Object  name, samaccountname
        if(!($Namesearch)){
            Write-Host "User not found in AD."
            break
        }
        $Namesearch
        $Question = Read-Host "do you wish to continue?(y/n)"
        if($Question -eq "y"){
            Nurses $Namesearch.samaccountname
        }
    }
}

#Created by offir.k 
#Last modified by offir.k 12.1.2017

function CurrentLoggedOnUser{
    Echo "Write hostname:"
    $comp = Read-Host
    $comp = $comp.Trim()
    "___________________________"
    if(Test-Connection $comp -Count 1 -Quiet){
	    #"Please wait..."
	    $CompInfo = Get-WmiObject -Class win32_computersystem -ComputerName $comp | Select-Object name,username
        #$CompInfo |fl
        "Computername: "+ $CompInfo.name
        "Username: "+ $CompInfo.username
        "___________________________"
        if(!$CompInfo.username){
            Write-Host "Computer is not in use."# -NoNewline
            break
        }
        if($CompInfo.username){
            "Users AD info:"       
            $UserInfor = uid ($CompInfo.username.trimstart("MEUHEDET\"))
            $UserInfor
        }
    }
    else{Write-Host "No ping to $comp" -ForegroundColor Red}
}

#Created by offir.k
#Last modified by offir.k 20.4.16
#àéôåñ ñéñîä ìîùúîù áà÷èéá
function ResetADUserPass {
    $Search = Select-GuiDropDown -data ("UID","Samaccountname","Hebrew name") -title "Find User"
        $PARAMETER = Read-Host "Enter Parameter"
        if($Search -eq "UID"){
            $UID = "*"+$PARAMETER
            $UIDsearch = Get-ADUser -Filter {UID -like $UID } -Properties *|Select-Object  name, samaccountname
            $result = $UIDsearch;$result
            if(!($result)){Echo "`nUser does not exist."}
        }
        elseif($Search -eq "Samaccountname"){
            $Samsearch = Get-ADUser $PARAMETER |Select-Object  name, samaccountname
            $result = $Samsearch;$result
            if(!($result)){Echo "`nUser does not exist."}
        }
        elseif($Search -eq "Hebrew name"){
            $Name = $PARAMETER+"*"
            $Namesearch = Get-ADUser -Filter {name -like $Name } -Properties *|Select-Object  name, samaccountname
            $result = $Namesearch;$result
            if(!($result)){Echo "`nUser does not exist."}
        }
        #else{Echo "`nUser does not exist."}

        if(!($result -eq $Null)){
            $newPassword = (Read-Host -Prompt "Provide New Password" -AsSecureString); Set-ADAccountPassword -Identity $result.SamAccountName -NewPassword $newPassword -Reset
            Set-ADUser -Identity $result.SamAccountName -ChangePasswordAtLogon $true
            $Sam = $result.SamAccountName
            echo "Password has been set for $Sam."
            $result.SamAccountName|clip ; echo "Username has been copied to clipboard."
    }
}

<#
     Created by Eliav.f
     Last modified by Eliav.f 1.3.2017

     Retrieve useful data From "Vm_Info.csv" about VMWare and Hyper-v Virtual Machines.
     The data is Updated on daily basis in 22:00	
     Create by Eliav 28.2.17
#>
 
 

Function Read-VM{
    $CSV = Import-Csv \\docserver1\SYSTEM_DOCS\files\VM_Info\VM_Info.csv
    $Read = Read-Host "Enter VMName or IP Address"
    $SearchName = $("*$Read*")
    
    If ($VMName = ($CSV | Where-Object {$_.VMName -like $SearchName -or $_.ComputerName -like $SearchName })){
        $VMName | Out-GridView -PassThru
       # $VMName    
    }
    ElseIf($IPAddress =  ($CSV | Where-Object {$_."IP Address" -like $Read})){
        $IPAddress | Out-GridView -PassThru
       # $IPAddress
    }
    Else{
        Write-Host "`nThere was no matches, please try again." -ForegroundColor Red
        Read-VM
    }
}
Clear-Host