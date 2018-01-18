
$date = (Get-Date).ToString("dd-MM-yyyy")

$inputXML = @"
<Window x:Name="Active_Directory_Tool" x:Class="AD_TOOL.MainWindow"
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:local="clr-namespace:AD_TOOL"
            mc:Ignorable="d"
            Title="Active Directory Tool" Width="772" Height="550" FontWeight="Bold">
    <Grid Margin="0,0,0,-21" Width="764" Height="537">
        <TabControl HorizontalAlignment="Left" VerticalAlignment="Top" Background="#FF263B99" Height="472" Width="758">
            <TabItem Header="User Information Export" FontWeight="Bold">
                <Grid Background="#FF06225F" HorizontalAlignment="Left" Margin="0,-2,-8,-1" Width="auto">
                    <Label x:Name="Select_Mode" Content="Select Mode" HorizontalAlignment="Left" Margin="158,30,0,0" VerticalAlignment="Top" Foreground="#FFFFFBFB"/>
                    <RadioButton x:Name="RadioSelective" Content="Selective Users Mode" HorizontalAlignment="Left" Margin="243,41,0,0" VerticalAlignment="Top" Background="White" Foreground="#FFF5EDED"/>
                    <RadioButton x:Name="Radioallusers" Content="All Users" HorizontalAlignment="Left" Margin="407,44,0,0" VerticalAlignment="Top" Foreground="#FFFBF7F7"/>
                    <Grid Background="White" Margin="0,64,0,185" Width="auto">
                        <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="198" VerticalAlignment="Top" Width="750"/>
                        <Label x:Name="Selectusers" Content="Select Users text file" HorizontalAlignment="Left" Margin="198,84,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                        <Button x:Name="btn_SelectiveBrowseFile" Content="Browse File" HorizontalAlignment="Left" Margin="362,87,0,0" VerticalAlignment="Top" Width="75" IsEnabled="False" FontWeight="Normal"/>
                        <Button x:Name="Btn_ExportSelective" Content="Export" HorizontalAlignment="Left" Margin="362,153,0,0" VerticalAlignment="Top" Width="75" IsEnabled="False" FontWeight="Normal"/>
                        <Label x:Name="Selective_Mode" Content="Selective Mode" HorizontalAlignment="Left" Margin="341,0,0,0" VerticalAlignment="Top" FontWeight="Bold" Height="23"/>

                        <TextBlock HorizontalAlignment="Left" Margin="478,70,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Normal" Foreground="#FFFF0707"><Run Text="Users List must have"/><Run FontWeight="Bold" Text=" "/><Run FontWeight="Bold" Text="SamaccountName "/><Run Text="or"/><LineBreak/><Run FontWeight="Bold" Text="Email "/><Run Text="or "/><Run FontWeight="Bold" Text="Employee ID"/></TextBlock>
                    </Grid>
                    <Grid>
                        <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="185" Margin="0,262,0,0" VerticalAlignment="Top" Width="750"/>
                        <Rectangle Height="185" Margin="0,262,0,0" Width="auto" Fill="#FFFDFDFF"/>
                        <Label Content="Do You want to export by OU" HorizontalAlignment="Left" Margin="166,312,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                        <RadioButton x:Name="Radio_OU_Yes" Content="Yes" HorizontalAlignment="Left" Margin="357,318,0,0" VerticalAlignment="Top" FontWeight="Normal" IsEnabled="False"/>
                        <RadioButton x:Name="Radio_OU_No" Content="No" HorizontalAlignment="Left" Margin="407,318,0,114" VerticalAlignment="Center" FontWeight="Normal" IsEnabled="False"/>
                        <Button x:Name="Btn_Export_All" Content="Export" HorizontalAlignment="Left" Margin="367,407,0,0" VerticalAlignment="Top" Width="75" IsEnabled="False" FontWeight="Normal"/>
                        <Label Content="All Users Export Mode" HorizontalAlignment="Left" Margin="322,262,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                        <TextBlock HorizontalAlignment="Left" Margin="10,2,0,0" TextWrapping="Wrap" Text="Script will export all attributes of the users" VerticalAlignment="Top" FontWeight="Bold" Foreground="#FF37EA25" Height="14"/>
                        <TextBlock HorizontalAlignment="Left" Margin="472,2,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="278" Height="25" Foreground="Yellow"  ><Run Text="Log Path :- "/><Run Text="C:\logs\AD_Tool_log_$date.txt" /></TextBlock>
                        <Button x:Name="Btn_SelectOU" Content="Select OU" HorizontalAlignment="Left" Margin="367,362,0,0" VerticalAlignment="Top" Width="75" IsEnabled="False" FontWeight="Normal"/>

                    </Grid>

                </Grid>
            </TabItem>
            <TabItem Header="AD Group Tool" FontWeight="Bold">
                <Grid Background="#FF434D91" Margin="0,0,0,-1">
                    <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="192" Margin="0,245,0,0" VerticalAlignment="Top" Width="752"/>

                    <Grid Margin="0,30,0,202" Background="White">
                        <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="215" VerticalAlignment="Top" Width="752" UseLayoutRounding="False" Cursor="None">
                            <Border.Effect>
                                <DropShadowEffect/>
                            </Border.Effect>
                        </Border>
                        <Label Content="Group Membership Tool" HorizontalAlignment="Left" Margin="292,-3,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                        <Label x:Name="GroupUsersList" Content="Users List :-" HorizontalAlignment="Left" Margin="254,42,0,0" VerticalAlignment="Top" Width="77" RenderTransformOrigin="0.584,0.115"/>
                        <Button x:Name="Btn_GroupUserList" Content="Browse" HorizontalAlignment="Left" Margin="340,45,0,0" VerticalAlignment="Top" Width="75" FontWeight="Normal"/>
                        <Label Content="Groups List:-" HorizontalAlignment="Left" Margin="254,73,0,0" VerticalAlignment="Top"/>
                        <Button x:Name="Btn_groupslist" Content="Browse" HorizontalAlignment="Left" Margin="340,76,0,0" VerticalAlignment="Top" Width="75" FontWeight="Normal" IsEnabled="False" RenderTransformOrigin="0.067,1.15"/>
                        <RadioButton x:Name="RadioAddtogroup" Content="Add to Groups" HorizontalAlignment="Left" Margin="262,123,0,0" VerticalAlignment="Top" FontWeight="Normal" IsEnabled="False"/>
                        <RadioButton x:Name="RadioGroupfromRemove" Content="Remove From Groups" HorizontalAlignment="Left" Margin="373,123,0,0" VerticalAlignment="Top" FontWeight="Normal" IsEnabled="False"/>
                        <Button x:Name="btn_GroupGo" Content="Go" HorizontalAlignment="Left" Margin="342,175,0,0" VerticalAlignment="Top" Width="75" FontWeight="Normal" IsEnabled="False"/>
                        <TextBlock HorizontalAlignment="Left" Margin="430,76,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="33" FontWeight="Normal" Foreground="#FFF91A1A" RenderTransformOrigin="0.849,-0.482"><Run Text="Group List Must Have"/><Run FontWeight="Bold" Text=" "/><Run FontWeight="Bold" Text="Samaccountname "/><Run Text="or "/><Run FontWeight="Bold" Text="DN "/><Run Text="of Goups"/><Run Text=" "/><Run Text=" "/></TextBlock>
                        <TextBlock HorizontalAlignment="Left" Margin="429,33,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontWeight="Normal" Foreground="#FFFD2121"><Run Text="Users List must have"/><Run FontWeight="Bold" Text=" "/><Run FontWeight="Bold" Text="S"/><Run FontWeight="Bold" Text="amaccountName "/><Run Text="or"/><LineBreak/><Run FontWeight="Bold" Text="Email "/><Run Text="or "/><Run FontWeight="Bold" Text="Employee ID"/></TextBlock>
                        <TextBlock HorizontalAlignment="Left" Margin="12,-25,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Foreground="#FFF5F513" Height="20"><Run Text="Log Path :- "/><Run Text="C:\logs\AD_Tool_log_$date.txt"/></TextBlock>
                    </Grid>
                    <Grid Margin="0,245,0,0" Background="White">
                        <Label Content="Export Members" HorizontalAlignment="Left" Margin="332,0,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.625,0.423"/>
                        <Label Content="Do You want export all Groups" HorizontalAlignment="Left" Margin="149,25,0,0" VerticalAlignment="Top"/>
                        <RadioButton x:Name="RadioExportallgroupsyes" Content="Yes" HorizontalAlignment="Left" Margin="341,31,0,0" VerticalAlignment="Top" FontWeight="Normal"/>
                        <RadioButton x:Name="RadioExportAllGroupsNO" Content="No" HorizontalAlignment="Left" Margin="382,31,0,0" VerticalAlignment="Top" FontWeight="Normal"/>
                        <Label x:Name="Groups_List" Content="Groups List" HorizontalAlignment="Left" Margin="259,56,0,0" VerticalAlignment="Top"/>
                        <Button x:Name="Btn_GroupBrowse_Export" Content="Browse" HorizontalAlignment="Left" Margin="342,59,0,118" VerticalAlignment="Center" Width="75" FontWeight="Normal" IsEnabled="False"/>
                        <TextBlock HorizontalAlignment="Left" Margin="486,46,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="256" Height="56" FontWeight="Normal" Foreground="#FFF91A1A" RenderTransformOrigin="0.849,-0.482"><Run Text="Group"/><Run Text="s"/><Run Text=" List Must Have "/><Run FontWeight="Bold" Text="Samaccountname "/><Run Text="or "/><Run FontWeight="Bold" Text="DN "/><Run Text="of Goups"/><Run Text=" "/><Run Text=" "/></TextBlock>
                    </Grid>
                    <Grid Margin="0,330,0,0" Background="White">
                        <RadioButton x:Name="Radio_Specific_OU_Yes" Content="Yes" HorizontalAlignment="Left" Margin="347,24,0,0" VerticalAlignment="Top" FontWeight="Normal" IsEnabled="False"/>
                        <RadioButton x:Name="Radio_Specific_OU_No" Content="No" HorizontalAlignment="Left" Margin="387,24,0,0" VerticalAlignment="Top" IsEnabled="False" FontWeight="Normal"/>
                        <Button x:Name="Btn_Export_Specific_OU" Content="Select OU" HorizontalAlignment="Left" Margin="347,55,0,0" VerticalAlignment="Top" Width="75" FontWeight="Normal" IsEnabled="False"/>
                        <Button x:Name="Btn_ExportMembers" Content="Export" HorizontalAlignment="Left" Margin="347,92,0,0" VerticalAlignment="Top" Width="75" FontWeight="Normal" IsEnabled="False"/>
                        <Label Content="Do You want Export From specific OU" HorizontalAlignment="Left" Margin="110,18,0,0" VerticalAlignment="Top"/>
                    </Grid>

                </Grid>
            </TabItem>
            <ListBox Height="100" Width="100"/>
        </TabControl>
        <Grid Margin="0,472,6,0">
            <ProgressBar x:Name="Info_progress" RenderTransformOrigin="0.7,1.167" Margin="567,1,4,30" Height="28" Width="187"/>
            <StatusBar x:Name="Info_status" Margin="4,2,194,29" FontFamily="Consolas" Height="28" FontWeight="Bold" Width="560"/>
        </Grid>
    </Grid>
</Window>

"@

$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML



$reader = (New-Object System.Xml.XmlNodeReader $XAML)

try { $Form = [Windows.Markup.XamlReader]::Load($reader) }
catch {

    #Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."

    throw
}
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================

$XAML.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) }


function Get-FormVariables {
    if ($global:ReadmeDisplay -ne $true) { Write-Host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow; $global:ReadmeDisplay = $true }
    Write-Host "Found the following interactable elements from our form" -ForegroundColor Cyan
    Get-Variable WPF*
}

Get-FormVariables




#===========================================================================
# Actually make the objects work
#===========================================================================

#Sample entry of how to add data to a field

#$vmpicklistView.items.Add([pscustomobject]@{'VMName'=($_).Name;Status=$_.Status;Other="Yes"})

$form.Add_Loaded( {


        $global:SaveFile = ""

        $global:SelectedOU = ""

        $global:getuserslist = ""

        $global:getgrouplist = ""

        $global:users = ""

        $global:groups = ""

        $global:samaccountnames = @()

        $global:GroupDN = ""

        function Global:Get-FileName ($initialDirectory) {
            [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
                Out-Null

            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory = $initialDirectory
            $OpenFileDialog.filter = "Text (*.txt)| *.txt"
            $OpenFileDialog.ShowDialog() | Out-Null
            $OpenFileDialog.filename
        }

        #Global:Write-Log function credit to lazywinadmin
        function Global:Write-Log {
            <#
       .SYNOPSIS
        Function to create or append a log file
       #>
            [CmdletBinding()]
            param(
                [Parameter()]
                $Path = "c:\logs",
                $LogName = "AD_Tool_log_$(Get-Date -f 'yyyyMMdd').log",

                [Parameter(Mandatory = $true)]
                $Message = "",

                [Parameter()]
                [ValidateSet('INFORMATIVE', 'SUCCESS', 'ERROR')]
                $Type = "INFORMATIVE",
                $Category
            )
            begin {
                # Verify if the log already exists, else create it
                if (-not (Test-Path -Path $(Join-Path -Path $Path -ChildPath $LogName))) {
                    New-Item -Path $(Join-Path -Path $Path -ChildPath $LogName) -ItemType file
                }

            }
            process {
                try {

                    switch ($TYPE) {
                        'INFORMATIVE' {

                            $WPFInfo_status.Items.Clear()
                            $WPFInfo_status.Background = "White"
                            #$WPFInfo_status.Items.Add($Message)
                            $Form.Dispatcher.Invoke([action]{$WPFInfo_status.Items.Add($Message)},"Render")
                            "$(Get-Date -Format dd-mm-yyyy:hh:mm) [$TYPE]  $Message" | Out-File -FilePath (Join-Path -Path $Path -ChildPath $LogName) -Append
                            #$Form.Dispatcher.Invoke([action]{$wWPFInfo_status}

                        }
                        'ERROR' {
                            $WPFInfo_status.Items.Clear()
                            $WPFInfo_status.Background = "Red"
                            $Form.Dispatcher.Invoke([action]{$WPFInfo_status.Items.Add($Message)},"Render")
                            "$(Get-Date -Format dd-mm-yyyy:hh:mm) [$TYPE]  $Message" | Out-File -FilePath (Join-Path -Path $Path -ChildPath $LogName) -Append
                        }
                        'SUCCESS' {

                            $WPFInfo_status.Items.Clear()
                            $WPFInfo_status.Background = "Green"
                            $Form.Dispatcher.Invoke([action]{$WPFInfo_status.Items.Add($Message)},"Render")
                            "$(Get-Date -Format dd-mm-yyyy:hh:mm) [$TYPE]  $Message" | Out-File -FilePath (Join-Path -Path $Path -ChildPath $LogName) -Append

                        }

                    }

                }
                catch {
                    Write-Error -Message "Could not write into $(Join-Path -Path $Path -ChildPath $LogName)"
                    Write-Error -Message "Last Error:$($error[0].exception.message)"
                }
            }

        }

        function Global:Get-SaveFile ($initialDirectory) {
            [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
                Out-Null

            $global:SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $global:SaveFileDialog.initialDirectory = $initialDirectory
            $global:SaveFileDialog.filter = "Csv (*.Csv)| *.csv"
            $global:SaveFileDialog.ShowDialog() | Out-Null
            $global:SaveFileDialog.filename
        }

        function Global:Get-OUDialog {
            <#
       .SYNOPSIS
        A self contained WPF/XAML treeview organizational unit selection dialog box.
       .DESCRIPTION
        A self contained WPF/XAML treeview organizational unit selection dialog box. No AD modules required, just need to be joined to the domain.
       .EXAMPLE
        $OU = Get-OUDialog
       .NOTES
        Author: Zachary Loeber
        Requires: Powershell 4.0
        Version History
        1.0.0 - 03/21/2015
        - Initial release (the function is a bit overbloated because I'm simply embedding some of my prior functions directly
        in the thing instead of customizing the code for the function. Meh, it gets the job done...
       .LINK
        https://github.com/zloeber/Powershell/blob/master/ActiveDirectory/Select-OU/Get-OUDialog.ps1
       .LINK
        http://www.the-little-things.net
       #>
            [CmdletBinding()]
            param()

            function Get-ChildOUStructure {
                <#
           .SYNOPSIS
            Create JSON exportable tree view of AD OU (or other) structures.
           .DESCRIPTION
            Create JSON exportable tree view of AD OU (or other) structures in Canonical Name format.
           .PARAMETER ouarray
            Array of OUs in CanonicalName format (ie. domain/ou1/ou2)
           .PARAMETER oubase
            Base of OU
           .EXAMPLE
            $OUs = @(Get-ADObject -Filter {(ObjectClass -eq "OrganizationalUnit")} -Properties CanonicalName).CanonicalName
            $test = $OUs | Get-ChildOUStructure | ConvertTo-Json -Depth 20
           .NOTES
            Author: Zachary Loeber
            Requires: Powershell 3.0, Lync
            Version History
            1.0.0 - 12/24/2014
            - Initial release
           .LINK
            https://github.com/zloeber/Powershell/blob/master/ActiveDirectory/Get-ChildOUStructure.ps1
           .LINK
            http://www.the-little-things.net
           #>
                [CmdletBinding()]
                param(
                    [Parameter(Position = 0, ValueFromPipeline = $true, Mandatory = $true, HelpMessage = 'Array of OUs in CanonicalName formate (ie. domain/ou1/ou2)')]
                    [string[]]$ouarray,
                    [Parameter(Position = 1, HelpMessage = 'Base of OU.')]
                    [string]$oubase = ''
                )
                begin {
                    $newarray = @()
                    $base = ''
                    $firstset = $false
                    $ouarraylist = @()
                }
                process {
                    $ouarraylist += $ouarray
                }
                end {
                    $ouarraylist = $ouarraylist | Where-Object { ($_ -ne $null) -and ($_ -ne '') } | Select-Object -Unique | Sort-Object
                    if ($ouarraylist.count -gt 0) {
                        $ouarraylist | ForEach-Object {
                            # $prioroupath = if ($oubase -ne '') {$oubase + '/' + $_} else {''}
                            $firstelement = @( $_ -split '/')[0]
                            $regex = "`^`($firstelement`?`)"
                            $tmp = $_ -replace $regex, '' -replace "^(\/?)", ''

                            if (-not $firstset) {
                                $base = $firstelement
                                $firstset = $true
                            }
                            else {
                                if (($base -ne $firstelement) -or ($tmp -eq '')) {
                                    Write-Verbose "Processing Subtree for: $base"
                                    $fulloupath = if ($oubase -ne '') { $oubase + '/' + $base } else { $base }
                                    New-Object psobject -Property @{
                                        'name'     = $base
                                        'path'     = $fulloupath
                                        'children' = if ($newarray.count -gt 0) {, @( Get-ChildOUStructure -ouarray $newarray -oubase $fulloupath) } else { $null }
                                    }
                                    $base = $firstelement
                                    $newarray = @()
                                    $firstset = $false
                                }
                            }
                            if ($tmp -ne '') {
                                $newarray += $tmp
                            }
                        }
                        Write-Verbose "Processing Subtree for: $base"
                        $fulloupath = if ($oubase -ne '') { $oubase + '/' + $base } else { $base }
                        New-Object psobject -Property @{
                            'name'     = $base
                            'path'     = $fulloupath
                            'children' = if ($newarray.count -gt 0) {, @( Get-ChildOUStructure -ouarray $newarray -oubase $fulloupath) } else { $null }
                        }
                    }
                }
            }

            function Connect-ActiveDirectory {
                [CmdletBinding()]
                param(
                    [Parameter(ParameterSetName = 'Credential')]
                    [Parameter(ParameterSetName = 'CredentialObject')]
                    [Parameter(ParameterSetName = 'Default')]
                    [string]$ComputerName,

                    [Parameter(ParameterSetName = 'Credential')]
                    [string]$DomainName,

                    [Parameter(ParameterSetName = 'Credential', Mandatory = $true)]
                    [string]$UserName,

                    [Parameter(ParameterSetName = 'Credential', HelpMessage = 'Password for Username in remote domain.', Mandatory = $true)]
                    [string]$Password,

                    [Parameter(ParameterSetName = 'CredentialObject', HelpMessage = 'Full credential object', Mandatory = $True)]
                    [System.Management.Automation.PSCredential]$Creds,

                    [Parameter(HelpMessage = 'Context to return, forest, domain, or DirectoryEntry.')]
                    [ValidateSet('Domain', 'Forest', 'DirectoryEntry', 'ADContext')]
                    [string]$ADContextType = 'ADContext'
                )

                $UsingAltCred = $false

                # If the username was passed in domain\<username> or username@domain then gank the domain name for later use
                if (($UserName -split "\\").count -gt 1) {
                    $DomainName = ($UserName -split "\\")[0]
                    $UserName = ($UserName -split "\\")[1]
                }
                if (($UserName -split "\@").count -gt 1) {
                    $DomainName = ($UserName -split "\@")[1]
                    $UserName = ($UserName -split "\@")[0]
                }

                switch ($PSCmdlet.ParameterSetName) {
                    'CredentialObject' {
                        if ($Creds.GetNetworkCredential().Domain -ne '') {
                            $UserName = $Creds.GetNetworkCredential().UserName
                            $Password = $Creds.GetNetworkCredential().Password
                            $DomainName = $Creds.GetNetworkCredential().Domain
                            $UsingAltCred = $true
                        }
                        else {
                            throw 'The credential object must include a defined domain.'
                        }
                    }
                    'Credential' {
                        if (-not $DomainName) {
                            Write-Error 'Username must be in @domainname.com or <domainname>\<username> format or the domain name must be manually passed in the DomainName parameter'
                            return $null
                        }
                        else {
                            $UserName = $DomainName + '\' + $UserName
                            $UsingAltCred = $true
                        }
                    }
                }

                $ADServer = ''

                # If a computer name was specified then we will attempt to perform a remote connection
                if ($ComputerName) {
                    # If a computername was specified then we are connecting remotely
                    $ADServer = "LDAP://$($ComputerName)"
                    $ContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::DirectoryServer

                    if ($UsingAltCred) {
                        $ADContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext $ContextType, $ComputerName, $UserName, $Password
                    }
                    else {
                        if ($ComputerName) {
                            $ADContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext $ContextType, $ComputerName
                        }
                        else {
                            $ADContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext $ContextType
                        }
                    }

                    try {
                        switch ($ADContextType) {
                            'ADContext' {
                                return $ADContext
                            }
                            'DirectoryEntry' {
                                if ($UsingAltCred) {
                                    return New-Object System.DirectoryServices.DirectoryEntry ($ADServer, $UserName, $Password)
                                }
                                else {
                                    return New-Object -TypeName System.DirectoryServices.DirectoryEntry $ADServer
                                }
                            }
                            'Forest' {
                                return [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ADContext)
                            }
                            'Domain' {
                                return [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ADContext)
                            }
                        }
                    }
                    catch {
                        throw
                    }
                }

                # If using just an alternate credential without specifying a remote computer (dc) to connect they
                # try connecting to the locally joined domain with the credentials.
                if ($UsingAltCred) {
                    # *** FINISH ME ***
                }
                # We have not specified another computer or credential so connect to the local domain if possible.
                $ContextType = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain
                try {
                    switch ($ADContextType) {
                        'ADContext' {
                            return New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext $ContextType
                        }
                        'DirectoryEntry' {
                            return [System.DirectoryServices.DirectoryEntry]''
                        }
                        'Forest' {
                            return [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
                        }
                        'Domain' {
                            return [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
                        }
                    }
                }
                catch {
                    throw
                }
            }

            function Search-AD {
                # Original Author (largely unmodified btw):
                #  http://becomelotr.wordpress.com/2012/11/02/quick-active-directory-search-with-pure-powershell/
                [CmdletBinding()]
                param(
                    [string[]]$Filter,
                    [string[]]$Properties = @( 'Name', 'ADSPath'),
                    [string]$SearchRoot = '',
                    [switch]$DontJoinAttributeValues,
                    [System.DirectoryServices.DirectoryEntry]$DirectoryEntry = $null
                )

                if ($DirectoryEntry -ne $null) {
                    if ($SearchRoot -ne '') {
                        $DirectoryEntry.set_Path($SearchRoot)
                    }
                }
                else {
                    $DirectoryEntry = [System.DirectoryServices.DirectoryEntry]$SearchRoot
                }

                if ($Filter) {
                    $LDAP = "(&({0}))" -f ($Filter -join ')(')
                }
                else {
                    $LDAP = "(name=*)"
                }
                try {
                    (New-Object System.DirectoryServices.DirectorySearcher -ArgumentList @(
                            $DirectoryEntry,
                            $LDAP,
                            $Properties
                        ) -Property @{
                            PageSize = 1000
                        }).FindAll() | ForEach-Object {
                        $ObjectProps = @{}
                        $_.Properties.GetEnumerator() |
                            ForEach-Object {
                            $Val = @( $_.Value)
                            if ($_.Name -ne $null) {
                                if ($DontJoinAttributeValues -and ($Val.count -gt 1)) {
                                    $ObjectProps.Add($_.Name, $_.Value)
                                }
                                else {
                                    $ObjectProps.Add($_.Name, ( -join $_.Value))
                                }
                            }
                        }
                        if ($ObjectProps.psbase.keys.count -ge 1) {
                            New-Object PSObject -Property $ObjectProps | Select-Object $Properties
                        }
                    }
                }
                catch {
                    Write-Warning -Message ('Search-AD: Filter - {0}: Root - {1}: Error - {2}' -f $LDAP, $Root.Path, $_.Exception.Message)
                }
            }

            function Convert-CNToDN {
                param([string]$CN)
                $SplitCN = $CN -split '/'
                if ($SplitCN.count -eq 1) {
                    return 'DC=' + (($SplitCN)[0] -replace '\.', ',DC=')
                }
                else {
                    $basedn = '.' + ($SplitCN)[0] -replace '\.', ',DC='
                    [array]::Reverse($SplitCN)
                    $ous = ''
                    for ($index = 0; $index -lt ($SplitCN.count - 1); $index++) {
                        $ous += 'OU=' + $SplitCN[$index] + ','
                    }
                    $result = ($ous + $basedn) -replace ',,', ','
                    return $result
                }
            }

            function Add-TreeItem {
                param(
                    $TreeObj,
                    $Name,
                    $Parent,
                    $Tag
                )

                $ChildItem = New-Object System.Windows.Controls.TreeViewItem
                $ChildItem.Header = $Name
                $ChildItem.Tag = $Tag
                $Parent.Items.Add($ChildItem) | Out-Null

                if (($TreeObj.children).count -gt 0) {
                    foreach ($ou in $TreeObj.children) {
                        $treeparent = Add-TreeItem -TreeObj $ou -Name $ou.Name -Parent $ChildItem -Tag $ou.Path
                    }
                }
            }

            if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
                Write-Warning 'Run PowerShell.exe with -Sta switch, then run this script.'
                Write-Warning 'Example:'
                Write-Warning '    PowerShell.exe -noprofile -Sta'
                break
            }

            [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
            [xml]$xamlMain = @'
<Window x:Name="windowSelectOU"
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
Title="Select OU" Height="367" Width="525">
<Grid>
<TreeView x:Name="treeviewOUs" Margin="10,10,10,46" UseLayoutRounding="False" Background="{DynamicResource {x:Static SystemColors.MenuBrushKey}}"/>
<Button x:Name="btnCancel" Content="Cancel" Margin="0,0,10,12" ToolTip="Filter" Height="23" VerticalAlignment="Bottom" HorizontalAlignment="Right" Width="71" IsCancel="True"/>
<Button x:Name="btnSelect" Content="Select" Margin="0,0,86,12" ToolTip="Filter" HorizontalAlignment="Right" Width="71" Height="23" VerticalAlignment="Bottom" IsDefault="True"/>
<TextBlock x:Name="txtSelectedOU" Margin="10,0,162,6" TextWrapping="Wrap" VerticalAlignment="Bottom" Height="35" Background="{DynamicResource {x:Static SystemColors.ControlBrushKey}}" IsEnabled="False"/>
</Grid>
</Window>
'@

            # Read XAML
            $reader = (New-Object System.Xml.XmlNodeReader $xamlMain)
            $window = [Windows.Markup.XamlReader]::Load($reader)

            $namespace = @{ x = 'http://schemas.microsoft.com/winfx/2006/xaml' }
            $xpath_formobjects = "//*[@*[contains(translate(name(.),'n','N'),'Name')]]"

            # Create a variable for every named xaml element
            Select-Xml $xamlMain -Namespace $namespace -XPath $xpath_formobjects | ForEach-Object {
                $_.Node | ForEach-Object {
                    Set-Variable -Name ($_.Name) -Value $window.FindName($_.Name)
                }
            }

            $conn = Connect-ActiveDirectory -ADContextType:DirectoryEntry
            $domstruct = @( Search-AD -DirectoryEntry $conn -Filter '(ObjectClass=organizationalUnit)' -Properties CanonicalName).CanonicalName | sort | Get-ChildOUStructure

            Add-TreeItem -TreeObj $domstruct -Name $domstruct.Name -Parent $treeviewOUs -Tag $domstruct.Path

            $treeviewOUs.add_SelectedItemChanged( {
                    $txtSelectedOU.Text = Convert-CNToDN $this.SelectedItem.Tag
                })

            $btnSelect.add_Click( {
                    $script:DialogResult = $txtSelectedOU.Text
                    $windowSelectOU.Close()
                })
            $btnCancel.add_Click( {
                    $script:DialogResult = $null
                })

            # Due to some bizarre bug with showdialog and xaml we need to invoke this asynchronously
            #  to prevent a segfault
            $async = $windowSelectOU.Dispatcher.InvokeAsync( {
                    $retval = $windowSelectOU.ShowDialog()
                })
            $async.Wait() | Out-Null

            # Clear out previously created variables for every named xaml element to be nice...
            Select-Xml $xamlMain -Namespace $namespace -XPath $xpath_formobjects | ForEach-Object {
                $_.Node | ForEach-Object {
                    Remove-Variable -Name ($_.Name)
                }
            }
            return $DialogResult
        }

        function Global:Getusername ($user) {
            try {

                $samaccountname = Get-ADUser -Filter { samaccountname -EQ $user -or mail -EQ $user -or employeeID -EQ $user } -Properties samaccountname

            }
            catch {

                Global:Write-Log -Path $logpath -Message "Error While collecting $user samaccountname" -Type ERROR

            }

            return $samaccountname
        }

        function Global:getgroupdn ($group) {
            try {

                $GroupDistinguishedName = Get-ADGroup -Filter { samaccountname -EQ $group -or DistinguishedName -EQ $group -or Name -EQ $group } -Properties DistinguishedName, Name

            }
            catch {

                Global:Write-Log -Path $logpath -Message "Error While collecting $user samaccountname" -Type ERROR

            }

            return $GroupDistinguishedName
        }

        $global:logpath = "c:\logs"

        $global:Attributes = "DisplayName", "City", "CN", "co", "codePage", "Company", "Country", "countryCode", "Department", "Description", "DistinguishedName", "EmailAddress", "EmployeeID", "EmployeeNumber", "Enabled", "Fax", "GivenName", "HomeDirectory", "HomedirRequired", "HomeDrive", "HomePage", "HomePhone", "Initials", "lastLogon", "LastLogonDate", "lastLogonTimestamp", "mail", "Manager", "mobile", "MobilePhone", "Name", "Office", "OfficePhone", "Organization", "physicalDeliveryOfficeName", "PostalCode", "SamAccountName", "sAMAccountType", "sn", "State", "StreetAddress", "Surname", "telephoneNumber", "Title", "UserPrincipalName"
        
        function Global:Export-Group {
            [CmdletBinding()]
            [Alias()]
            Param
            (
                # Param1 help description
                [Parameter(Mandatory = $false,
                    ValueFromPipelineByPropertyName = $true,
                    Position = 0)]
                $GroupNames,

                # Param2 help description
                [Parameter(Mandatory = $true)]
                [ValidateSet('All', 'GroupList')]
                $Type = "",

                [Parameter(Mandatory = $true,
                    ValueFromPipelineByPropertyName = $true)]
                $savefilepath,

                $SearchBaseOU
				
            )

            Begin {
            }
            Process {
                switch ($Type) {
                    "All" {
                        if ($SearchBaseOU) {
							
                            $F_getgroups = Get-ADGroup -Filter * -SearchBase $SearchBaseOU -Properties Name, DistinguishedName
							
                        }
                        Else {
								
                            $F_getgroups = Get-ADGroup -Filter * -Properties Name, DistinguishedName
                        }
							
                        $Report = @()
                        $count = 0
                        
                        foreach ($F_getgroup in $F_getgroups) {
                            try {
                                $Count ++
                                $F_GroupName = $null
                                $F_GroupName = $F_getgroup.Name
                               
                                $F_GroupDN = $null
                               
                                $F_GroupDN = $F_getgroup.DistinguishedName
                               
                                $memberlist = @()
                        
                                Global:Write-Log -Path $logpath -Message "Collecting $F_GroupName Members" -Type "INFORMATIVE"
                                
                                $collectcount = $(($Count/$F_getgroups.count)*100)

                                update-ProgressBar -PercentComplete $collectcount

                                $F_Members = Get-ADGroupMember -Identity $F_GroupDN
                                
                                foreach ($F_member in $F_Members) {
                                    $memberlist += $F_member.name
                                    [string]$F_Memberlist = $memberlist -join "`n"
                                }
                                
                                $csvObject = New-Object PSObject 
                                
                                Add-Member -inputObject $csvObject -memberType NoteProperty -name "Group" -value $F_GroupName
                                
                                Add-Member -inputObject $csvObject -memberType NoteProperty -name "Name" -value $F_Memberlist
                                

                                $Report += $csvObject
                            }
                            catch {
                                Global:Write-Log -Path $logpath -Message "Error While collecting $F_GroupName Members" -Type "Error"
                                $F_Memberlist = $_.exception.Message
                                $csvObject = New-Object PSObject 
                                Add-Member -inputObject $csvObject -memberType NoteProperty -name "Group" -value $F_GroupName
                                Add-Member -inputObject $csvObject -memberType NoteProperty -name "Name" -value $F_Memberlist
                                $Report += $csvObject
                                $ERRORCount = 1
                            }
						}
                            if ($ERRORCount -ne "0" -or $ERRORCount -eq $null -or $ERRORCount -eq '') 
                            {
                                
                                $Report | Export-Csv -NoTypeInformation -Path $savefilepath
                                
                                Global:Write-Log -Path $logpath -Message "Groups Exported to $savefilepath" -Type "Success"

                            }
                            else {
                                
                                Global:Write-Log -Path $logpath -Message "ERROR While Exporting" -Type "ERROR"
                            }
								
                        }
 
              "GroupList" {
							   
                        $getgroups = Get-Content -Path $global:getgrouplist

                        $Report = @()
                        $count = 0
                        foreach ($getgroup in $getgroups) {
                            try {
                                $count ++
                                $F_GroupDN = (Global:getgroupdn -group $getgroup).DistinguishedName
                                
                                $F_GroupName = (Global:getgroupdn -group $getgroup).Name

                                    
                                $memberlist = @()

                                Global:Write-Log -Path $logpath -Message "Collecting $F_GroupName Members" -Type "INFORMATIVE"

                                $CountNo  = $(($count/$getgroup.count)*100)

                                Global:update-ProgressBar -PercentComplete $CountNo

                                $F_Members = Get-ADGroupMember -Identity $F_GroupDN

                                foreach ($F_member in $F_Members) {
                                    $memberlist += $F_member.name
                                    [string]$F_Memberlist = $memberlist -join "`n"
                                }
                                $csvObject = New-Object PSObject 
                                Add-Member -inputObject $csvObject -memberType NoteProperty -name "Group" -value $F_GroupName
                                Add-Member -inputObject $csvObject -memberType NoteProperty -name "Name" -value $F_Memberlist
                                $Report += $csvObject
                            }
                            catch {
								
                                Global:Write-Log -Path $logpath -Message "Error While collecting $F_GroupName Members" -Type "Error"
                                $F_Memberlist = $_.exception.Message
                                $csvObject = New-Object PSObject 
                                Add-Member -inputObject $csvObject -memberType NoteProperty -name "Group" -value $F_GroupName
                                Add-Member -inputObject $csvObject -memberType NoteProperty -name "Name" -value $F_Memberlist
                                $Report += $csvObject
                                $ERRORCount = 1
							
                            }
                        }
					
                       if ($ERRORCount -ne "0" -or $ERRORCount -eq $null -or $ERRORCount -eq '') 
                            {
                            $Report | Export-Csv -NoTypeInformation -Path $savefilepath
                            Global:Write-Log -Path $logpath -Message "Groups Exported to $savefilepath" -Type "Success"
                        }
                        else {
                            Global:Write-Log -Path $logpath -Message "ERROR While Exporting" -Type "ERROR"
                        }
                    }

                }
            }
        }

        Function Global:update-ProgressBar([int]$PercentComplete )
        {
           
            $Form.Dispatcher.Invoke([action]{$WPFInfo_progress.Value = $PercentComplete},              
            "Render"
            

        )
        }
		
		
        Global:Write-Log -Path $logpath -Message "Ready" -Type INFORMATIVE


        Import-Module activedirectory -ErrorAction SilentlyContinue


    })

#### Users Information Export Control##########################
$WPFRadioSelective.add_Click( {

        Global:Write-Log -Path $logpath -Message "Ready" -Type INFORMATIVE
        $WPFbtn_SelectiveBrowseFile.IsEnabled = $true
        $WPFBtn_Export_All.IsEnabled = $false
        $WPFRadio_OU_Yes.IsEnabled = $false
        $WPFRadio_OU_No.IsEnabled = $false
        $WPFBtn_SelectOU.IsEnabled.$false

    })

$WPFRadioallusers.add_Click( {
        Global:Write-Log -Path $logpath -Message "Ready" -Type INFORMATIVE
        $WPFbtn_SelectiveBrowseFile.IsEnabled = $false
        $wpfBtn_ExportSelective.IsEnabled = $false
        $WPFRadio_OU_Yes.IsEnabled = $true
        $WPFRadio_OU_No.IsEnabled = $true
    })

$WPFbtn_SelectiveBrowseFile.add_Click( {

        $global:getuserslist = $null

        $global:getuserslist = Get-FileName


        if ($global:getuserslist.length -ne "0") {

            $wpfBtn_ExportSelective.IsEnabled = $true
            $WPFbtn_SelectiveBrowseFile.Content = "Selected"
            $WPFbtn_SelectiveBrowseFile.Background = "Green"

            Global:Write-Log -Path $logpath -Message "File Selected $global:getuserslist" -Type INFORMATIVE

        }


    })

$wpfBtn_ExportSelective.add_Click( {


        $global:SaveFile = Get-SaveFile

        Global:Write-Log -Path $logpath -Message "Getting Users from File" -Type INFORMATIVE

        $global:users = Get-Content $global:getuserslist

        $report = @()
        $count = 0
        foreach ($user in $global:users) {
            try {
                
                $count ++

                Global:Write-Log -Path $logpath -Message "Working on $user" -Type INFORMATIVE

                $getuserDetails = $null
                
                $Runcount = $(($Count/$global:users.count)*100)
                Global:update-ProgressBar -PercentComplete

                $getuserDetails = Get-ADUser -Filter { samaccountname -EQ $user -or mail -EQ $user -or employeeID -EQ $user } -Properties * | Select-Object $Attributes

                $csvObject = New-Object PSObject

                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "DisplayName" -Value $getuserDetails.DisplayName
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "City" -Value $getuserDetails.City
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "CN" -Value $getuserDetails.CN
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "co" -Value $getuserDetails.co
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "codePage" -Value $getuserDetails.codePage
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Company" -Value $getuserDetails.Company
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Country" -Value $getuserDetails.Country
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "countryCode" -Value $getuserDetails.countryCode
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Department" -Value $getuserDetails.Department
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Description" -Value $getuserDetails.Description
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "DistinguishedName" -Value $getuserDetails.DistinguishedName
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "EmailAddress" -Value $getuserDetails.EmailAddress
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "EmployeeID" -Value $getuserDetails.EmployeeID
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "EmployeeNumber" -Value $getuserDetails.EmployeeNumber
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Enabled" -Value $getuserDetails.Enabled
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Fax" -Value $getuserDetails.Fax
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "GivenName" -Value $getuserDetails.GivenName
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "HomeDirectory" -Value $getuserDetails.HomeDirectory
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "HomedirRequired" -Value $getuserDetails.HomedirRequired
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "HomeDrive" -Value $getuserDetails.HomeDrive
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "HomePage" -Value $getuserDetails.HomePage
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "HomePhone" -Value $getuserDetails.HomePhone
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Initials" -Value $getuserDetails.Initials
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "lastLogon" -Value $getuserDetails.lastLogon
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "LastLogonDate" -Value $getuserDetails.LastLogonDate
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "lastLogonTimestamp" -Value $getuserDetails.lastLogonTimestamp
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "mail" -Value $getuserDetails.mail
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Manager" -Value $getuserDetails.Manager
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "mobile" -Value $getuserDetails.mobile
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "MobilePhone" -Value $getuserDetails.MobilePhone
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Name" -Value $getuserDetails.Name
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Office" -Value $getuserDetails.Office
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "OfficePhone" -Value $getuserDetails.OfficePhone
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Organization" -Value $getuserDetails.Organization
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "physicalDeliveryOfficeName" -Value $getuserDetails.physicalDeliveryOfficeName
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "PostalCode" -Value $getuserDetails.PostalCode
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "SamAccountName" -Value $getuserDetails.SamAccountName
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "sAMAccountType" -Value $getuserDetails.sAMAccountType
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "sn" -Value $getuserDetails.sn
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "State" -Value $getuserDetails.State
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "StreetAddress" -Value $getuserDetails.StreetAddress
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Surname" -Value $getuserDetails.Surname
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "telephoneNumber" -Value $getuserDetails.telephoneNumber
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "Title" -Value $getuserDetails.Title
                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "UserPrincipalName" -Value $getuserDetails.UserPrincipalName

                $Report += $csvObject

            }
            catch {


                Global:Write-Log -Path $logpath -Message "Error while Collecting information for $user" -Type ERROR

                $csvObject = New-Object PSObject

                Add-Member -InputObject $csvObject -MemberType NoteProperty -Name "DisplayName" -Value "$user Error While collecting Information"

                $Report += $csvObject

            }
        }



        if ($report -ne $null) {

            $report | Export-Csv -Path $global:SaveFile -NoTypeInformation

            Global:Write-Log -Path $logpath -Message "File Successfully Export to $global:SaveFile" -Type SUCCESS

        }
        else {

            Global:Write-Log -Path $logpath -Message "$errorMsg" -Type ERROR
        }


        $global:users = $null

        $report = $null

        $global:SaveFile = $null

        $global:getuserslist = $null

        $WPFbtn_SelectiveBrowseFile.Content = "Browse"
        $WPFbtn_SelectiveBrowseFile.Background = "#FFDDDDDD"
    })

$WPFRadioallusers.add_click({
    $WPFRadio_OU_Yes.IsEnabled =  $true
    $WPFRadio_OU_No.IsEnabled = $true    
})

$WPFRadio_OU_Yes.add_Click( {

        $WPFBtn_SelectOU.IsEnabled = $true

    })

$WPFRadio_OU_No.add_Click( {

        $WPFBtn_SelectOU.IsEnabled = $false
        $WPFBtn_Export_All.IsEnabled = $true
    })

$WPFBtn_SelectOU.add_Click( {

        $global:SelectedOU = Get-OUDialog

        Global:Write-Log -Path $logpath -Message "$global:SelectedOU" -Type INFORMATIVE

        $WPFBtn_Export_All.IsEnabled = $true


    })

$WPFBtn_Export_All.add_Click( {

        $WPFBtn_Export_All.Content =  "Please Wait"
        $global:SaveFile = Get-SaveFile

        if ($global:SelectedOU -eq $null -or $global:SelectedOU.Length -eq "0" -or $global:SelectedOU -eq "" ) {


             try {
               
                Get-ADUser -Filter * -Properties * | Select-Object $Attributes | Export-Csv -Path $global:SaveFile -NoTypeInformation

                Global:Write-Log -Path $logpath -Message "All Users information exported" -Type SUCCESS
            }
            catch {

                $errorMsg = $_.Exception.Message

                Global:Write-Log -Path $logpath -Message "$errorMsg" -Type ERROR
            }


        }
        else {


                       try {
               
                Get-ADUser -Filter * -Properties * -SearchBase $global:SelectedOU | Select-Object $Attributes | Export-Csv -Path $global:SaveFile -NoTypeInformation

                Global:Write-Log -Path $logpath -Message "All Users information exported" -Type SUCCESS
            }
            catch {

                $errorMsg = $_.Exception.Message

                Global:Write-Log -Path $logpath -Message "$errorMsg" -Type ERROR
            }

        }

        $global:SaveFile = $null

        $global:SelectedOU = $null

        $global:getuserslist = $null

        $WPFBtn_Export_All.Content = "Exort"
    })

####END##########################

######AD Group Tool##############

$WPFBtn_GroupUserList.add_Click( {

        $global:getuserslist = $null

        $global:getuserslist = Get-FileName

        if ($global:getuserslist.length -ne "0") {

            $WPFBtn_groupslist.IsEnabled = $true
            $WPFBtn_GroupUserList.Content = "Selected"
            $WPFBtn_GroupUserList.Background = "Green" ##FFDDDDDD

            Global:Write-Log -Path $logpath -Message "File Selected $global:getuserslist" -Type INFORMATIVE

        }


    })

$WPFBtn_groupslist.add_Click( {

        $global:getgrouplist = $null

        $global:getgrouplist = Get-FileName

        if ($global:getgrouplist.length -ne "0") {

            Global:Write-Log -Path $logpath -Message "File Selected $global:getgrouplist" -Type INFORMATIVE

            $WPFBtn_groupslist.Background = "Green"
            $WPFBtn_groupslist.Content = "Selected"
            $WPFRadioAddtogroup.IsEnabled = $true
            $wpfRadioGroupfromRemove.IsEnabled = $true

        }


    })

$WPFRadioAddtogroup.add_Click( {

        $WPFbtn_GroupGo.IsEnabled = $true
    })

$wpfRadioGroupfromRemove.add_Click( {

        $WPFbtn_GroupGo.IsEnabled = $true
    })

$WPFbtn_GroupGo.add_Click( {


        if ($WPFRadioAddtogroup.IsChecked -eq $true) {

            $global:samaccountnames = @()

            $getuserscontents = $null

            Global:Write-Log -Path $logpath -Message "Working on Users Samaccountname" -Type INFORMATIVE

            $getuserscontents = Get-Content $global:getuserslist

            foreach ($getuserscontent in $getuserscontents) {
                try {

                    $getsamaccountname = $null
                    $getsamaccountname =
                    $global:samaccountnames += (Getusername -user $getuserscontent).SamAccountName

                }
                catch {
                    Global:Write-Log -Path $logpath -Message "could not collect $getuserscontent" -Type ERROR



                }

            }


            Global:Write-Log -Path $logpath -Message "Collecting Groups Name" -Type INFORMATIVE

            $getgroupscontents = Get-Content $global:getgrouplist

            foreach ($getgroupscontent in $getgroupscontents) {
                try {
                    $getgroupdn = $null

                    $getgroupdn = (getgroupdn -group $getgroupscontent)

                    $global:GroupDN = $null

                    $global:GroupDN = $getgroupdn.DistinguishedName

                    Global:Write-Log -Path $logpath -Message "Adding $global:samaccountnames to $getgroupscontent" -Type INFORMATIVE

                    Add-ADGroupMember -Identity $global:GroupDN -Members $global:samaccountnames -Confirm $false -ErrorAction Stop


                }
                catch {
                    Global:Write-Log -Path $logpath -Message "Unable to add Users to $getuserscontent" -Type ERROR

                    throw
                    $errorMsg = $_.Exception.Message

                    Global:Write-Log -Path $logpath -Message "$errorMsg" -Type ERROR
                }

                Global:Write-Log -Path $logpath -Message "Users Added to $getgroupscontent" -Type SUCCESS
            }

            $global:getuserslist = $null

            $global:getgrouplist = $null

            $global:users = $null

            $global:samaccountnames = $null

        }
        elseif ($wpfRadioGroupfromRemove.IsChecked -eq $true) {

            $global:samaccountnames = @()
            $getuserscontents = $null

            Global:Write-Log -Path $logpath -Message "Working on Users Samaccountname" -Type INFORMATIVE

            $getuserscontents = Get-Content $global:getuserslist

            foreach ($getuserscontent in $getuserscontents) {
                try {

                    $getsamaccountname = $null
                    $global:samaccountnames += (Getusername -user $getuserscontent).SamAccountName

                }
                catch {

                    Global:Write-Log -Path $logpath -Message "could not collect $getuserscontent" -Type ERROR

                }

            }



            Global:Write-Log -Path $logpath -Message "Collecting Groups Name" -Type INFORMATIVE

            $getgroupscontents = Get-Content $global:getgrouplist
            $errorCount = $null
            foreach ($getgroupscontent in $getgroupscontents) {
                try {
                    $getgroupdn = $null

                    $getgroupdn = (getgroupdn -group $getgroupscontent)

                    $global:GroupDN = $null

                    $global:GroupDN = $getgroupdn.DistinguishedName

                    Global:Write-Log -LogName $logpath -Message "Removing $global:samaccountnames to $getgroupscontent" -Type INFORMATIVE

                    Remove-ADGroupMember -Identity $global:GroupDN -Members $global:samaccountnames -Confirm:$false -ErrorAction Stop


                }
                catch {
                    Global:Write-Log -l $logpath -Message "Unable to remove Users from $getuserscontent" -Type ERROR

                    $errorMsg = $_.Exception.Message

                    Global:Write-Log -Path $logpath -Message "$errorMsg" -Type ERROR

                    $errorCount = "1"
                }

                if ($errorCount -eq "1") {

                    Global:Write-Log -Path $logpath -Message "Error while removing some users, Please check the log" -Type ERROR

                }
                else {

                    Global:Write-Log -Path $logpath -Message "Users removed from the groups" -Type ERROR
                }

            }

            $global:getuserslist = $null

            $global:getgrouplist = $null

            $global:users = $null

            $global:samaccountnames = $null
        }
        $WPFBtn_groupslist.Background = "#FFDDDDDD"

        $WPFBtn_GroupUserList.Background = "#FFDDDDDD"

        $WPFBtn_groupslist.Content = "Browse"

    })
	
	$wpfRadioExportallgroupsyes.add_Click( {
			
		$wpfBtn_GroupBrowse_Export.IsEnabled =  $false
		$WpfBtn_ExportMembers.IsEnabled = $false
		$WPFRadio_Specific_OU_Yes.IsEnabled = $true
		$WPFRadio_Specific_OU_No.IsEnabled = $true

		})

	$wpfRadioExportAllGroupsNO.add_Click( {
			
			$wpfBtn_GroupBrowse_Export.IsEnabled = $true
			$WpfBtn_ExportMembers.IsEnabled = $false
			$WPFRadio_Specific_OU_Yes.IsEnabled = $false
			$WPFRadio_Specific_OU_No.IsEnabled = $false
            $WPFBtn_Export_Specific_OU.IsEnabled = $false
		})

	$wpfBtn_GroupBrowse_Export.add_Click( {
        $global:getgrouplist = $null
        $global:getgrouplist = Get-FileName
        if ($global:getgrouplist.length -ne "0") {
			
			$WpfBtn_ExportMembers.IsEnabled = $true
			$wpfBtn_GroupBrowse_Export.Background = "Green"
			$wpfBtn_GroupBrowse_Export.Content = "Selected"
        }
    })
	
	$WPFRadio_Specific_OU_Yes.add_Click({
		
		$WPFBtn_Export_Specific_OU.IsEnabled = $true

	})

	$WPFRadio_Specific_OU_No.add_Click({

	
		$WPFBtn_Export_Specific_OU.IsEnabled = $false
		$WpfBtn_ExportMembers.IsEnabled = $true
	})
	
	$WPFBtn_Export_Specific_OU.add_Click({
		$Global:SelectedOU = $null
		$Global:SelectedOU = Global:Get-OUDialog
		if($Global:SelectedOU.Length -ne "0"){
			
			$WpfBtn_ExportMembers.IsEnabled =  $true
            $WPFBtn_Export_Specific_OU.Background = "Green"
            $WPFBtn_Export_Specific_OU.Content = "Selected"
			
		}
	})

	$WpfBtn_ExportMembers.add_Click( {
	   
        $WpfBtn_ExportMembers.Content = "Please wait"
		$Global:SaveFile = get-SaveFile
		
		if($wpfRadioExportallgroupsyes.IsChecked -eq $true)
		{
			
			if($WPFRadio_Specific_OU_Yes.IsChecked -eq $true)
			{
				Global:Export-Group -Type All -savefilepath $Global:SaveFile -SearchBase $Global:SelectedOU
			}
			elseif($WPFRadio_Specific_OU_No.IsChecked -eq $true)
			{
				Global:Export-Group -Type All -savefilepath $Global:SaveFile
			}
			
		}
		elseif ($wpfRadioExportAllGroupsNO.IsChecked -eq $true) {
			Global:Export-Group -GroupNames $global:getgrouplist -Type GroupList -savefilepath $Global:SaveFile 
			}
    $WPFBtn_Export_Specific_OU.Background = "#FFDDDDDD"
    
    $wpfBtn_GroupBrowse_Export.Background = "#FFDDDDDD"

    $WPFBtn_Export_Specific_OU.Content = "Select OU"
    
    $wpfBtn_GroupBrowse_Export.Content = "Browse"

    $WpfBtn_ExportMembers.Content = "Export"

    $WPFInfo_progress.Value
		
    })


#===========================================================================
# Shows the form
#===========================================================================
Write-Host "To show the form, run the following" -ForegroundColor Cyan
$Form.ShowDialog() | Out-Null