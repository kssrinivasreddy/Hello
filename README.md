<# 
.SYNOPSIS
    WinBot GUI program
This script is a GUI event based program to execute some of the run books used by MOI IT Windows.

Author: svinjarapu@micron.com
#> 

# .NET Framework classes
Add-Type -AssemblyName PresentationFramework

#Import functions
. "$PSScriptRoot\GetRFCListData.ps1"
. "$PSScriptRoot\Create_RFC_For_Patching.ps1"
. "$PSScriptRoot\StartPatching.ps1"
. "$PSScriptRoot\GetPatchingProgress.ps1"
. "$PSSCriptRoot\PerformPostPatchingValidation.ps1"
. "$PSScriptRoot\DeleteSnapshots.ps1"
. "$PSScriptRoot\UpgradeVMTools.ps1"
. "$PSScriptRoot\GetHARestartedEvents.ps1"
. "$PSScriptRoot\StartPostBuildActivities.ps1"
. "$PSScriptRoot\StartRetirement.ps1"

#Create Synchronized Hashtable for sharing variables between threads
$syncHash = [hashtable]::Synchronized(@{})

# Get XAML
[xml]$xaml = Get-Content "$PSScriptRoot\UI.xaml" -ErrorAction Stop

#Load XAML content in to XAML reader
$syncHash.Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml -ErrorAction Stop))

#Make GetRFCListData function available to Runspace
$Definition = Get-Content Function:\GetRFCListData -ErrorAction Stop
$SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'GetRFCListData', $Definition

#Make CreateRFCs function available to Runspace
$CreateRFCsDefinition = Get-Content Function:\CreateRFCs -ErrorAction Stop
$CreateRFCsFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'CreateRFCs', $CreateRFCsDefinition

#Make StartPatching function available to Runspace
$StartPatchingDefinition = Get-Content Function:\StartPatching -ErrorAction Stop
$StartPatchingFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'StartPatching', $StartPatchingDefinition

#Make GetPatchingProgress function available to Runspace
$GetPatchingProgressDefinition = Get-Content Function:\GetPatchingProgress -ErrorAction Stop
$GetPatchingProgressFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'GetPatchingProgress', $GetPatchingProgressDefinition

#Make PerformPostPatchingValidation function available to Runspace
$PerformPostPatchingValidationDefinition = Get-Content Function:\PerformPostPatchingValidation -ErrorAction Stop
$PerformPostPatchingValidationFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'PerformPostPatchingValidation', $PerformPostPatchingValidationDefinition

$DeleteSnapshotsDefinition = Get-Content Function:\DeleteSnapshots -ErrorAction Stop
$DeleteSnapshotsFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'DeleteSnapshots', $DeleteSnapshotsDefinition

$UpgradeVMToolsDefinition = Get-Content Function:\UpgradeVMTools -ErrorAction Stop
$UpgradeVMToolsFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'UpgradeVMTools', $UpgradeVMToolsDefinition

$GetHARestartedEventsDefinition = Get-Content Function:\GetHARestartedEvents -ErrorAction Stop
$GetHARestartedEventsFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'GetHARestartedEvents', $GetHARestartedEventsDefinition

$StartPostBuildActivitiesDefinition = Get-Content Function:\StartPostBuildActivities -ErrorAction Stop
$StartPostBuildActivitiesFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'StartPostBuildActivities', $StartPostBuildActivitiesDefinition

$StartRetirementDefinition = Get-Content Function:\StartRetirement -ErrorAction Stop
$StartRetirementFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'StartRetirement', $StartRetirementDefinition

#Create a SessionStateFunction
$InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$InitialSessionState.Commands.Add($SessionStateFunction)
$InitialSessionState.Commands.Add($CreateRFCsFunction)
$InitialSessionState.Commands.Add($StartPatchingFunction)
$InitialSessionState.Commands.Add($GetPatchingProgressFunction)
$InitialSessionState.Commands.Add($PerformPostPatchingValidationFunction)
$InitialSessionState.Commands.Add($DeleteSnapshotsFunction)
$InitialSessionState.Commands.Add($UpgradeVMToolsFunction)
$InitialSessionState.Commands.Add($GetHARestartedEventsFunction)
$InitialSessionState.Commands.Add($StartPostBuildActivitiesFunction)
$InitialSessionState.Commands.Add($StartRetirementFunction)

#Create variables for the elements in the UI (Buttons and Text boxes)
$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | Where-Object { ($_.TargetName -ne "Border") -and ($_.Name -ne "Border") -and ($_.Name -ne "ContentSite") } | ForEach-Object {
    $syncHash.Add($_.Name,$syncHash.Window.FindName($_.Name))
}

$SyncHash.CreateRFCs.Add_Click({
        #Check if Remedy credentails are entered
        if (($syncHash.RemedyUserName.Text.Length -eq 0) -or ($syncHash.RemedyPassword.Password.Length -eq 0)) {
         [System.Windows.MessageBox]::Show("Please enter Remedy user name and password","Credentials Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
         return
        }        
        $SyncHash.CreateRFCs.IsEnabled = $false
        $syncHash.Host = $host
        $Runspace = [runspacefactory]::CreateRunspace($InitialSessionState)
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()
        $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
        $Runspace.SessionStateProxy.SetVariable("remedyUserName",$SyncHash.RemedyUserName.Text)
        $Runspace.SessionStateProxy.SetVariable("remedyPassword",$SyncHash.RemedyPassword.Password)
        #Wait-Debugger
        $code1 = {
            #Wait-Debugger
            CreateRFCs -syncHash $syncHash -remedyUserName $remedyUserName -remedyPassword $remedyPassword
            $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.CreateRFCs.IsEnabled = $true } )           
        }
        $PSinstance = [powershell]::Create().AddScript($Code1)
        $PSinstance.Runspace = $Runspace
        $job = $PSinstance.BeginInvoke()
})

$syncHash.GetRFCDataFromSharePoint.Add_Click({    
    $syncHash.GetRFCDataFromSharePoint.IsEnabled = $false
    $syncHash.Host = $host
    $Runspace = [runspacefactory]::CreateRunspace($InitialSessionState)
    $PSinstance = [powershell]::Create()
    $PSinstance.Runspace = $Runspace
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)  
    $code1 = {
        #Wait-Debugger
        $syncHash.Window.Dispatcher.invoke( [action]{ $syncHash.RFCList.Items.Clear() } )
        $syncHash.Window.Dispatcher.invoke( [action]{ $syncHash.RFCOutputBox.Text = "`nCollecting Data from SharePoint. Please wait.." } )       
        $RFC_To_Be_Created_List = GetRFCListData
        if ($RFC_To_Be_Created_List.Count -eq 0) {
            $syncHash.Window.Dispatcher.invoke( [action]{ $syncHash.RFCOutputBox.Text = "`nNo servers found in SharePoint for which an RFC needs to be created" } )
            $syncHash.Window.Dispatcher.invoke( [action]{ $syncHash.GetRFCDataFromSharePoint.IsEnabled = $true })
            return
        }
        foreach ($RFCObject in $RFC_To_Be_Created_List) {
            $RFCObject.StartTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId( ($RFCObject.StartTime), 'Mountain Standard Time')
            $RFCObject.EndTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId( ($RFCObject.EndTime), 'Mountain Standard Time')
            $syncHash.Window.Dispatcher.invoke( [action]{ $syncHash.RFCList.Items.Add($RFCObject) } )
        }
        $syncHash.Window.Dispatcher.invoke( [action]{ $syncHash.RFCOutputBox.Text = "`nData Gathering from SharePoint is completed" } )
        $syncHash.Window.Dispatcher.invoke( [action]{ $syncHash.GetRFCDataFromSharePoint.IsEnabled = $true })
    }
    [void]$PSinstance.AddScript($Code1)
    $job = $PSinstance.BeginInvoke()
})

$SyncHash.StartPatching.Add_Click({
        if ($syncHash.PostPatchingValidation.IsEnabled -eq $false) {
            [System.Windows.MessageBox]::Show("Post Patching validation in progress. Please wait","Validation in progress",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
            return
        }
        if ($syncHash.ServersToBePatched.Text -eq "") {
            [System.Windows.MessageBox]::Show("Please enter Server Name(s)","Server Name(s) Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
            return
        }        
        if ($syncHash.TakeSnapshot.IsChecked -eq $true) {
            if (($syncHash.adUserName.Text -eq "") -or ($syncHash.adPassword.Password -eq "")) {
                [System.Windows.MessageBox]::Show("Please enter ad- credentials to connect to vCenter","ad- Credentials Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
                return
            }
            $syncHash.TakeSnapshot.IsEnabled = $false            
        }
        $SyncHash.StartPatching.IsEnabled = $false
        $syncHash.Host = $host
        $Runspace = [runspacefactory]::CreateRunspace($InitialSessionState)
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()
        $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
        $Runspace.SessionStateProxy.SetVariable("ServerNames",$SyncHash.ServersToBePatched.Text)     
        $Runspace.SessionStateProxy.SetVariable("adUserName",$SyncHash.adUserName.Text)
        $Runspace.SessionStateProxy.SetVariable("adPassword",$SyncHash.adPassword.Password)
        $Runspace.SessionStateProxy.SetVariable("takeSnapshot",$SyncHash.TakeSnapshot.IsChecked)
        $Runspace.SessionStateProxy.SetVariable("autoReboot",$SyncHash.AutoReboot.IsChecked)
        #Wait-Debugger
        $code1 = {
            #Wait-Debugger
            StartPatching -syncHash $syncHash -serverNames $ServerNames -adUserName $adUserName -adPassword $adPassword -takeSnapshot $takeSnapshot -autoReboot $autoReboot
            $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.StartPatching.IsEnabled = $true })
            $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.TakeSnapshot.IsEnabled = $true })
            $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.TakeSnapshot.IsChecked = $false })
            $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.ServersToBePatched.Clear() })
        }
        $PSinstance = [powershell]::Create().AddScript($Code1)
        $PSinstance.Runspace = $Runspace
        $job = $PSinstance.BeginInvoke()        
})

$syncHash.ShowDetailedStatus.Add_Click({
    if ($syncHash.ServersToBePatched.Text -eq "") {
            [System.Windows.MessageBox]::Show("Please enter Server Name(s)","Server Name(s) Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
            return
    }    
    $SyncHash.ShowDetailedStatus.IsEnabled = $false
    $syncHash.PatchingStatus.Visibility = "Visible"
    $syncHash.PatchingStatus.IsSelected = $true    
    #$syncHash.PatchingProgressBar.Value=50
    $syncHash.Host = $host
    $Runspace = [runspacefactory]::CreateRunspace($InitialSessionState)
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $Runspace.SessionStateProxy.SetVariable("ServerNames",$SyncHash.ServersToBePatched.Text)        
    #Wait-Debugger
    $code1 = {
        #Wait-Debugger
        GetPatchingProgress -syncHash $syncHash -serverNames $ServerNames
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.ShowDetailedStatus.IsEnabled = $true } )           
        }
    $PSinstance = [powershell]::Create().AddScript($Code1)
    $PSinstance.Runspace = $Runspace
    $job = $PSinstance.BeginInvoke()    
})

$syncHash.PostPatchingValidation.Add_Click({
    if ($syncHash.StartPatching.IsEnabled -eq $false) {
        [System.Windows.MessageBox]::Show("Patching in progress. Please wait","Patching in progress",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    if ($syncHash.ServersToBePatched.Text -eq "") {
            [System.Windows.MessageBox]::Show("Please enter Server Name(s)","Server Name(s) Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
            return
    }
    if (($syncHash.adUserName.Text -eq "") -or ($syncHash.adPassword.Password -eq "")) {
        [System.Windows.MessageBox]::Show("Please enter ad- credentials to connect to vCenter","ad- Credentials Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    $SyncHash.PostPatchingValidation.IsEnabled = $false
    $syncHash.Host = $host
    $Runspace = [runspacefactory]::CreateRunspace($InitialSessionState)
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $Runspace.SessionStateProxy.SetVariable("ServerNames",$SyncHash.ServersToBePatched.Text)
    $Runspace.SessionStateProxy.SetVariable("adUserName",$SyncHash.adUserName.Text)
    $Runspace.SessionStateProxy.SetVariable("adPassword",$SyncHash.adPassword.Password)
    #Wait-Debugger
    $code1 = {
        #Wait-Debugger
        PerformPostPatchingValidation -syncHash $syncHash -serverNames $ServerNames -adUserName $adUserName -adPassword $adPassword
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.PostPatchingValidation.IsEnabled = $true })
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.ServersToBePatched.clear()})
        }
    $PSinstance = [powershell]::Create().AddScript($Code1)
    $PSinstance.Runspace = $Runspace
    $job = $PSinstance.BeginInvoke()
})

$syncHash.DeleteSnapshots.Add_Click({
    if ($syncHash.GetHARestartedVMs.IsEnabled -eq $false) {
        [System.Windows.MessageBox]::Show("Get HA Restarted VMs is in progress. Please try after some time","Get HA Restarted VMs in progress",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    if ($syncHash.UpgradeVMTools.IsEnabled -eq $false) {
        [System.Windows.MessageBox]::Show("Upgrade VMTools is in progress. Please try after some time","Upgrade VMTools in progress",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    if ($syncHash.VMNames.Text -eq "") {
            [System.Windows.MessageBox]::Show("Please enter VM Name(s)","VM Name(s) Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
            return
    }
    if (($syncHash.ADUserName3.Text -eq "") -or ($syncHash.ADPassword3.Password -eq "")) {
        [System.Windows.MessageBox]::Show("Please enter ad- credentials to connect to vCenter","ad- Credentials Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    $SyncHash.DeleteSnapshots.IsEnabled = $false    
    $syncHash.Host = $host
    $Runspace = [runspacefactory]::CreateRunspace($InitialSessionState)
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $Runspace.SessionStateProxy.SetVariable("ServerNames",$SyncHash.VMNames.Text)
    $Runspace.SessionStateProxy.SetVariable("adUserName",$SyncHash.ADUserName3.Text)
    $Runspace.SessionStateProxy.SetVariable("adPassword",$SyncHash.ADPassword3.Password)
    #Wait-Debugger
    $code1 = {
        #Wait-Debugger
        DeleteSnapshots -syncHash $syncHash -VMNames $ServerNames -adUserName $adUserName -adPassword $adPassword
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.DeleteSnapshots.IsEnabled = $true })        
    }
    $PSinstance = [powershell]::Create().AddScript($Code1)
    $PSinstance.Runspace = $Runspace
    $job = $PSinstance.BeginInvoke()
})

$syncHash.UpgradeVMTools.Add_Click({
    if ($syncHash.GetHARestartedVMs.IsEnabled -eq $false) {
        [System.Windows.MessageBox]::Show("Get HA Restarted VMs is in progress. Please try after some time","Get HA Restarted VMs in progress",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    if ($syncHash.DeleteSnapshots.IsEnabled -eq $false) {
        [System.Windows.MessageBox]::Show("Delete Snapshots is in progress. Please try after some time","Delete Snapshots in progress",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    if ($syncHash.VMNames.Text -eq "") {
            [System.Windows.MessageBox]::Show("Please enter VM Name(s)","VM Name(s) Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
            return
    }
    if (($syncHash.ADUserName3.Text -eq "") -or ($syncHash.ADPassword3.Password -eq "")) {
        [System.Windows.MessageBox]::Show("Please enter ad- credentials to connect to vCenter","ad- Credentials Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    $SyncHash.UpgradeVMTools.IsEnabled = $false    
    $syncHash.Host = $host
    $Runspace = [runspacefactory]::CreateRunspace($InitialSessionState)
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $Runspace.SessionStateProxy.SetVariable("ServerNames",$SyncHash.VMNames.Text)
    $Runspace.SessionStateProxy.SetVariable("adUserName",$SyncHash.ADUserName3.Text)
    $Runspace.SessionStateProxy.SetVariable("adPassword",$SyncHash.ADPassword3.Password)
    #Wait-Debugger
    $code1 = {
        #Wait-Debugger
        UpgradeVMTools -syncHash $syncHash -VMNames $ServerNames -adUserName $adUserName -adPassword $adPassword
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.UpgradeVMTools.IsEnabled = $true })        
    }
    $PSinstance = [powershell]::Create().AddScript($Code1)
    $PSinstance.Runspace = $Runspace
    $job = $PSinstance.BeginInvoke()
})

$syncHash.GetHARestartedVMs.Add_Click({
    if ($syncHash.DeleteSnapshots.IsEnabled -eq $false) {
        [System.Windows.MessageBox]::Show("Delete Snapshots is in progress. Please try after some time","Delete Snapshots in progress",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    if ($syncHash.UpgradeVMTools.IsEnabled -eq $false) {
        [System.Windows.MessageBox]::Show("Upgrade VMTools is in progress. Please try after some time","Upgrade VMTools in progress",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }    
    if (($syncHash.ADUserName3.Text -eq "") -or ($syncHash.ADPassword3.Password -eq "")) {
        [System.Windows.MessageBox]::Show("Please enter ad- credentials to connect to vCenter","ad- Credentials Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    $SyncHash.GetHARestartedVMs.IsEnabled = $false 
    $syncHash.Host = $host
    $Runspace = [runspacefactory]::CreateRunspace($InitialSessionState)
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)    
    $Runspace.SessionStateProxy.SetVariable("adUserName",$SyncHash.ADUserName3.Text)
    $Runspace.SessionStateProxy.SetVariable("adPassword",$SyncHash.ADPassword3.Password)
    #Wait-Debugger
    $code1 = {
        #Wait-Debugger
        GetHARestartedEvents -syncHash $syncHash -adUserName $adUserName -adPassword $adPassword
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.GetHARestartedVMs.IsEnabled = $true })        
    }
    $PSinstance = [powershell]::Create().AddScript($Code1)
    $PSinstance.Runspace = $Runspace
    $job = $PSinstance.BeginInvoke()
})

$syncHash.AddToServerList.Add_Click({
    if ($syncHash.PostBuildServerName.Text -eq "") {
        [System.Windows.MessageBox]::Show("Please enter server name","Server name required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    $serverObject = [PSCustomObject]@{
        ServerName = $syncHash.PostBuildServerName.Text
        ServerTier = $syncHash.ServerTier.Text
    }
    $syncHash.ServerList.Items.Add($serverObject)
    $syncHash.PostBuildServerName.Clear()
    $serverNames = $syncHash.ServerList.Items
})

$syncHash.RemoveFromServerList.Add_Click({
    
    $item = $syncHash.ServerList.SelectedItem
    $syncHash.ServerList.Items.Remove($item)
    
})

$syncHash.StartPostBuildActivities.Add_Click({
    if (($syncHash.ADUserName1.Text -eq "") -or ($syncHash.ADPassword1.Password -eq "")) {
        [System.Windows.MessageBox]::Show("Please enter ad- credentials to connect to vCenter","ad- Credentials Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    if ($syncHash.ServerList.Items.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No server added to Server List","Empty Server List",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    $SyncHash.StartPostBuildActivities.IsEnabled = $false
    $SyncHash.AddToServerList.IsEnabled = $false
    $SyncHash.RemoveFromServerList.IsEnabled = $false
    $syncHash.Host = $host
    $Runspace = [runspacefactory]::CreateRunspace($InitialSessionState)
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $Runspace.SessionStateProxy.SetVariable("serverList",$syncHash.ServerList)
    $Runspace.SessionStateProxy.SetVariable("adUserName",$SyncHash.ADUserName1.Text)
    $Runspace.SessionStateProxy.SetVariable("adPassword",$SyncHash.ADPassword1.Password)
    #Wait-Debugger
    $code1 = {
        #Wait-Debugger
        StartPostBuildActivities -syncHash $syncHash -ServerList $serverList -adUserName $adUserName -adPassword $adPassword
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.StartPostBuildActivities.IsEnabled = $true })
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.AddToServerList.IsEnabled = $true })
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.RemoveFromServerList.IsEnabled = $true })
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.ServerList.Items.Clear()})
    }
    $PSinstance = [powershell]::Create().AddScript($Code1)
    $PSinstance.Runspace = $Runspace
    $job = $PSinstance.BeginInvoke()
})

$syncHash.StartRetirement.Add_Click({
    if (($syncHash.DecomServerName.Text -eq "") -or ($syncHash.DecomRFCNumber.Text -eq "") -or ($syncHash.DecomJiraCase.Text -eq "")) {
        [System.Windows.MessageBox]::Show("Please enter Server Name, RFC Number and Jira case","Mandatory details missing",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    if (($syncHash.ADUserName2.Text -eq "") -or ($syncHash.ADPassword2.Password -eq "")) {
        [System.Windows.MessageBox]::Show("Please enter ad- credentials","ad- Credentials Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    if (($syncHash.RemedyUserName1.Text -eq "") -or ($syncHash.RemedyPassword1.Password -eq "")) {
        [System.Windows.MessageBox]::Show("Please enter Remedy credentials","Remedy Credentials Required",[System.Windows.MessageBoxButton]::Ok,[System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $SyncHash.StartRetirement.IsEnabled = $false
    $syncHash.Host = $host
    $Runspace = [runspacefactory]::CreateRunspace($InitialSessionState)
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $Runspace.SessionStateProxy.SetVariable("serverList",$syncHash.DecomServerName.Text)
    $Runspace.SessionStateProxy.SetVariable("adUserName",$SyncHash.ADUserName2.Text)
    $Runspace.SessionStateProxy.SetVariable("adPassword",$SyncHash.ADPassword2.Password)
    $Runspace.SessionStateProxy.SetVariable("remedyUserName",$SyncHash.RemedyUserName1.Text)
    $Runspace.SessionStateProxy.SetVariable("remedyPassword",$SyncHash.RemedyPassword1.Password)
    $Runspace.SessionStateProxy.SetVariable("RFCNumber",$SyncHash.DecomRFCNumber.Text)
    $Runspace.SessionStateProxy.SetVariable("jiraCase",$SyncHash.DecomJiraCase.Text)
    #Wait-Debugger
    $code1 = {
        #Wait-Debugger
        StartRetirement -syncHash $syncHash -ServerList $serverList -adUserName $adUserName -adPassword $adPassword -remedyUserName $remedyUserName -remedyPassword $remedyPassword -RFCNumber $RFCNumber -jiraCase $jiracase
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.StartRetirement.IsEnabled = $true })
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.DecomServerName.Clear() })
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.DecomRFCNumber.Clear() })
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.DecomJiraCase.Clear() })
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.ADUserName2.Clear() })
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.ADPassword2.Clear() })
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.RemedyUserName1.Clear() })
        $SyncHash.Window.Dispatcher.invoke( [action]{ $syncHash.RemedyPassword1.Clear() })
    }
    $PSinstance = [powershell]::Create().AddScript($Code1)
    $PSinstance.Runspace = $Runspace
    $job = $PSinstance.BeginInvoke()
})


#Show the GUI
$syncHash.Window.ShowDialog() | Out-Null
