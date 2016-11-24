#========================================================================
# Name		: Hyper-V_VMConvertToStandard.ps1
# Author 	: Florian Clisson
# Email	    : florian.clisson@gmail.com
# 
#========================================================================

#****************************************************************************************************
<#  
.SYNOPSIS  
    Applies the Standard regarding the Virtual machine after convertion VMware to Hyper-V.  
.DESCRIPTION  
    
    This script does : 
      1 - Detect the not Standard Virtual Machines      
      2 - Move the Virtual Machine in the standard location : "C:\ClusterStorage\<DataStoreName>\<VMname>\Virtual Hard Disks\"
      3 - Rename the VHDX File with the standard <VMName>.vhdx and <VMName>_1.vhdx
      4 - Find the Virtual Machine on VMware environment and get the Memory
      5 - Disabled the dynamic memory and applies the got memory  on the VMware Environment 
      6 - Disable Time Synchronization
      7 - Added the virtual machine at the Cluster 

.NOTES  
    File Name  : VM_ConvertToStandard.ps1   
    Requires   : PowerShell V4 , Powercli 5.5 
    

.WARNING 

    This script should not be launched with ISE, because  otherwise there are a conflict between the cmdlet PowerCli and PowerShell


.INFORMATION 

      - The OutViewNoStandardVHDX Function is the entered Point of the Script , this is function that define the Virtual Machines not Standard ,and built
         the list on Virtual Machines.
     
      - The Script check if the virtual Machine is Off
         
      - The cmdlet "Move-VMStorage" does not delete the source folder after Moving, thus the script delete the source folder if it is empty.
        
      - The script check if PowerCli is installed , if it's not the case the step 4 is skiped.

      - This step is skipped , if the virtual machine is not find in the vCenter

      -  If no Cluster is discovered, the environment is considered as a Micro Site, and the Virtual Machine is not added to Cluster.

.UPDATE 
 Add line : Move-VMStorage -ea SilentlyContinue -VMname $VMName -ComputerName $Node –VirtualMachinePath $Destination
#>
#****************************************************************************************************

[CmdletBinding()]
Param(
  [parameter(Mandatory=$false)]
  [string]
   $Version 

)




$versionScript = "1.1.2"
$Realse = "0923106"
$ScriptDir = split-path -parent $MyInvocation.MyCommand.Path

if($Version)
{
    return $versionScript
}



#********************************************************
function LogDisplay
{
  param([string]$Message="", [string]$TypeMess="")
  If ($TypeMess -eq "Error") {Write-Host $Message -foregroundcolor Red}
  If ($TypeMess -eq "Information") {Write-Host $Message -foregroundcolor Yellow}
  If ($TypeMess -eq "Success") {Write-Host $Message -foregroundcolor Green}
  
}

#********************************************************
Function ShowMessage {
    param($Message, $Title , $Btn, $Icon)
    #http://cnf1g.com/?p=279

	[System.Windows.Forms.MessageBox]::Show($Message, $Title , $Btn, $Icon)
}   #Example : ShowMessage -Message "The Virtual Machine must be Off" -Title "ConvertVMtoStandard" -Btn 0 -Icon 48


#********************************************************
function Load-ListBox 
{
<#
	.SYNOPSIS
		This functions helps you load items into a ListBox or CheckedListBox.

	.DESCRIPTION
		Use this function to dynamically load items into the ListBox control.

	.PARAMETER  ListBox
		The ListBox control you want to add items to.

	.PARAMETER  Items
		The object or objects you wish to load into the ListBox's Items collection.

	.PARAMETER  DisplayMember
		Indicates the property to display for the items in this control.
	
	.PARAMETER  Append
		Adds the item(s) to the ListBox without clearing the Items collection.
	
	.EXAMPLE
		Load-ListBox $ListBox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Load-ListBox $listBox1 "Red" -Append
		Load-ListBox $listBox1 "White" -Append
		Load-ListBox $listBox1 "Blue" -Append
	
	.EXAMPLE
		Load-ListBox $listBox1 (Get-Process) "ProcessName"
#>
	Param (
		[ValidateNotNull()]
		[Parameter(Mandatory=$true)]
		[System.Windows.Forms.ListBox]$ListBox,
		[ValidateNotNull()]
		[Parameter(Mandatory=$true)]
		$Items,
	    [Parameter(Mandatory=$false)]
		[string]$DisplayMember,
		[switch]$Append
	)
	
	if(-not $Append)
	{
		$listBox.Items.Clear()	
	}
	
	if($Items -is [System.Windows.Forms.ListBox+ObjectCollection])
	{
		$listBox.Items.AddRange($Items)
	}
	elseif ($Items -is [Array])
	{
		$listBox.BeginUpdate()
		foreach($obj in $Items)
		{
			$listBox.Items.Add($obj)
		}
		$listBox.EndUpdate()
	}
	else
	{
		$listBox.Items.Add($Items)	
	}

	$listBox.DisplayMember = $DisplayMember	
}
#endregion


#********************************************************
Function ResetForm {

$listBox.Items.Clear()
$progress.Value = 0
$VMName = ""
$Path = ""
$Node = ""
$textbox_node.Text = "Please type Node name"

}

#********************************************************
Function ProgressBar {

        param($step)

        $MaxStep = 11

        if($checkbox1.Checked)
        {
            $MaxStep = 12    
        }
        if($checkbox2.Checked)
        {
            $MaxStep = 12    
        }
        if($checkbox1.Checked -AND $checkbox2.Checked )
        {
            $MaxStep = 14    
        }

        $Value = ( $Step * 100 / $MaxStep)
 
        Return $Value

}


#********************************************************
Function CheckStatusVM
{
        param($Node,$VMName)
        
        $VMStatus =  Get-VM -ComputerName $Node | Where Name -Like $VMName
        
        if($VMStatus.State -ne "off")
	    {
        
		    #Write-Host "The Virtual Machine must be Off" `n -ForegroundColor Red
            ShowMessage -Message "The Virtual Machine must be Off ! " -Title "Convert VM toStandard" -Btn 0 -Icon 48
        
		    ResetForm
	    }
}

#********************************************************
#--------------------------------------------
# Define a not Standard Virtual Machines
#--------------------------------------------
Function OutViewNoStandardVHDX {
        param($Node="")
    
        # This line is the Secret, it defines the Virtual Machine displayed in the listbo
        $GetVMHardDisk = @(Get-VMHardDiskDrive * -ComputerName $Node |  Where-Object { $_.Path -notLike "*"+$_.VMName+"\Virtual Hard Disks\"+$_.VMName+"*.vhdx"} | Where-Object { $_.Path -notLike "*.avhdx"}) + (Get-VMHardDiskDrive * -ComputerName $Node|  Where-Object {$_.Path -notlike "C:\ClusterStorage\*"}) 
        $VM = @()
        $GetVMHardDisk | Foreach-Object{

            $MyDetails = "" | select-Object VMName, ControllerType, ControllerNumber, ControllerLocation, Path, DiskNumber

            $isExist = $_.VMname
            $CheckDuplication = ($VM | ? { $_.VMname -like $isExist }).Count
        
            if($CheckDuplication -eq 0 )
            {

                $MyDetails.VMName = $_.VMname

                $MyDetails.ControllerType = $_.ControllerType

                $MyDetails.ControllerNumber = $_.ControllerNumber

                $MyDetails.ControllerLocation = $_.ControllerLocation

                $MyDetails.Path = $_.Path

                $MyDetails.DiskNumber = $_.DiskNumber
         
     
                $VM += $MYDetails
            }

        } 
    
        Return $VM
        #If you want out-view = $VM | Out-GridView  -OutputMode Single -Title "Select one or more VMs to Live Migrate"
}


#********************************************************
#--------------------------------------------
# List the Datastores discovered
#--------------------------------------------
Function OutViewVolumes {  
               
                param($Node="")

                $myCol = @()
                ForEach( $Volume in (Invoke-Command  -ComputerName $Node -ScriptBlock  {Get-ClusterSharedVolume})){


                    $VolumeInfo = "" | select-Object Name, FreeSpaceGB, UsedSpaceGB, SizeGB, PercentFree, FileSystem  
                    $VolumeInfo.Name = $Volume.Name
                    
                    ForEach ($a in (Invoke-Command  -ComputerName $Node -ScriptBlock {Get-ClusterSharedVolume $args[0] | select -Expand SharedVolumeInfo | select -ExpandProperty Partition } -ArgumentList (,$Volume.Name)))
                    {
                        
   
                       
                       $VolumeInfo.FreeSpaceGB = [math]::round($a.FreeSpace/1GB,2)
                       $VolumeInfo.UsedSpaceGB = [math]::round($a.UsedSpace/1GB,2)
                       $VolumeInfo.SizeGB = [math]::round($a.Size/1GB,2)
                       $VolumeInfo.PercentFree = $a.PercentFree 
                       $VolumeInfo.FileSystem = $a.FileSystem
   
                       
                    }
                    $myCol += $VolumeInfo 
                    
                }
                 
               
                 $Volume =  $myCol | Out-GridView -OutputMode Single
                 Return $Volume
}

#********************************************************
Function RenameVHDX {
        param([string]$VMName,[string]$Node)

        $label2.Text = "Rename VHDX in progress :"
        $VMHardDisk = Get-VMHardDiskDrive -VMName $VMName -ComputerName $Node
        $Value = ProgressBar 5 
        $progress.Value = $Value 
        "Disable Time Synchronization..." | Out-Host
        $DisableTimeSynchronization = Disable-VMIntegrationService -name "Time Synchronization" -vmname $VMName -computername $Node
        $ihd = 0
    
	    $VMHardDisk | ForEach-Object{

                $Path = $_.Path
                $VHDName = $Path.Substring($Path.LastIndexOf("\")+1) 
                #NetworkPath
                $NetworkPathVHDX = $Path.replace("C:","c$") 
                $NetworkPathVHDX = "\\$Node\$NetworkPathVHDX"

                if($VHDName -notlike ("$VMname*"+".vhdx"))
                {
        
                     #DiskNumber
                     if( $ihd -eq 0 ){
           
                         $NewVHDname = "$VMName"+".vhdx"
                         $NewPath = ($Path.replace("C:\","C:\\")).replace($VHDName,$NewVHDname)
                 
                 
                         Rename-item -path "$NetworkPathVHDX" -newname $NewVHDname
                         $Value = ProgressBar 6 
                         $progress.Value = $Value
                         Set-VMHardDiskDrive -VMName $VMName -Computername $Node –ControllerType $_.ControllerType -ControllerNumber $_.ControllerNumber -ControllerLocation $_.ControllerLocation -Path $NewPath 
		                 $Value = ProgressBar 7 
                         $progress.Value = $Value 
                 

                    }else{
				 
				         $DiskNumber = $ihd.ToString("00")
				         $NewVHDname = "$VMName"+ "_" + $DiskNumber + ".vhdx"
                         $NewPath = ($Path.replace("C:\","C:\\")).replace($VHDName,$NewVHDname)
                 
                         $error.Clear()
                         Rename-item -path $NetworkPathVHDX -newname $NewVHDname
                         $Value = ProgressBar 8
                         $progress.Value = $Value
                 
                         if(!$?)
                         {

                            ShowMessage -Message ("Rename VHDName - Error "+ $error[0]) -Title "ConvertVMtoStandard" -Btn 0 -Icon 16
                    
                         }
                
                         Set-VMHardDiskDrive -VMName $VMName -Computername $Node –ControllerType $_.ControllerType -ControllerNumber $_.ControllerNumber -ControllerLocation $_.ControllerLocation -Path $NewPath 
                         $Value = ProgressBar 9
                         $progress.Value = $Value

                         $error.Clear()
                         if(!$?)
                         {
                             ShowMessage -Message ("Change Path Folder Path - Error "+$error[0]) -Title "ConvertVMtoStandard" -Btn 0 -Icon 16
                    
                         }else
                         {

                            "Change Path Folder Path - Success... " | Out-Host
                            #ShowMessage -Message ("Change Path Folder Path - Success ") -Title "ConvertVMtoStandard" -Btn 0 -Icon 64
                            $Value = ProgressBar 10
                            $progress.Value = $Value
                         }
                
                    }

                        $ihd = $ihd + 1

                }else
                {
                    Write-Host "Already Renammed"
                    $ihd = $ihd + 1
                }

            }
	
	        $CheckConf = @(Get-VMHardDiskDrive $VMname -ComputerName $Node |  Where-Object { $_.Path -notLike "*"+$_.VMName+"\Virtual Hard Disks\"+$_.VMName+"*.vhdx"} | Where-Object { $_.Path -notLike "*.avhdx"}) + (Get-VMHardDiskDrive $VMname -ComputerName $Node |  Where-Object {$_.Path -notlike "C:\ClusterStorage\*"}).Count
   
            if($CheckConf -like 0)
            {  
               $Value = ProgressBar 11 
               $progress.Value = $Value
               $label2.Text = "Rename VHDX Success"
      
            }
   
}

#********************************************************
function Move-VMStorage2
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # A name of a VM or a VM object
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $VMName,
        # The name of the Hyper-V host
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Node,
        # The path where the VM is going to be relocated to.
        [string]
        $Path 
    )
        
        
        $label2.Text = "Move Virtual Machine in progress :"
        # Lets move and tidy the source folder
        $VMName = Get-VM $VMName -ComputerName $Node
        "Move Storage...In progress"
        
        if( $Path -notcontains $VMName.Name)
        {

             $error.clear()
             $Value = ProgressBar 1
             $progress.Value = $Value

             $Path = $Path.Substring(0,$Path.Length-1)
             $VM =$VMName.Name
             $Command = "Move-VMStorage -Computername $Node -VMname $VM -DestinationStoragePath $Path"
             Invoke-Expression $command

             
             If (!$?){

                ShowMessage -Message ("Move Storage to $Path - Error "+ $error[0]) -Title "ConvertVMtoStandard" -Btn 0 -Icon 16
                
                ResetForm
                    

             }else
             {
                
                "Move Storage to $Path - Success" | Out-Host
                #ShowMessage -Message ("Move Storage to $Path - Success ") -Title "ConvertVMtoStandard" -Btn 0 -Icon 64
                $Value = ProgressBar 2
                $progress.Value = $Value
             }
            
            
             $VMOldPath = $VMName.Path           
             $FolderEmpty = 1
             $FolderEmpty = Invoke-Command -ComputerName $Node -ScriptBlock { @(Get-ChildItem $args[0] -Recurse | where {!$_.PsIsContainer}).Count } -ArgumentList ($VMOldPath)
              
             if($FolderEmpty -eq 0 -AND $FolderEmpty -ne $null)
             {
                Invoke-Command -ComputerName $Node {Remove-Item -Path $args[0] -Recurse} -ArgumentList ($VMOldPath)
             }


             If(!$?)
             {
                ShowMessage -Message ("VM Source Folder not removed") -Title "ConvertVMtoStandard" -Btn 0 -Icon 16
             }
                              
             $Value = ProgressBar 4
             $progress.Value = $Value
                  
        }
}       


#********************************************************
function Move-VMStorage1
{
param([string]$VMName,[string]$Node)

   
    $VMHardDisk = Get-VMHardDiskDrive -VMName $VMName -ComputerName $Node

    $VMHardDisk | Foreach-Object{


                $Path = $_.Path 
                $VHDName = $Path.Substring($Path.LastIndexOf("\")+1)

                #Current location network 
                $LocationNetwork = $Path.replace("C:","c$")
                $LocationNetwork = "\\$Node\$LocationNetwork"
                #New location network 
                $NewLocation = $Path.Substring(0,$Path.LastIndexOf("\")+1)  + "Virtual Hard Disks\"
                $NewLocationNetwork = $NewLocation.replace("C:","c$")
                $NewLocationNetwork = "\\$Node\$NewLocationNetwork"
 
                $CheckFolder = test-path $NewLocationNetwork
 
                 if($CheckFolder -eq 0)
                 {

                    $Folder = New-Item -path "$NewLocationNetwork" -ItemType directory

                 }

    
                 Move-Item -Path $LocationNetwork -Destination $NewLocationNetwork
         
                 $NewPath = $NewLocation.replace("C:\","C:\\") +"$VHDName"
 
                 Set-VMHardDiskDrive -VMName $VMName -Computername $Node –ControllerType $_.ControllerType -ControllerNumber $_.ControllerNumber -ControllerLocation $_.ControllerLocation -Path $NewPath

         
                 $Destination = $Path.Substring(0,$Path.LastIndexOf("\")+1).replace("Virtual Hard Disks\","")        
    }
     
    #Alow to move pagefile SmartPagingFile SnapShotFilePath VirtualMachinePath
    $MoveVirtualMachineFiles = "Move-VMStorage -ea SilentlyContinue -VMname $VMName -ComputerName $Node -SmartPagingFilepath $Destination -SnapShotFilePath $Destination –VirtualMachinePath $Destination"
    Invoke-Expression $MoveVirtualMachineFiles
}


#--------------------------------------------------------
# Get-Virtual Machien Memory on the VMware Infrastructure
#--------------------------------------------------------
#********************************************************
function GetMemory
{

    param($VMName)
    Add-PSSnapin VMware.VimAutomation.Core -ea SilentlyContinue

    #find the itc code
    try{
        $itc = (((gpresult /r /scope computer | ? {$_ -like ‘*=*’}) -split ‘,’) | ? {$_ -like 'OU=??'}).replace("OU=","")
    }catch{
    
        $itc = (((gpresult /r /scope computer | ? {$_ -like ‘*=*’}) -split ‘,’) | ? {$_ -like 'OU=???'}).replace("OU=","")
    }
    
    $itc= $itc.ToLower()
    $ITC_VC="$itc-vc"

    $ConnectVcenter = Connect-VIServer $ITC_VC
	"Connect-VIServer... " +  $ITC_VC | Out-Host
    Write-Host "Connect to $ITC_VC"
    Write-Host "Please wait ..." -ForegroundColor Yellow 

    if(@(Get-VM $VMName -ea SilentlyContinue) -ne 0){
        
        $VMwareVM = Get-VM $VMName
        $global:MemorySize = $VMwareVM.MemoryGB

    }else
    {

        ShowMessage -Message ( "Virtual Machine $VMName not found on the vCenter $ITC_VC") -Title "ConvertVMtoStandard" -Btn 0 -Icon 16

    }
    
    #Disconnect virtual machine connection
    Disconnect-VIServer -server $ITC_VC -Confirm:$false 
	"Disconnect to... " + $ITC_VC | Out-Host
   
}

#********************************************************
Function ConfigureMemorySize
{

param($VMName,$Node)

        
        $MemorySize = $global:MemorySize*1024
        $MemorySize = [string]$MemorySize+"MB" 
    
        try
        {
        
            $command = "Set-VMMemory $VMName -ComputerName $Node -DynamicMemoryEnabled `$false -StartupBytes $MemorySize" 
            Invoke-Expression $command 
      
            Write-Host "Success - Dynamic Memory disable "`n  

        }
        catch
        {
                    
            Write-Host "Error to disable dynamic memory "`n -ForegroundColor Red
        }
}

Function DisableDynamicMemory
{
        param($VMName,$Node)
  
        try
        {
            $command = "Set-VMMemory $VMName -ComputerName $Node -DynamicMemoryEnabled `$false" 
            Invoke-Expression $command
         
            Write-Host "Success - Dynamic Memory disable "`n 
        }catch
        {

           Write-Host "Error to disable dynamic memory "`n -ForegroundColor Red
        }
    
}


Function AddToCluster
{
        param($Node,$VMname)

        $ClusterExist = Invoke-Command -ComputerName $Node -ScriptBlock { @(Get-Cluster -ea Ignore).Count }

        if($ClusterExist -ne 0)
        {
      
            try
            {
        
                Invoke-Command -ComputerName $Node -ScriptBlock { Get-VM $args[0] | Add-ClusterVirtualMachineRole} -ArgumentList ($VMname)
            
                #LogDisplay "$VMname added to Cluster" "Success" `n

            }
            catch
            {
        
                #LogDisplay "$VMname not added to Cluster" "Error" `n
                ShowMessage -Message ( "$VMname not added to Cluster.") -Title "ConvertVMtoStandard" -Btn 0 -Icon 16
            }
        }
        else
        {

         ShowMessage -Message ( "Micro Site detected, you can't add node to Cluster, if it's not a Micro Site please do it manually.") -Title "ConvertVMtoStandard" -Btn 0 -Icon 48
         #LogDisplay "Cluster not detected, if a cluster exist please add the node to cluster manually" "Information" `n 

        }
}


#region
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
#endregion            

#################################################
# CONFIGURE THE WINDOWS FORM
#################################################

# Create form 
$form = New-Object Windows.Forms.Form

# Block Resize
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $False
$form.MinimizeBox = $False

# Define title
$form.Text = "Convert To Standard VM"

# Define Size form 
$form.Size = New-Object System.Drawing.Size(400,470)

#################################################
# ADDING COMPONENTS
#################################################

# Button OK
$button_ok = New-Object System.Windows.Forms.Button
$button_ok.Text = "OK"
$button_ok.Size = New-Object System.Drawing.Size(355,40)
$button_ok.Location = New-Object System.Drawing.Size(20,340)

# Button Connect
$button_Connect = New-Object System.Windows.Forms.Button
$button_Connect.Text = "Connect"
$button_Connect.Size = New-Object System.Drawing.Size(90,25)
$button_Connect.Location = New-Object System.Drawing.Size(180,9)

# Button Exit
$button_exit = New-Object System.Windows.Forms.Button
$button_exit.Text = "Cancel"
$button_exit.Size = New-Object System.Drawing.Size(355,40)
$button_exit.Location = New-Object System.Drawing.Size(20,390)



# Label 1
$label_prez = New-Object System.Windows.Forms.Label
$label_prez.AutoSize = $true
$label_prez.Location = New-Object System.Drawing.Point(20,45)
$label_prez.Size = New-Object System.Drawing.Size(100,20)
$label_prez.Text = "Please select Virtual Machine :"



# CheckBox 1
$checkbox1 = New-Object System.Windows.Forms.CheckBox
$checkbox1.AutoSize = $true
$checkbox1.Location = New-Object System.Drawing.Point(20,210)
$checkbox1.Name = 'checkbox1'
$checkbox1.Size = New-Object System.Drawing.Size(80,20)
$checkbox1.Text = 'Update Memory with vCenter Server'
$checkbox1.Checked = $true

# CheckBox 2
$checkbox2 = New-Object System.Windows.Forms.CheckBox
$checkbox2.AutoSize = $true
$checkbox2.Location = New-Object System.Drawing.Point(20,230)
$checkbox2.Name = 'checkbox2'
$checkbox2.Size = New-Object System.Drawing.Size(80,20)
$checkbox2.Text = 'Add VM to Cluster'
$checkbox2.Checked = $true


# CheckBox 3
$checkbox3 = New-Object System.Windows.Forms.CheckBox
$checkbox3.AutoSize = $true
$checkbox3.Location = New-Object System.Drawing.Point(20,250)
$checkbox3.Name = 'checkbox3'
$checkbox3.Size = New-Object System.Drawing.Size(80,20)
$checkbox3.Text = 'Change Virtual Machine Location (Slow)'
$checkbox3.Checked = $false

# Label 2
$label1 = New-Object System.Windows.Forms.Label
$label1.AutoSize = $true
$label1.Location = New-Object System.Drawing.Point(20,250)
$label1.Name = 'label_password'
$label1.Size = New-Object System.Drawing.Size(100,20)
$label1.Text = "a simple Label"

# TextBox Node
$textbox_node = New-Object System.Windows.Forms.TextBox
$textbox_node.AutoSize = $true
$textbox_node.Location = New-Object System.Drawing.Point(20,12)
$textbox_node.Name = 'textbox_sw'
$textbox_node.Size = New-Object System.Drawing.Size(140,40)
$Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Italic)
$textbox_node.Font = $Font
$textbox_node.Text = "Please type Node name"


# Label 3 - Text - Progress Bar
$label2 = New-Object System.Windows.Forms.Label
$label2.AutoSize = $true
$label2.Location = New-Object System.Drawing.Point(20,280)
$label2.Name = 'label_complex'
$label2.Size = New-Object System.Drawing.Size(100,20)
$label2.Text = "Progress bar :"

# Label 4 - Script Version
$label_version = New-Object System.Windows.Forms.Label
$label_version.AutoSize = $true
$label_version.Location = New-Object System.Drawing.Point(355,10)
$label_version.Name = 'version'
$label_version.Size = New-Object System.Drawing.Size(100,20)
$label_version.Text = "$versionScript"

# Progress Bar
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20,300)
$progress.Name = 'progressBar'
$progress.Size = New-Object System.Drawing.Size(350,23)

# ListBox

$listBox2 = New-Object System.Windows.Forms.ListBox 
$listBox2.Location = New-Object System.Drawing.Size(10,130)  
$listBox2.Size = New-Object System.Drawing.Size(345,20) 
$listBox2.Height = 80
$listBox2.Text = "Please select datastore : "
$Datastore= $listBox2.SelectedItem

# ListBox

$listBox = New-Object System.Windows.Forms.ListBox 
$listBox.Location = New-Object System.Drawing.Size(20,70)
$listBox.Size = New-Object System.Drawing.Size(345,20) 
$listBox.Height = 120
$listBox.Text = "Please select  Virtual Machine : "

#################################################
# MAIN - MANAGE EVENTS
#################################################

$button_Connect.Add_Click(
{


#----------------------------------------------
#region Import Hyper-V Module
#----------------------------------------------

Import-Module Hyper-V
Import-Module failoverclusters
#endregion Import Hyper-V Module

Start-Transcript -path $ScriptDir\logs.txt  -append


$Node = [string]$textbox_node.text 

$label2.Text = "Please wait connection in progress ..."
$ping = new-object System.Net.NetworkInformation.Ping 
try{
   $ReponsePing = $ping.Send("$Node") 
}catch{}

    if($node -like "????-s*")
    {

        if ($ReponsePing.status –eq “Success”)
        {
            #List Box 1
            $VolumePathNotStandard = OutViewNoStandardVHDX -Node "$Node"
            $Box2 = Load-ListBox $listBox ($VolumePathNotStandard) "VMName"

            $label2.Text = "Connected"
            $VMname = $listBox.SelectedItem

            #Disable CheckBox
            $ClusterExist = @(Get-Cluster -Name $Node -ErrorAction Ignore | Get-ClusterNode).count
            if($ClusterExist -eq 0){
                $Checkbox2.Checked = $false
                $Checkbox2.Enabled=$False
            }
   
 
        }else
        {

            $label2.Text = "Error Connection ..."
            ShowMessage -Message "Cannot contact node !" -Title "ConvertVMtoStandard" -Btn 0 -Icon 16 

        }
    }else{

        $label2.Text = "Please specify the NodeName not the ClusterName ..."

    }

})


# Manage event when we click to the Close Button
$button_exit.Add_Click(
{
$form.Close();

})

# Manage event when we click to the Ok Button
$button_ok.Add_Click(
{
    Stop-Transcript
    $label2.Text = "Tasks in Progress ... :" 
    
    #Get VM Selected 
    $VMName = ($listBox.SelectedItem).VMName 


    $LogPath = "$ScriptDir\$VMName`_logging.txt"

    Start-Transcript -path $LogPath  -append    
   
    $global:Node = [string]$textbox_node.text
    

    $global:StartTime = Get-Date

    "Start Time..." + " $global:StartTime" | Out-Host

    
    #Check if the VM is Off 
    CheckStatusVM -VMName $VMname -Node $Node
     
    if($Checkbox3.Checked)
    {
            $Volume = OutViewVolumes $Node 
    }
    
    $Path = (Get-VMHardDiskDrive $VMName -ComputerName $Node).Path

    if($Path[0] -notlike  "*"+($Volume.Name)+"*")
    {
        #$Volume = OutViewVolumes $Node 

        if($Volume.Name -eq $null)
        {
            ShowMessage -Message "DataStore not Selected" -Title "ConvertVMtoStandard" -Btn 0 -Icon 16 
            $form.Close(); 
        }
        $VMname = $VMname.ToUpper()
        $Path = "C:\ClusterStorage\"+$Volume.Name+"\$VMname\" 

        $global:MoveStorageStartTime = Get-Date
        "Move Storage Start Time..." + " $global:MoveStorageStartTime" | Out-Host
		"Move Storage to ... " + $Path | Out-Host
        Move-VMStorage2 -Node $Node -VMName $VMName -Path $Path 
        $global:MoveStorageEndTime = Get-Date
        "Move Storage End Time..." + " $global:MoveStorageEndTime" | Out-Host    
               
    }
    else
    { 
        if($Path -notlike ("C:\ClusterStorage\*"+"\$VMname*"+"\Virtual Hard Disks\*"))
        {
                $Path = (Get-VMHardDiskDrive $VMName -ComputerName $Node).Path
                $Path = $Path.split("\")[2]
                if($Path -like ("????"+"-D"+"?????"+"-LUN*")) 
                {   
                    
                    $Path = "C:\ClusterStorage\"+$Path+"\$VMname\"
					"Move Storage to ... " + $Path | Out-Host
					$global:MoveStorageStartTime = Get-Date
					"Move Storage Start Time..." + " $global:MoveStorageStartTime" | Out-Host
                     Move-VMStorage1 -Node $Node -VMName $VMName -Path $Path
					$global:MoveStorageEndTime = Get-Date
					"Move Storage End Time..." + " $global:MoveStorageEndTime" | Out-Host
                }
                else
                {

                    $Volume = OutViewVolumes $Node

                    if($Volume.Name -eq $null){

                        ShowMessage -Message "DataStore not Selected" -Title "ConvertVMtoStandard" -Btn 0 -Icon 16
                        $form.Close();
                    }
                    $VMname = $VMname.ToUpper()
                    $Path = "C:\ClusterStorage\"+$Volume.Name+"\$VMname\"
                    "Move Storage to ... " + $Path | Out-Host
                    $global:MoveStorageStartTime = Get-Date
			
                    $VMHardDisk = Get-VMHardDiskDrive -VMName $VMname -ComputerName $node

                    $SamePath = $VMHardDisk | ? { $_.Path -like $Path+"*"}

                    "Move Storage Start Time..." + " $global:MoveStorageStartTime" | Out-Host
                     
                     $global:global:MoveStorageEndTime = Get-Date
                    "Move Storage End Time..." + " $global:MoveStorageEndTime" | Out-Host

                     if($SamePath.Count -eq 0 )
                     {
                        Move-VMStorage2 -Node $Node -VMName $VMName -Path $Path

                      }else
                     {   
                        Move-VMStorage1
                     }

                    
                    


                }
         }else{ Write-Host "Already Done"} 
    }
   
    "Rename Virtual vhdx ..." | Out-Host
    #RenameVHDX Function 
    RenameVHDX -VMName $VMName -Node $Node
	
	$VMclustered  = (Get-VM $VMName -ComputerName $Node).isClustered
    
    if($checkbox1.Checked)
    {
       $Powercli = @(Get-WmiObject -Class Win32_Product | ? {$_.Name -like "VMware vSphere PowerCLI" }).Count
       
	   
       if( $Powercli -eq 1)
       {
		        
            $DynamicMemoryEnabled = (Get-VMMemory $VMName -ComputerName $Node).DynamicMemoryEnabled
            
            "PowerCLI found ..." | Out-Host
            
            if($DynamicMemoryEnabled -like "True")
            {
                #$MemorySize = GetMemory -VMname $VMName
                $label2.Text = "Get Virtual Machine Memory in vCenter in progress ..."
				"$VMname - Get Virtual Machine Memory in vCenter in progress ..." | Out-Host
                GetMemory -VMname $VMName
                #GetMemory -VMname $VMName
				"Disble Dynamic Memory and fix Startup Memory ..." | Out-Host
                $label2.Text = "Disbled Dynamic Memory and fix Startup Memory ..."
				"Configure Memory Size ..." | Out-Host
                ConfigureMemorySize -VMName $VMName -Node $Node 

            }else{ShowMessage -Message ("Dynamic memory is already disabled, if you want change the memory size please do manually") -Title "ConvertVMtoStandard" -Btn 0 -Icon 48}
        }
        else{
          
          ShowMessage -Message ("VMware vSphere PowerCLI not installed, the virtual memory can't be updated") -Title "ConvertVMtoStandard" -Btn 0 -Icon 48

        }

    }else{

        $DynamicMemoryEnabled = (Get-VMMemory $VMName -ComputerName $Node).DynamicMemoryEnabled
        
        if($DynamicMemoryEnabled -like "True"){
            DisableDynamicMemory -VMName $VMName -Node $Node
        }

    }

   
    $label2.Text = "Add VM to Cluster ..."
    if($checkbox2.Checked -AND $VMclustered -ne "True"){
        # Comment to removed
		"Add VM to Cluster ..." | Out-Host
         AddToCluster -Node $Node -VMname $VMName
    }

    #$Value = ProgressBar 14
    $progress.Value = 100 
    $label2.Text = "Success Convertion"
    #ShowMessage -Message ("Success") -Title "ConvertVMtoStandard" -Btn 0 -Icon 64

    sleep 5
    ResetForm

    $EndTime = get-date 

    "************************EXECUTION TIME SCRIPT*******************************" | Out-Host
    try{
        NEW-TIMESPAN –Start $StartTime –End $EndTime | Out-Host
    }catch{}
    "************************EXECUTION TIME MOVE STORAGE************************" | Out-Host

    "Start to ... " + $global:MoveStorageStartTime + " End to ... " +  $global:MoveStorageEndTime | Out-Host
    try{
    NEW-TIMESPAN –Start $global:MoveStorageStartTime –End  $global:MoveStorageEndTime | Out-Host
	}catch{}
	 "************************VHDX**********************************************" | Out-Host


     $vmdisk= Get-VMHardDiskDrive $VMName -ComputerName $Node
     try{
        $Size = Get-VHD -ComputerName $Node -VMId $vmdisk.VMId 
     }catch{}
 

     $vmdisk |ForEach-Object {

        $Path = $_.Path

  
        $VHDsize = ($Size | ? { $_.Path -like $Path } | select -first 1 ).Size

        Write-Host ($VHDsize/1024/1024/1024) "GB" "|" $Path | Out-Host

    }
 
    Stop-Transcript
  
})



#################################################
# Insert Compenent
#################################################

# Add compenent to the form
$form.Controls.Add($label_prez)
$form.Controls.Add($label_version)
$Form.Controls.Add($listBox)
#$Form.Controls.Add($listBox2)
$form.Controls.Add($DisplayMember)
$form.Controls.Add($checkbox1)
$form.Controls.Add($checkbox2)
$form.Controls.Add($checkbox3)
#$form.Controls.Add($label1)
$form.Controls.Add($label2)
$form.Controls.Add($textbox_node)
$form.Controls.Add($progress)
$form.Controls.Add($button_ok)
$form.Controls.Add($button_Connect)
$form.Controls.Add($button_exit)

# Display Windows 
$form.ShowDialog()


#################################################
# END OF PROGRAM
#################################################

