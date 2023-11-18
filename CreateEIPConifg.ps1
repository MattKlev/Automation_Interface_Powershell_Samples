  $source = @" 
  namespace EnvDteUtils
  {
    using System; 
    using System.Runtime.InteropServices; 
    
    public class MessageFilter : IOleMessageFilter 
    { 
      public void Register() 
      { 
        IOleMessageFilter newFilter = new MessageFilter();  
        IOleMessageFilter oldFilter = null;  
        CoRegisterMessageFilter(newFilter, out oldFilter); 
      } 

      public void Revoke() 
      { 
        IOleMessageFilter oldFilter = null;  
        CoRegisterMessageFilter(null, out oldFilter); 
      } 
      
      int IOleMessageFilter.HandleInComingCall(int dwCallType, System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr lpInterfaceInfo)
      { 
        return 0; 
      } 

      int IOleMessageFilter.RetryRejectedCall(System.IntPtr hTaskCallee, int dwTickCount, int dwRejectType) 
      { 
        if (dwRejectType == 2) 
        { 
          return 99; 
        } 
        return -1; 
      } 

      int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee, int dwTickCount, int dwPendingType) 
      { 
        return 2;  
      } 

      [DllImport("Ole32.dll")] 
      private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter); 
    }  

    [ComImport(), Guid("00000016-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)] 
    interface IOleMessageFilter  
    { 
      [PreserveSig] 
      int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

      [PreserveSig]
      int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

      [PreserveSig]
      int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }
  }
"@
  Add-Type -TypeDefinition $source
$MessageFilter=New-Object -TypeName EnvDteUtils.MessageFilter
$MessageFilter.Register()

<#Only change these parameters#>
$ProjectName = "EIP_Test"
$targetDir = "E:\tmp"

#name of the Xti file
 $XtiFileName = "Box 1 (ConveyLinxLogix)"

#This number of new boxes will be created with an IP address starting at 1
 $NumOfNewBox = 120 
<##############################>

# Determine script location for PowerShell
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path


<#Delete the old project folder and crate a new one#>
$ProjectFolder =$targetDir+'\'+$ProjectName

if(test-path $ProjectFolder -pathtype container)
{
    Remove-Item $ProjectFolder -Recurse -Force
}

New-Item $ProjectFolder –type directory
<#################################>


<#Create a folder for the TwinCAT project#>
$TwinCATProjectFolder = $ProjectFolder+'\'+$ProjectName

New-Item $TwinCATProjectFolder –type directory

<#########################################>

<# Create new XTI files based on the template#>
    
    #location of the template xti file
    $XtiTemplate = ($ScriptDir+'\'+$XtiFileName +".xti")
    
    #Storage location for the new XTI files
    $ModifedXtiFolderName =$ProjectFolder + "\XTI_Files"
    
    New-Item $ModifedXtiFolderName –type directory
    
    #Load the template XTI file and select the "NewSlavePara" node
    $doc = New-Object System.Xml.XmlDocument
    $doc.Load($XtiTemplate)
    
    $EthernetIP = $doc.SelectSingleNode("//EthernetIp")
    $NewSlavePara = $EthernetIP.NewSlavePara
    
    $NewSlaveParaCharArray = $NewSlavePara.ToCharArray()
   
    $XtiFilePaths = New-Object System.Collections.ArrayList
    #Modify the template XTI file - and save a new copy
    For ($i=1; $i -le $NumOfNewBox; $i++) 
    {
           $Chars = ("{0:X}" -f $i).ToCharArray()
   
            

        
            if ($Chars.Length -ge 2)
            {
                $NewSlaveParaCharArray[265] = $Chars[1]
                $NewSlaveParaCharArray[264] = $Chars[0]
            }
            elseIf($Chars.Length -eq 1)
            {
                $NewSlaveParaCharArray[265] = $Chars[0]  
            }

            $NewSlavePara = -join $NewSlaveParaCharArray

            $EthernetIP.NewSlavePara = $NewSlavePara

            $XtiFilePaths.Add($ModifedXtiFolderName +"\"+ $XtiFileName + "_" + $i + ".xti")
            $doc.Save($XtiFilePaths[$i-1])
    }
    
    ###########


<#Create an instance of Visual studio#>
$dte = new-object -com VisualStudio.DTE.14.0
$dte.SuppressUI = $true
$dte.MainWindow.Visible = $false
<#####################################>


<#Create a new TwinCAT project from the template#>
$template = "C:\TwinCAT\3.1\Components\Base\PrjTemplate\TwinCAT Project.tsproj"
$sln = $dte.Solution
$project = $sln.AddFromTemplate($template,$TwinCATProjectFolder,$ProjectName +'.tsp')
<################################################>

<#Get the targeted AMS Net if of the TwinCAT project#>
$systemManager = $project.Object
$targetNetId = $systemManager.GetTargetNetId()
write-host $targetNetId
<####################################################>


<#Save#>
$project.Save();
$sln.SaveAs($TwinCATProjectFolder)
<######>

<#Create an EIP Master#>
$devices = $systemManager.LookupTreeItem("TIID")
$EIPMaster = $devices.CreateChild("EIP Master", 133, $null, $null)
<#####################>

<#adds slaves to the EIP Scanner#>
For ($i=1; $i -le $NumOfNewBox; $i++) 
    {
     Write-Host "Adding: " $XtiFilePaths[$i-1]
     $NewChild= $EIPMaster.ImportChild($XtiFilePaths[$i-1],$null,$null,$null)
    }

<#Save#>
$project.Save();
$sln.SaveAs($TwinCATProjectFolder)
<######>

<#Activate the TwinCAT configuration#>
#$systemManager.ActivateConfiguration()
#Start-Sleep 5
<###################################>

<#Start TwinCAT in run mode#>
#$systemManager.StartRestartTwinCAT()
<###########################>

$MessageFilter.Revoke()