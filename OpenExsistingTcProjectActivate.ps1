#$netId=$args[0]

#$targetSln=$args[1]

#$Variant=$args[2]


$targetSln = "C:\Users\Matt\Documents\TcXaeShell\AI_Deploy\AI_Deploy.sln"

$netId = '5.24.131.22.1.1'

$Variant  = 'Model_B'

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

####Calls to AI code
$dte = new-object -com TcXaeShell.DTE.15.0
$dte.SuppressUI = $false
$dte.MainWindow.Visible = $true

$sln = $dte.Solution
$sln.Open($targetSln)

$project = $sln.Projects.Item(1)
$systemManager = $project.Object

$systemManager.SetTargetNetId($netId)
$systemManager.CurrentProjectVariant = $Variant;
$systemManager.ActivateConfiguration()
$systemManager.StartRestartTwinCAT() 


#closes Visual studio
$dte.Quit()

####End of calls to AI code

$MessageFilter.Revoke()