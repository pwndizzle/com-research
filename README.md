# COM Research

A collection of random bits of COM research and POC code.


- com-enum.ps1
  - A script to recursively enumerate exposed COM methods.

- clsids.txt
  - A list of CLSIDs compiled from a test system and here - https://www.scriptjunkie.us/wp-content/uploads/2019/05/iids_tiraniddo
  - The clsids-crash.txt file contains ids that were unstable.

- shellex.cpp + VS Project
  - Example code to show how to interact with COM methods using IDispatch in C++. I've included the full Visual Stuido project for easier import/testing.
  
## COM Examples

```
$Obj = [System.Activator]::CreateInstance([Type]::GetTypeFromProgID("Excel.Application"));$Obj | gm

$Obj = [System.Activator]::CreateInstance([Type]::GetTypeFromCLSID("00020812-0000-0000-C000-000000000046"));$Obj | gm
```

```
([System.Activator]::CreateInstance([Type]::GetTypeFromCLSID("C08AFD90-F2A1-11D1-8455-00A0C91F3880"))).Document.Application.ShellExecute("cmd","/c notepad")
```

```
Sub test()
Set obj = GetObject("new:C08AFD90-F2A1-11D1-8455-00A0C91F3880")
obj.Document.Application.ShellExecute "cmd", "/c calc", "", Null, 0
End Sub
```
