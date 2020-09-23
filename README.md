<div align="center">

## Is Process Running


</div>

### Description

This is do determine if any exe is already running. This handles the multithreading issues of NT, and it works on 95,98,NT. I got most of this straight from Microsoft, but have wrapped and cleaned it up alot.
 
### More Info
 
Pass it the EXE Name you want to know if it is running

Test it with Notepad.exe or something first, and keep in mind that if you are debugging the application you want to use this with, that VB6.EXE is the process you are running, and until you compile and run your exe it will not see the app you are using.

True if process is running, else false.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[James N\. Wink](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-n-wink.md)
**Level**          |Advanced
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/james-n-wink-is-process-running__1-5449/archive/master.zip)

### API Declarations

```
Option Explicit
'Used to determine Process Information
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const MAX_PATH = 260
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const hNull = 0
Public Const WIN95_System_Found = 1
Public Const WINNT_System_Found = 2
Public Const Default_Log_Size = 10000000
Public Const Default_Log_Days = 0
'Types Used by Win API's
Public Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long   ' This process
 th32DefaultHeapID As Long
 th32ModuleID As Long   ' Associated exe
 cntThreads As Long
 th32ParentProcessID As Long  ' This process's parent process
 pcPriClassBase As Long   ' Base priority of process threads
 dwFlags As Long
 szExeFile As String * 260  ' MAX_PATH
End Type
Public Type OSVERSIONINFO
 dwOSVersionInfoSize As Long
 dwMajorVersion As Long
 dwMinorVersion As Long
 dwBuildNumber As Long
 dwPlatformId As Long   '1 = Windows 95.
         '2 = Windows NT
 szCSDVersion As String * 128
End Type
'Used to determine process information
Public Declare Function Process32First Lib "kernel32" ( _
 ByVal hSnapshot As Long, _
 lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" ( _
 ByVal hSnapshot As Long, _
 lppe As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "Kernel32.dll" ( _
 ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32.dll" ( _
 ByVal dwDesiredAccessas As Long, _
 ByVal bInheritHandle As Long, _
 ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" ( _
 ByRef lpidProcess As Long, _
 ByVal cb As Long, _
 ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" ( _
 ByVal hProcess As Long, _
 ByVal hModule As Long, _
 ByVal ModuleName As String, _
 ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" ( _
 ByVal hProcess As Long, _
 ByRef lphModule As Long, _
 ByVal cb As Long, _
 ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
 ByVal dwFlags As Long, _
 ByVal th32ProcessID As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" ( _
 lpVersionInformation As OSVERSIONINFO) As Integer
```


### Source Code

```
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsProcessRunning
'
' Date: 07/13/1999
' Comapany: WEI
' Web Site: http://www.winkenterprises.com
' Author: James N.Wink
' Email: james@winkenterprises.com
'
' Description: Used to determine if a process is running.
'
' Input: EXEName - String  EXE name of the Process
'
' Output: IsProcessRunning - Boolean Returns True if running
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsProcessRunning(ByVal EXEName As String) As Boolean
 'Used if Win 95 is detected
 Dim booResult As Boolean
 Dim lngLength As Long
 Dim lngProcessID As Long
 Dim strProcessName As String
 Dim lngSnapHwnd As Long
 Dim udtProcEntry As PROCESSENTRY32
 'Used if NT is detected
 Dim lngCBSize As Long 'Specifies the size, in bytes, of the lpidProcess array
 Dim lngCBSizeReturned As Long 'Receives the number of bytes returned
 Dim lngNumElements As Long
 Dim lngProcessIDs() As Long
 Dim lngCBSize2 As Long
 Dim lngModules(1 To 200) As Long
 Dim lngReturn As Long
 Dim strModuleName As String
 Dim lngSize As Long
 Dim lngHwndProcess As Long
 Dim lngLoop As Long
 'Turn on Error handler
 On Error GoTo IsProcessRunning_Error
 booResult = False
 EXEName = UCase$(Trim$(EXEName))
 lngLength = Len(EXEName)
Select Case getVersion()
  Case WIN95_System_Found 'Windows 95/98
  'Get SnapShot of Threads
  lngSnapHwnd = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
  'Check to see if SnapShot was made
  If lngSnapHwnd = hNull Then GoTo IsProcessRunning_Exit
  'Set Size in UDT, must be done, prior to calling API
  udtProcEntry.dwSize = Len(udtProcEntry)
  ' Get First Process
  lngProcessID = Process32First(lngSnapHwnd, udtProcEntry)
  Do While lngProcessID
   'Get Full Path Process Name
   strProcessName = StrZToStr(udtProcEntry.szExeFile)
   'Check for Matching Upper case result
   strProcessName = Ucase$(Trim$(strProcessName))
   If Right$(strProcessName, lngLength) = EXEName Then
    'Found
    booResult = True
    GoTo IsProcessRunning_Exit
   End If
   'Not found, get next Process
   lngProcessID = Process32Next(lngSnapHwnd, udtProcEntry)
  Loop
  Case WINNT_System_Found 'Windows NT
  'Get the array containing the process id's for each process objec
  '  t
  'Set Default Size
  lngCBSize = 8 ' Really needs to be 16, but Loop will increment prior to calling API
  lngCBSizeReturned = 96
  'Check to see if Process ID's were returned
  Do While lngCBSize <= lngCBSizeReturned
   'Increment Size
   lngCBSize = lngCBSize * 2
   'Allocate Memory for Array
   ReDim lngProcessIDs(lngCBSize / 4) As Long
   'Get Process ID's
   lngReturn = EnumProcesses(lngProcessIDs(1), lngCBSize, lngCBSizeReturned)
  Loop
  'Count number of processes returned
  lngNumElements = lngCBSizeReturned / 4
  'Loop thru each process
  For lngLoop = 1 To lngNumElements
   'Get a handle to the Process and Open
   lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
   Or PROCESS_VM_READ, 0, lngProcessIDs(lngLoop))
   'Check to see if Process handle was returned
   If lngHwndProcess <> 0 Then
    'Get an array of the module handles for the specified process
    lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)
    'If the Module Array is retrieved, Get the ModuleFileName
    If lngReturn <> 0 Then
     'Buffer with spaces first to allocate memory for byte array
     strModuleName = Space(MAX_PATH)
     'Must be set prior to calling API
     lngSize = 500
     'Get Process Name
     lngReturn = GetModuleFileNameExA(lngHwndProcess, lngModules(1), _
     strModuleName, lngSize)
     'Remove trailing spaces
     strProcessName = Left(strModuleName, lngReturn)
     'Check for Matching Upper case result
     strProcessName = UCase$(Trim$(strProcessName))
     If Right$(strProcessName, lngLength) = EXEName Then
      'Found
      booResult = True
      GoTo IsProcessRunning_Exit
     End If
    End If
   End If
   'Close the handle to this process
   lngReturn = CloseHandle(lngHwndProcess)
  Next
 End Select
GoTo IsProcessRunning_Exit
IsProcessRunning_Error:
Err.Clear
booResult = False
IsProcessRunning_Exit:
'Turn off Error handler
On Error GoTo 0
IsProcessRunning = booResult
End Function
Private Function getVersion() As Long
 Dim osinfo As OSVERSIONINFO
 Dim retvalue As Integer
 osinfo.dwOSVersionInfoSize = 148
 osinfo.szCSDVersion = Space$(128)
 retvalue = GetVersionExA(osinfo)
 getVersion = osinfo.dwPlatformId
End Function
Private Function StrZToStr(s As String) As String
 StrZToStr = Left$(s, Len(s) - 1)
End Function
```

