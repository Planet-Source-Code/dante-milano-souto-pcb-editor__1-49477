Attribute VB_Name = "modAssossiar"
Option Explicit


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
            "RegCreateKeyA" (ByVal hKey As Long, _
            ByVal lpSubKey As String, _
            phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias _
            "RegSetValueA" (ByVal hKey As Long, _
            ByVal lpSubKey As String, _
            ByVal dwType As Long, _
            ByVal lpData As String, _
            ByVal cbData As Long) As Long

' Return codes from Registration functions.
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 260&
Private Const REG_SZ = 1

Public Sub Associar()
'
'   ASSOCIANDO '*.mpcb' AO Milano PCB
'

On Error GoTo errTrat

Dim sKeyName    As String   ' Holds Key Name in registry.
Dim sKeyValue   As String   ' Holds Key Value in registry.
Dim ret         As Long     ' Holds error status if any from API calls.
Dim lphKey      As Long     ' Holds created key handle from RegCreateKey.

    ' Registrando o MilanoPCB

    'This creates a Root entry called "MyApp".
    sKeyName = "MilanoPCB"
    sKeyValue = "Milano Printed Circuit Board Editor"
    ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    ret = RegSetValue&(lphKey, "", REG_SZ, sKeyValue, 0&)
    
    'This creates a Root entry called .BAR associated with "MyApp".
    sKeyName = ".mpcb"
    sKeyValue = "MilanoPCB"
    ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    ret = RegSetValue&(lphKey, "", REG_SZ, sKeyValue, 0&)
    
    'This sets the command line for "MyApp".
    sKeyName = "MilanoPCB"
    sKeyValue = App.Path & "\MilanoPCB.exe %1"
    ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
    ret = RegSetValue&(lphKey, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)
    
    sKeyValue = App.Path & "\mpcb_file.ico"
    ret = RegSetValue&(lphKey, "DefaultIcon", REG_SZ, sKeyValue, MAX_PATH)

Exit Sub
errTrat:


End Sub


