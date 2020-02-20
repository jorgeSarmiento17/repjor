Attribute VB_Name = "Module1"
Public access As String
Public db As Database
Public rs As Recordset
Public var_id As Integer
Public up As String
Public sql As String
'Public var_id As Integer
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Global ContadorSalir As Integer


