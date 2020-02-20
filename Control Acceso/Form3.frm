VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permisos de Usuario"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "Form3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7065
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Programas Pencas de Visual\sistemas pencas\ultima version 8.0\bodega.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "clave"
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6855
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   6240
         MouseIcon       =   "Form3.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Form3.frx":0614
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5040
         MouseIcon       =   "Form3.frx":0A56
         MousePointer    =   99  'Custom
         Picture         =   "Form3.frx":0D60
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txt_password 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmd_eliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         MouseIcon       =   "Form3.frx":106A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Eliminar"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         MouseIcon       =   "Form3.frx":1374
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancelar"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmd_volver 
         Caption         =   "&Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         MouseIcon       =   "Form3.frx":167E
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volver al Menu Principal"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Left            =   120
         Top             =   1800
      End
      Begin VB.TextBox txt_nombre 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmd_modificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MouseIcon       =   "Form3.frx":1988
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "modificar"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmd_grabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         MouseIcon       =   "Form3.frx":1C92
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Grabar"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   1200
         Top             =   0
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "Form3.frx":1F9C
         Left            =   2040
         List            =   "Form3.frx":1FA6
         MouseIcon       =   "Form3.frx":1FC3
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1320
         Width           =   2295
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label5 
         Height          =   495
         Left            =   5160
         MouseIcon       =   "Form3.frx":22CD
         MousePointer    =   99  'Custom
         TabIndex        =   18
         ToolTipText     =   "Grabar"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label4 
         Height          =   495
         Left            =   1800
         MouseIcon       =   "Form3.frx":25D7
         MousePointer    =   99  'Custom
         TabIndex        =   17
         ToolTipText     =   "Eliminar"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label3 
         Height          =   495
         Left            =   120
         MouseIcon       =   "Form3.frx":28E1
         MousePointer    =   99  'Custom
         TabIndex        =   16
         ToolTipText     =   "Modificar"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6000
         MouseIcon       =   "Form3.frx":2BEB
         MousePointer    =   99  'Custom
         Picture         =   "Form3.frx":2EF5
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre     :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lb_codigo 
         BackStyle       =   0  'Transparent
         Caption         =   "Password  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo        :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   2055
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Top             =   3045
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4895
            Text            =   "Estado:"
            TextSave        =   "Estado:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4895
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MousePointer    =   99
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AgregarOk As Boolean
Dim CodigoB, Opcion, Puerta As String
Dim k As Variant
Private Declare Function ShellAbout Lib "shell32.dll" Alias _
 "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, _
ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Combo1_Click()
cmd_grabar.Enabled = True
If txt_password.Text = "" Or txt_nombre.Text = "" Then
cmd_grabar.Enabled = False
End If
If cmd_modificar.Enabled = True Then
cmd_grabar.Enabled = False
End If

End Sub
Private Sub cmd_modificar_Click()
On Error Resume Next
cmd_grabar.Enabled = False
If txt_nombre.Text = "" Or txt_password.Text = "" Or Combo1.Text = "" Then
MsgBox "No Se Puede Modificar No hay datos en pantalla", vbCritical, ("Modificar")
cmd_modificar.Enabled = False
cmd_eliminar.Enabled = False
Exit Sub
Else
If MsgBox("    ¿Confirma los cambios en " & " " & txt_nombre & "?", vbYesNo, "Modificacion") = vbYes Then
Data1.UpdateRecord
Data1.Refresh
Data1.Recordset.Edit
Data1.Recordset.Fields!nombre = txt_nombre
Data1.Recordset.Fields!Password = txt_password
Data1.Recordset.Fields!grupo = Combo1
Data1.UpdateRecord
Timer1.Interval = 25
ProgressBar1.Min = 0
ProgressBar1.Max = 100
MostrarProgresoEnStatusBar True
Timer1.Interval = 50
StatusBar1.Panels(1).Text = "Estado: Procesando..."
txt_nombre.Text = ""
txt_password.Text = ""
Combo1.Text = ""
txt_nombre.SetFocus
cmd_modificar.Enabled = False
cmd_eliminar.Enabled = False
Else
MsgBox "Los cambios no se realizaron", vbInformation, Me.Caption
txt_password.Text = ""
cmd_modificar.Enabled = False
cmd_eliminar.Enabled = False
txt_nombre.Text = ""
Combo1.Text = ""
txt_nombre.SetFocus
End If
End If
End Sub
Private Sub cmd_eliminar_Click()
cmd_grabar.Enabled = False
On Error Resume Next
If Data1.Recordset.RecordCount = 0 Then
MsgBox "No hay datos que Eliminar", vbExclamation, ("Usuarios")
txt_nombre.SetFocus
Exit Sub
Else
If MsgBox("    ¿Desea Eliminar a" & " " & txt_nombre & "?", vbYesNo, "Eliminacion") = vbYes Then
If MsgBox("    ¿Realmente esta seguro de eliminar a " & " " & txt_nombre & "?", vbYesNo + vbExclamation, "Confirmacion de Eliminacion") = vbYes Then
Else
MsgBox txt_nombre & "   No sera Eliminado(a)", vbInformation, Me.Caption
cmd_eliminar.Enabled = False
cmd_modificar.Enabled = False
txt_nombre.Text = ""
txt_password.Text = ""
Combo1.Text = ""
txt_nombre.SetFocus
Exit Sub
End If
Data1.Recordset.Delete
cmd_cancelar.Enabled = False
Data1.Refresh
Timer1.Interval = 25
ProgressBar1.Min = 0
ProgressBar1.Max = 100
MostrarProgresoEnStatusBar True
Timer1.Interval = 50
StatusBar1.Panels(1).Text = "Estado: Procesando..."
cmd_eliminar.Enabled = False
cmd_modificar.Enabled = False
txt_nombre.SetFocus
Else
MsgBox txt_nombre & "   No sera Eliminado(a)", vbExclamation, Me.Caption
cmd_eliminar.Enabled = False
cmd_modificar.Enabled = False
txt_nombre.Text = ""
txt_password.Text = ""
Combo1.Text = ""
txt_nombre.SetFocus
Exit Sub
End If
End If
End Sub
Private Sub cmd_cancelar_Click()
cmd_grabar.Enabled = False
txt_nombre.Text = ""
txt_password.Text = ""
Combo1.Text = ""
txt_nombre.SetFocus
cmd_eliminar.Enabled = False
cmd_modificar.Enabled = False
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
End If
If (KeyAscii >= 1 And KeyAscii <= 7 Or KeyAscii >= 9 And KeyAscii <= 12 Or KeyAscii >= 14 And KeyAscii <= 26 Or KeyAscii >= 28 And KeyAscii <= 31 Or KeyAscii >= 33 And KeyAscii <= 64) Then
KeyAscii = 0
End If
If (KeyAscii >= 91 And KeyAscii <= 96 Or KeyAscii >= 123 And KeyAscii <= 255) Then
KeyAscii = 0
End If
If (KeyAscii >= 1 And KeyAscii <= 7 Or KeyAscii >= 9 And KeyAscii <= 12 Or KeyAscii >= 14 And KeyAscii <= 26 Or KeyAscii >= 28 And KeyAscii <= 31 Or KeyAscii >= 33 And KeyAscii <= 47 Or KeyAscii >= 65 And KeyAscii <= 122 Or KeyAscii >= 58 And KeyAscii <= 255) Then
KeyAscii = 0
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub
Private Sub Form_Load()
txt_password.Text = ""
Data1.DatabaseName = App.Path & "\bodega.mdb"
Data2.DatabaseName = App.Path & "\bodega.mdb"
End Sub
Private Sub cmd_grabar_Click()
On Error Resume Next
If txt_nombre = "" Or txt_password = "" Or Combo1 = "" Then
MsgBox " Debe llenar todos los campos antes de grabar", vbCritical, Me.Caption
cmd_grabar.Enabled = False
txt_nombre.SetFocus
Exit Sub
Else
If MsgBox("¿Son Correctos los datos?", vbYesNo, "Grabar Registro") = vbYes Then
cmd_eliminar.Enabled = False
cmd_modificar.Enabled = False
Data1.Recordset.AddNew
Data1.Recordset.Fields!Password = txt_password
Data1.Recordset.Fields!nombre = txt_nombre
Data1.Recordset.Fields!grupo = Combo1
Data1.UpdateRecord
Timer1.Interval = 25
ProgressBar1.Min = 0
ProgressBar1.Max = 100
MostrarProgresoEnStatusBar True
Timer1.Interval = 50
StatusBar1.Panels(1).Text = "Estado: Procesando..."
End If
End If
End Sub
Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_nombre.Text = LTrim(txt_nombre.Text)
If KeyAscii = 13 Then
If Trim(txt_nombre) <> "" Then
Data1.RecordSource = "select * from clave where NOMBRE='" & txt_nombre & "'"
Data1.Refresh
If Data1.Recordset.RecordCount <> 0 Then
txt_nombre = Trim(StrConv(txt_nombre, vbProperCase))

MsgBox txt_nombre & "  Ya existe en la base de datos", vbInformation, Me.Caption

'MsgBox "Este Registro ya existe en la base de datos", vbInformation, Me.Caption
txt_password.Text = Data1.Recordset.Fields!Password
Combo1.Text = Data1.Recordset.Fields!grupo
cmd_eliminar.Enabled = True
cmd_modificar.Enabled = True
cmd_grabar.Enabled = False
cmd_cancelar.SetFocus
End If
End If
End If
End Sub
Private Sub txt_nombre_LostFocus()
txt_nombre = Trim(StrConv(txt_nombre, vbProperCase))
End Sub
Private Sub txt_password_KeyPress(KeyAscii As Integer)
txt_password.Text = LTrim(txt_password.Text)
End Sub
Private Sub txt_password_LostFocus()
txt_password = Trim(StrConv(txt_password, vbProperCase))
End Sub
Private Sub Timer2_Timer()
If Form3.Caption = "Permisos de Usuarios" Then
Form3.Caption = "Password"
Else: Form3.Caption = "Permisos de Usuarios"
End If
End Sub
Private Sub cmd_volver_Click()
Unload Me
End Sub
Private Sub MostrarProgresoEnStatusBar(ByVal MostrarProgressBar As Boolean)
Dim tRC As RECT
If MostrarProgressBar Then
SendMessageAny StatusBar1.hwnd, SB_GETRECT, 1, tRC
With tRC
.Top = (.Top * Screen.TwipsPerPixelY)
.Left = (.Left * Screen.TwipsPerPixelX)
.Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
.Right = (.Right * Screen.TwipsPerPixelX) - .Left
End With
With ProgressBar1
SetParent .hwnd, StatusBar1.hwnd
.Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
.Visible = True
.Value = 0
End With
End If
End Sub
Private Sub Timer1_Timer()
Static lCount As Long
lCount = lCount + 10
If lCount > 100 Then
MostrarProgresoEnStatusBar False
StatusBar1.Panels(3).Text = ""
Timer1.Interval = 0
Else
StatusBar1.Panels(3).Text = Round((ProgressBar1.Value * 100 / ProgressBar1.Max), 0) & " %"
ProgressBar1.Value = lCount
End If
If lCount = 100 Then
StatusBar1.Panels(1).Text = "Estado: completado..."
StatusBar1.Panels(3).Text = "100%"
If cmd_grabar.Enabled = False And cmd_cancelar.Enabled = True Then
MsgBox "Registro Modificado", vbInformation, Me.Caption
ElseIf cmd_cancelar.Enabled = False Then
MsgBox txt_nombre & "  Fue Eliminado(a)", vbExclamation, Me.Caption
txt_nombre.Text = ""
txt_password.Text = ""
Combo1.Text = ""
cmd_cancelar.Enabled = True
Else
MsgBox txt_nombre & "  Fue ingresado(a) a la base de datos", vbInformation, "Grabar Registro"

'MsgBox "Registro ingresado a la base de datos", vbInformation, "Grabar Registro"
cmd_modificar.Enabled = False
cmd_eliminar.Enabled = False
End If
cmd_grabar.Enabled = False
ProgressBar1.Visible = False
StatusBar1.Panels(1).Text = "Estado: "
StatusBar1.Panels(3).Text = ""
txt_nombre.Text = ""
txt_password.Text = ""
Combo1.Text = ""
txt_nombre.SetFocus
lCount = 0
Timer1.Interval = 0
Else
StatusBar1.Panels(3).Text = Round((ProgressBar1.Value * 100 / ProgressBar1.Max), 0) & " %"
ProgressBar1.Value = lCount
End If
End Sub




