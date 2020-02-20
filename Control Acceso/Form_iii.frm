VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form Form_iii 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contraseña"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "Form_iii.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Bodega.mdb"
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
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Bodega.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ingresoclave"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H8000000A&
      Height          =   3015
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   6120
         Picture         =   "Form_iii.frx":030A
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   360
         Top             =   480
      End
      Begin VB.Timer TimerSalir 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   480
         Top             =   2280
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "BIENVENIDO(A) AL SISTEMA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   5895
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   5895
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Top             =   3030
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4710
            Text            =   "Estado:"
            TextSave        =   "Estado:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4710
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
      Height          =   2895
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6495
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5040
         MouseIcon       =   "Form_iii.frx":0614
         MousePointer    =   99  'Custom
         Picture         =   "Form_iii.frx":091E
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Left            =   120
         Top             =   1800
      End
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Salir"
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
         Left            =   4680
         MouseIcon       =   "Form_iii.frx":0C28
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Aceptar"
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
         Left            =   360
         MouseIcon       =   "Form_iii.frx":0F32
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aceptar"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Limpiar"
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
         Left            =   2520
         MouseIcon       =   "Form_iii.frx":123C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpiar"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txt_Password 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txt_nombre 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2640
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lb_codigo 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese Password :"
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
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese Nombre     :"
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
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   5760
         MouseIcon       =   "Form_iii.frx":1546
         MousePointer    =   99  'Custom
         Picture         =   "Form_iii.frx":1850
         Top             =   720
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form_iii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Unidad As Long
Private limite As Long
Private Progreso As Long
Dim i As Long
Dim a As String
Dim B As String
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"
End If
End Sub
Private Sub Form_Load()
StatusBar1.Panels(1).Text = "Estado:"
Data1.DatabaseName = App.Path & "\bodega.mdb"
Data2.DatabaseName = App.Path & "\bodega.mdb"
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ToolTipText = Label7.Caption
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ToolTipText = Label8.Caption
End Sub
Private Sub Timer1_Timer()
Dim F As String
c = Data1.Recordset.Fields!grupo
Label3 = Trim(StrConv(Label3, vbUpperCase))

B = "CARGANDO SU CONFIGURACION PERSONAL..."
F = "  "
Static lCount As Long
lCount = lCount + 1
If lCount > 100 Then
MostrarProgresoEnStatusBar False
StatusBar1.Panels(3).Text = ""
Timer1.Interval = 0
Else
StatusBar1.Panels(3).Text = Round((ProgressBar1.Value * 100 / ProgressBar1.Max), 0) & " %"
ProgressBar1.Value = lCount
cmdOK.Enabled = False
cmdCancel.Enabled = False
cmd_salir.Enabled = False
StatusBar1.MousePointer = 11
Frame2.MousePointer = 11
lb_codigo.MousePointer = 11
Label1.MousePointer = 11
Picture1.MousePointer = 11
If lCount = 3 Then
Frame1.Visible = True
Label2.Caption = Label2.Caption
Label3.Caption = Label3.Caption & a
Label4.Caption = Label4.Caption & B
Label4.ForeColor = &HC00000

End If
If lCount = 10 Then
Frame1.MousePointer = 11
Picture1.MousePointer = 11
End If
If lCount = 40 Then
Frame1.MousePointer = 13
Picture1.MousePointer = 13

End If
If lCount = 80 Then
Frame1.MousePointer = 11
Picture1.MousePointer = 11

End If
If lCount = 100 Then
StatusBar1.Panels(1).Text = "Estado: Completado..."
StatusBar1.Panels(3).Text = "100%"
principal.Show
Form_iii.Hide
Call principal.SetFrmMenu
lCount = 0
Timer1.Interval = 0
Else
StatusBar1.Panels(3).Text = Round((ProgressBar1.Value * 100 / ProgressBar1.Max), 0) & " %"
ProgressBar1.Value = lCount
End If
End If
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
Private Sub cmdCancel_Click()
ProgressBar1.Visible = False
StatusBar1.Panels(1).Text = "Estado:"
txt_Password.Text = ""
txt_nombre.Text = ""
txt_nombre.SetFocus
End Sub
Private Sub cmdOK_Click()
Dim access As Variant
Dim access1 As Variant
Dim r As String
Dim c As String
Dim s As String
s = "Error"
On Error Resume Next
r = txt_nombre
a = txt_nombre
Data1.RecordSource = "select * from clave where password= '" & txt_Password & "'"
Data1.RecordSource = "select * from clave where nombre= '" & txt_nombre & "'"

Data1.Refresh
If Trim(txt_Password.Text) = "" And Trim(txt_nombre.Text) = "" Then
MsgBox "Debe Ingresar Password/Nombre", vbOKOnly, Me.Caption
txt_nombre.SetFocus
Exit Sub
ElseIf Trim(txt_Password) = "" Then
MsgBox "Debe Ingresar la Password", vbOKOnly, Me.Caption
txt_Password.SetFocus
Exit Sub
ElseIf Trim(txt_nombre) = "" Then
MsgBox "Debe Ingresar el nombre", vbOKOnly, Me.Caption
txt_nombre.SetFocus
Exit Sub
End If
If Data1.Recordset.RecordCount = 1 Then
If Data1.Recordset.Fields!nombre = txt_nombre.Text And Data1.Recordset.Fields!Password = txt_Password Then
c = Data1.Recordset.Fields!grupo
Timer1.Interval = 25
ProgressBar1.Min = 0
ProgressBar1.Max = 100
MostrarProgresoEnStatusBar True
Timer1.Interval = 50
StatusBar1.Panels(1).Text = "Estado: Procesando..."
access = "AUTORIZADO"
Data2.Recordset.AddNew
Data2.Recordset.Fields!nombre = txt_nombre
Data2.Recordset.Fields!Password = txt_Password
Data2.Recordset.Fields!fecha = Format(Now, "dd/mm/yyyy")
Data2.Recordset.Fields!hora = Format(Time, "hh:mm:ss ")
Data2.Recordset.Fields!Status = access
Data2.Recordset.Fields!grupo = c
Data2.UpdateRecord
Else
'este codigo es necesario cuando la password es correcta y el nombre es erroneo
MsgBox "Password/Nombre Erroneo, Acceso Denegado   ", vbCritical, Me.Caption
access = "DENEGADO"
Data2.Recordset.AddNew
Data2.Recordset.Fields!nombre = r
Data2.Recordset.Fields!Password = txt_Password
Data2.Recordset.Fields!fecha = Format(Now, "dd/mm/yyyy")
Data2.Recordset.Fields!hora = Format(Time, "hh:mm:ss ")
Data2.Recordset.Fields!Status = access
Data2.Recordset.Fields!grupo = s
Data2.UpdateRecord
txt_Password.Text = ""
txt_nombre.Text = ""
txt_nombre.SetFocus
End If
Else
'este codigo se utiliza cuando la password y el nombre son erroneos
MsgBox "Password/Nombre Erroneo, Acceso Denegado   ", vbCritical, Me.Caption
access = "DENEGADO"
Data2.Recordset.AddNew
Data2.Recordset.Fields!nombre = r
Data2.Recordset.Fields!Password = txt_Password
Data2.Recordset.Fields!fecha = Format(Now, "dd/mm/yyyy")
Data2.Recordset.Fields!hora = Format(Time, "hh:mm:ss ")
Data2.Recordset.Fields!Status = access
Data2.Recordset.Fields!grupo = s
Data2.UpdateRecord
txt_Password.Text = ""
txt_nombre.Text = ""
txt_nombre.SetFocus
End If
End Sub
Private Sub cmd_salir_Click()
Dim Mensaje, Estilo, titulo, respuesta
Mensaje = "¿Desea Salir?"
Estilo = vbYesNo
titulo = "Terminar"
respuesta = MsgBox(Mensaje, Estilo, titulo)
If respuesta = vbYes Then
Unload Me
Else
txt_nombre.SetFocus
Exit Sub
End If
Set principal = Nothing
Unload Me
End Sub
Private Sub Timer2_Timer()
If Form_iii.Caption = "Contraseña" Then
Form_iii.Caption = "Password"
Else: Form_iii.Caption = "Contraseña"
End If
End Sub
Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
txt_nombre.Text = LTrim(txt_nombre.Text)
End Sub
Private Sub txt_nombre_LostFocus()
txt_nombre = Trim(StrConv(txt_nombre, vbProperCase))
End Sub
Private Sub txt_nombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txt_nombre.ToolTipText = txt_nombre.Text
End Sub
Private Sub txt_password_LostFocus()
txt_Password = Trim(StrConv(txt_Password, vbProperCase))
End Sub

