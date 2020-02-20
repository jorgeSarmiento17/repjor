VERSION 5.00
Begin VB.Form principal 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   6375
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8895
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1320
      Top             =   0
   End
   Begin VB.Image Image2 
      Height          =   1050
      Left            =   7920
      MouseIcon       =   "principal.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "principal.frx":0614
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -1200
      Picture         =   "principal.frx":2238
      Top             =   -840
      Width           =   15360
   End
   Begin VB.Menu productos 
      Caption         =   "Productos"
      Begin VB.Menu Ingreso 
         Caption         =   "Ingreso"
      End
      Begin VB.Menu modificacion 
         Caption         =   "Modificacion"
      End
      Begin VB.Menu busqueda 
         Caption         =   "Busqueda"
      End
      Begin VB.Menu eliminacion 
         Caption         =   "Eliminacion"
      End
      Begin VB.Menu listadoa 
         Caption         =   "Listado"
      End
   End
   Begin VB.Menu clientes 
      Caption         =   "Clientes"
      Begin VB.Menu ingresoc 
         Caption         =   "Ingreso"
      End
      Begin VB.Menu modificacionc 
         Caption         =   "Modificacion"
      End
      Begin VB.Menu busquedac 
         Caption         =   "Busqueda"
      End
      Begin VB.Menu eliminacionc 
         Caption         =   "Eliminacion"
      End
   End
   Begin VB.Menu proveedores 
      Caption         =   "Proveedores"
      Begin VB.Menu ingresop 
         Caption         =   "Ingreso"
      End
      Begin VB.Menu modificacionp 
         Caption         =   "Modificacion"
      End
      Begin VB.Menu busquedap 
         Caption         =   "Busqueda"
      End
      Begin VB.Menu eliminarp 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu permisos 
      Caption         =   "Permisos"
      Begin VB.Menu cambiar 
         Caption         =   "Cambiar de Usuario"
      End
      Begin VB.Menu permisu 
         Caption         =   "Permisos"
      End
      Begin VB.Menu listado 
         Caption         =   "Listado de Usuarios"
      End
      Begin VB.Menu Ingresosis 
         Caption         =   "Ingresos al Sistema"
      End
   End
   Begin VB.Menu acerca 
      Caption         =   "Acerca de..."
      Begin VB.Menu as 
         Caption         =   "Acerca de..."
      End
   End
   Begin VB.Menu terminat 
      Caption         =   "Terminar"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellAbout Lib "shell32.dll" Alias _
"ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, _
ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Sub acercade_Click()
Form1.Show 1
End Sub
Private Sub as_Click()
Call ShellAbout(Me.hwnd, "Password y Permisos de Usuario ©2003", "Ricardo99@Chile.com ,Ricardo Sanzana Analista de sistemas, Version 6.0 ", Me.Icon)
End Sub
Private Sub busqueda_Click()
MsgBox "Aqui va el formulario de Busqueda de Productos" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub
Private Sub busquedac_Click()
MsgBox "Aqui va el formulario de Busqueda de clientes" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub
Private Sub busquedap_Click()
MsgBox "Aqui va el formulario de Busqueda de Proveedores" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub
Private Sub cambiar_Click()
Dim Mensaje, Estilo, titulo, respuesta
Mensaje = "¿Desea Terminar Sesion?" & " " & Form_iii.txt_nombre
Estilo = vbYesNo + vbQuestion
titulo = "Terminar Sesion"
respuesta = MsgBox(Mensaje, Estilo, titulo)
If respuesta = vbYes Then
Unload Me
Load Form_iii
Form_iii.Show
Form_iii.Frame1.Visible = False
Form_iii.cmd_salir.Enabled = True
Form_iii.cmdCancel.Enabled = True
Form_iii.cmdOK.Enabled = True
Form_iii.MousePointer = 0
Form_iii.Label1.MousePointer = 0
Form_iii.lb_codigo.MousePointer = 0

Form_iii.Frame2.MousePointer = 0
Form_iii.Label2.Caption = "BIENVENIDO(A) AL SISTEMA : "
Form_iii.Label3.Caption = ""
Form_iii.Label4.Caption = ""
Form_iii.StatusBar1.Panels(1).Text = "Estado:"
Form_iii.StatusBar1.Panels(3).Text = ""
Form_iii.ProgressBar1.Visible = False
Form_iii.Label4.ForeColor = &H0&
Form_iii.txt_nombre.Text = ""
Form_iii.txt_Password.Text = ""
Form_iii.txt_nombre.SetFocus
Else
Exit Sub
End If
End Sub
Private Sub eliminacion_Click()
MsgBox "Aqui va el formulario de Eliminacion de Productos" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub
Private Sub eliminacionc_Click()
MsgBox "Aqui va el formulario de Eliminacion de clientes" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub
Private Sub eliminarp_Click()
MsgBox "Aqui va el formulario de Eliminacion de Proveedores" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub




Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu permisos
If Button = 1 Then PopupMenu productos
End Sub

Private Sub ingreso_Click()
MsgBox "Aqui va el formulario de Ingreso de Productos" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub
Private Sub ingresoc_Click()
MsgBox "Aqui va el formulario de Ingreso de clientes" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub
Private Sub lblData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblData.ToolTipText = Format(Date, "dd/mm/yyyy")
End Sub
Private Sub ingresop_Click()
MsgBox "Aqui va el formulario de Ingreso de Proveedores" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub

Private Sub Ingresosis_Click()
Form2.Show 1
End Sub

Private Sub listado_Click()
Form1.Show 1
End Sub

Private Sub listadoa_Click()
MsgBox "Aqui va el formulario de Listado de Productos" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub

Private Sub modificacion_Click()
MsgBox "Aqui va el formulario de Modificacion de Productos" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub
Private Sub modificacionc_Click()
MsgBox "Aqui va el formulario de Modificacion de clientes" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub
Public Sub SetFrmMenu()
Dim rs As Recordset
Dim C1 As Variant
Dim C2 As Variant
'Set db = OpenDatabase("c:\bodega.mdb")
Set db = OpenDatabase(App.Path & "\bodega.mdb")
sql = ""
sql = "select * from clave"
sql = sql + " where nombre= '" & Form_iii.txt_nombre & "'"

'sql = sql + " where password= '" & Form_iii.txt_password & "'"
Set rs = db.OpenRecordset(sql, 2)
If rs.RecordCount > 0 Then
C1 = rs!grupo
C2 = rs!nombre
End If
C1 = Trim(StrConv(C1, vbUpperCase))
If C1 = "ADMINISTRADOR" Then
With principal
.eliminacionc.Enabled = True
End With
End If
If C1 = "USUARIO" Then
With principal
.Ingreso.Enabled = False
.modificacion.Enabled = False
.eliminacion.Enabled = False
.modificacionc.Enabled = False
.eliminacionc.Enabled = False
.permisu.Enabled = False
.busqueda.Enabled = False
.ingresoc.Enabled = False
.busquedac.Enabled = False
.listadoa.Enabled = False
.listado.Enabled = False
.ingresop.Enabled = False
.busquedap.Enabled = False
.modificacionp.Enabled = False
.eliminarp.Enabled = False
.Ingresosis.Enabled = False
End With
End If
End Sub
Private Sub modificacionp_Click()
MsgBox "Aqui va el formulario de Modificacion de Proveedores" & " " & Form_iii.txt_nombre, vbDefaultButton1, Me.Caption
End Sub
Private Sub permisu_Click()
Form3.Show 1
End Sub
Private Sub salir_Click()
Dim Mensaje, Estilo, titulo, respuesta
Mensaje = "¿Desea salir?" & " " & Form_iii.txt_nombre
Estilo = vbYesNo + vbQuestion
titulo = "Terminar"
respuesta = MsgBox(Mensaje, Estilo, titulo)
If respuesta = vbYes Then
End
Else
Exit Sub
End If
Set principal = Nothing
Unload Me
End Sub
Private Sub Timer2_Timer()
If principal.Caption = "Password" Then
principal.Caption = "Ricardo Sanzana 2003 ©"
Else: principal.Caption = "Password"
End If
End Sub

