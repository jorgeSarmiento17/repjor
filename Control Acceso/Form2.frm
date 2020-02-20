VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos al sistema"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   6495
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   120
         Top             =   240
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Ingresos al Sistema"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   480
         MouseIcon       =   "Form2.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   360
         Width           =   5055
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   120
         Y1              =   960
         Y2              =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   6360
         X2              =   6360
         Y1              =   960
         Y2              =   240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   6360
         X2              =   120
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   6360
         X2              =   120
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         X1              =   6240
         X2              =   6240
         Y1              =   840
         Y2              =   360
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF0000&
         X1              =   6240
         X2              =   240
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FF0000&
         X1              =   6240
         X2              =   240
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         X1              =   240
         X2              =   240
         Y1              =   840
         Y2              =   360
      End
   End
   Begin VB.CommandButton cmd_cancelar 
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
      Left            =   5280
      MouseIcon       =   "Form2.frx":0614
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cerrar"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Contar"
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
      Left            =   240
      MouseIcon       =   "Form2.frx":091E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Contar"
      Top             =   2760
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   200
      Cols            =   6
      FixedCols       =   0
      ForeColorFixed  =   16711680
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancelar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
sql = ""
sql = "select count(nombre) as valor from ingresoclave "
Set rs = db.OpenRecordset(sql, 2)
a% = rs.Fields(0)
MsgBox "El Numero de Ingresos al Sistema " & Form_iii.txt_nombre & " es: " & a%, vbInformation, Me.Caption
End Sub

Private Sub Form_Load()
On Error GoTo control
grilla.TextMatrix(0, 0) = "Nombre"
grilla.TextMatrix(0, 1) = "Password"
grilla.TextMatrix(0, 2) = "Fecha"
grilla.TextMatrix(0, 3) = "Hora"
grilla.TextMatrix(0, 4) = "Status"
grilla.TextMatrix(0, 5) = "Grupo"
grilla.AddItem ""
grilla.ColWidth(0) = 3000
grilla.ColWidth(1) = 2000
grilla.ColWidth(2) = 1000
grilla.ColWidth(3) = 1000
grilla.ColWidth(4) = 2000
grilla.ColWidth(5) = 2000

'Set db = OpenDatabase("c:\bodega.mdb")
Set db = OpenDatabase(App.Path & "\bodega.mdb")
sql = "select * from ingresoclave "
Set rs = db.OpenRecordset(sql, 2)
If rs.RecordCount > 0 Then
While Not rs.EOF
grilla.TextMatrix(grilla.Row, 0) = rs!nombre
grilla.TextMatrix(grilla.Row, 1) = rs!Password
grilla.TextMatrix(grilla.Row, 2) = rs!fecha
grilla.TextMatrix(grilla.Row, 3) = rs!hora
grilla.TextMatrix(grilla.Row, 4) = rs!Status
grilla.TextMatrix(grilla.Row, 5) = rs!grupo
grilla.Row = grilla.Row + 1
rs.MoveNext
Wend
grilla.Rows = grilla.Rows - 1
End If
control:
If Err.Number > 0 Then
Resume Next
End If
End Sub
Private Sub Timer1_Timer()
If Form2.Caption = "Ingresos al Sistema" Then
Form2.Caption = "Password"
Else: Form2.Caption = "Ingresos al Sistema"
End If
If Line7.BorderColor = &HFFFFFF Then
Line7.BorderColor = &HFF0000
Else: Line7.BorderColor = &HFFFFFF
End If
If Line2.BorderColor = &HFFFFFF Then
Line2.BorderColor = &HFF0000
Else: Line2.BorderColor = &HFFFFFF
End If
If Line5.BorderColor = &HFF0000 Then
Line5.BorderColor = &H0&
Else: Line5.BorderColor = &HFF0000
End If
If Line6.BorderColor = &HFF0000 Then
Line6.BorderColor = &H0&
Else: Line6.BorderColor = &HFF0000
End If
If Line9.BorderColor = &HFF0000 Then
Line9.BorderColor = &H0&
Else: Line9.BorderColor = &HFF0000
End If
If Line8.BorderColor = &HFF0000 Then
Line8.BorderColor = &H0&
Else: Line8.BorderColor = &HFF0000
End If
If Line3.BorderColor = &HFFFFFF Then
Line3.BorderColor = &HFF0000
Else: Line3.BorderColor = &HFFFFFF
End If
If Line4.BorderColor = &HFFFFFF Then
Line4.BorderColor = &HFF0000
Else: Line4.BorderColor = &HFFFFFF
End If
End Sub


