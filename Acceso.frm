VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Acceso"
   ClientHeight    =   4965
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7395
   Icon            =   "Acceso.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Acceso.frx":9ED32
   ScaleHeight     =   4965
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H005971AE&
      Caption         =   "INICIAR SESION"
      DisabledPicture =   "Acceso.frx":1A71D5
      DownPicture     =   "Acceso.frx":1B44B4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      MaskColor       =   &H8000000B&
      Picture         =   "Acceso.frx":1C1793
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataField       =   "User"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7680
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   8040
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Poyecto Jampi\ProyectoJanpier.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Poyecto Jampi\ProyectoJanpier.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Clave"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   1740
      Left            =   960
      Picture         =   "Acceso.frx":1CEA72
      Top             =   240
      Width           =   6000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   1080
      TabIndex        =   2
      Top             =   2160
      Width           =   1725
   End
   Begin VB.Menu Opciones 
      Caption         =   "Opciones"
      Begin VB.Menu Cerrar 
         Caption         =   "Cerrar "
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c%, a%, bandera As Boolean, al%
Function clave()
c = 0
Adodc1.Recordset.MoveFirst
a = Adodc1.Recordset.RecordCount
While (c <> a)
If Adodc1.Recordset!clave = Trim(Text2.Text) And Adodc1.Recordset!User = Trim(Text1.Text) Then
c = Adodc1.Recordset.RecordCount
bandera = True
Unload Me
Form2.Show
Else
Adodc1.Recordset.MoveNext
c = c + 1
End If
Wend
If (bandera = False) Then
al = MsgBox("INFORMACION INCORRECTA", vbCritical)
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End If
End Function

Private Sub Ayuda_Click()
'ayuda.show
End Sub

Private Sub Cerrar_Click()
Unload Me
End
End Sub

Private Sub Command1_Click()
Call clave
End Sub


Private Sub Form_Load()
'******* texto sin simbolos
'Private Sub Text3_KeyPress(KeyAscii As Integer)
'If IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
'If (KeyAscii >= 33) And (KeyAscii <= 47) Or (KeyAscii >= 58) And (KeyAscii <= 100) Or _
 '  (KeyAscii >= 91) And (KeyAscii <= 96) Or (KeyAscii >= 123) And (KeyAscii <= 126) Then
  '  KeyAscii = 8
'End If
'End Sub

'If (KeyAscii >= 48) And (KeyAscii <= 57) Then KeyAscii = 0 validar solo texto numerico
'If IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0  validar texto
'If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii = 32 Or KeyAscii = 44 Then Exit Sub validar direcciones
End Sub
