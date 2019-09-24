VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   Caption         =   "Clientes"
   ClientHeight    =   9780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8295
   LinkTopic       =   "Form5"
   Picture         =   "Clientes.frx":0000
   ScaleHeight     =   9780
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   9720
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
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
      RecordSource    =   "Clientes"
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
   Begin VB.TextBox Text1 
      DataField       =   "Cedula_Cliente"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "REGRESAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      Picture         =   "Clientes.frx":1084A3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   15266
      _Version        =   393216
      Cols            =   3
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1740
      Left            =   360
      Picture         =   "Clientes.frx":115782
      Top             =   120
      Width           =   6000
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
Call GridClientes
Call datosclientes
End Sub
Function GridClientes()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 90
MSFlexGrid1.FormatString = "#|< Cedula |< Nombre-Apellido "
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 2000
MSFlexGrid1.ColWidth(2) = 6800

End Function
Function datosclientes()
c = 1
f = 1
a = Adodc1.Recordset.RecordCount
Adodc1.Recordset.MoveFirst
While (c <> a)
MSFlexGrid1.TextMatrix(f, 1) = Adodc1.Recordset!Cedula_Cliente
MSFlexGrid1.TextMatrix(f, 2) = Adodc1.Recordset!Apellidos_Nombre
Adodc1.Recordset.MoveNext
c = c + 1
f = f + 1
Wend
End Function
