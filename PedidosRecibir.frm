VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   Caption         =   "PEDIDOS A RECIBIR"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14445
   LinkTopic       =   "Form7"
   Picture         =   "PedidosRecibir.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   14445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "Regresar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      MaskColor       =   &H00E0E0E0&
      Picture         =   "PedidosRecibir.frx":1084A3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LLEGO EL PEDIDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Picture         =   "PedidosRecibir.frx":115782
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      DataField       =   "Nombre_Producto"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3840
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      RecordSource    =   "PedidosRecibir"
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   10610
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1740
      Left            =   120
      Picture         =   "PedidosRecibir.frx":122A61
      Top             =   120
      Width           =   6000
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a%, f%, c%

Private Sub Command1_Click()
c = 0
f = 1
a = Adodc1.Recordset.RecordCount
Adodc1.Recordset.MoveFirst
While (c <> a)
Adodc1.Recordset!CodigoPedido = MSFlexGrid1.TextMatrix(f, 1)
Adodc1.Recordset!Nombre_Producto = MSFlexGrid1.TextMatrix(f, 2)
Adodc1.Recordset!Proveedor = MSFlexGrid1.TextMatrix(f, 3)
Adodc1.Recordset!Cantidad_Botellas = MSFlexGrid1.TextMatrix(f, 4)
Adodc1.Recordset!ValorUnitario = MSFlexGrid1.TextMatrix(f, 5)
Adodc1.Recordset!TotalPago = MSFlexGrid1.TextMatrix(f, 6)
Adodc1.Recordset.Delete
c = c + 1
f = f + 1
Wend
End Sub

Private Sub Command2_Click()
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 80
MSFlexGrid1.FormatString = "#|< CodigoPedido |< Nombre Producto |< Nombre Proveedor |< Cantidad de botellas |<Precio Unitario |< Total Pago"
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1000
MSFlexGrid1.ColWidth(2) = 3500
MSFlexGrid1.ColWidth(3) = 3500
MSFlexGrid1.ColWidth(4) = 2000
MSFlexGrid1.ColWidth(5) = 2000
MSFlexGrid1.ColWidth(6) = 2000
Call datos
End Sub
Function datos()
c = 0
f = 1
a = Adodc1.Recordset.RecordCount
Adodc1.Recordset.MoveFirst
While (c <> a)
MSFlexGrid1.TextMatrix(f, 1) = Adodc1.Recordset!CodigoPedido
MSFlexGrid1.TextMatrix(f, 2) = Adodc1.Recordset!Nombre_Producto
MSFlexGrid1.TextMatrix(f, 3) = Adodc1.Recordset!Proveedor
MSFlexGrid1.TextMatrix(f, 4) = Adodc1.Recordset!Cantidad_Botellas
MSFlexGrid1.TextMatrix(f, 5) = Adodc1.Recordset!ValorUnitario
MSFlexGrid1.TextMatrix(f, 6) = Adodc1.Recordset!TotalPago
Adodc1.Recordset.MoveNext
c = c + 1
f = f + 1
Wend
End Function


