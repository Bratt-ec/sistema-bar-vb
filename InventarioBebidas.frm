VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   Caption         =   "Inventario"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15045
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   Picture         =   "InventarioBebidas.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Left            =   10440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "InventarioBebidas.frx":1084A3
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Codigo Proveedor"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      Picture         =   "InventarioBebidas.frx":115782
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Nombre"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      Picture         =   "InventarioBebidas.frx":21DC25
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Codigo"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      Picture         =   "InventarioBebidas.frx":3260C8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Buscar 
      BackColor       =   &H00000000&
      Caption         =   "Buscar"
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
      Left            =   7440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "InventarioBebidas.frx":42E56B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "N_Existencias"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   2040
      TabIndex        =   1
      Top             =   4920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7011
      _Version        =   393216
      BackColor       =   12640511
      BackColorFixed  =   12640511
      ForeColorSel    =   8438015
      BackColorBkg    =   16777215
      Appearance      =   0
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "StockBotellas"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTARIO DE BEBIDAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   4080
      TabIndex        =   0
      Top             =   1800
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   1740
      Left            =   4200
      Picture         =   "InventarioBebidas.frx":43B84A
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim f%, c%, a%, m%

Private Sub Command1_Click()
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
Call GridInventario
Call datos
End Sub
Function GridInventario()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 80
MSFlexGrid1.FormatString = "#|< Codigo |< Nombre |< N° Existencias|< Precio Compra |< Precio Venta |< Cod_Proveedor"
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1000
MSFlexGrid1.ColWidth(2) = 2800
MSFlexGrid1.ColWidth(3) = 1300
MSFlexGrid1.ColWidth(4) = 2500
MSFlexGrid1.ColWidth(5) = 1500
MSFlexGrid1.ColWidth(6) = 1800
End Function
Function datos()
c = 0
f = 1
a = Adodc1.Recordset.RecordCount
Adodc1.Recordset.MoveFirst
While (c <> a)
MSFlexGrid1.TextMatrix(f, 1) = Adodc1.Recordset!Codigo_Botella
MSFlexGrid1.TextMatrix(f, 2) = Adodc1.Recordset!Nombre_Botella
MSFlexGrid1.TextMatrix(f, 3) = Adodc1.Recordset!N_Existencias
MSFlexGrid1.TextMatrix(f, 4) = Adodc1.Recordset!Precio
MSFlexGrid1.TextMatrix(f, 5) = Adodc1.Recordset!PrecioVenta
MSFlexGrid1.TextMatrix(f, 6) = Adodc1.Recordset!CodigoProveedor
Adodc1.Recordset.MoveNext
c = c + 1
f = f + 1
Wend
End Function


Private Sub Option1_Click()
Text2.Enabled = True
Text3.Enabled = False
Text4.Enabled = False
Text4.Text = ""
Text3.Text = ""
End Sub

Private Sub Option2_Click()
Text2.Enabled = False
Text3.Enabled = True
Text4.Enabled = False
Text4.Text = ""
Text2.Text = ""
End Sub

Private Sub Option3_Click()
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = True
Text2.Text = ""
Text3.Text = ""
End Sub
Function buscarInv()
c = 0
Adodc1.Recordset.MoveFirst
a = Adodc1.Recordset.RecordCount
While (c <> a)
If Adodc1.Recordset!Codigo_Botella = Val(Text2.Text) Or Adodc1.Recordset!Nombre_Botella = Trim(Text3.Text) Or Adodc1.Recordset!CodigoProveedor = Trim(Text4.Text) Then
c = Adodc1.Recordset.RecordCount
Call GridInventario
MSFlexGrid1.TextMatrix(1, 1) = Adodc1.Recordset!Codigo_Botella
MSFlexGrid1.TextMatrix(1, 2) = Adodc1.Recordset!Nombre_Botella
MSFlexGrid1.TextMatrix(1, 3) = Adodc1.Recordset!N_Existencias
MSFlexGrid1.TextMatrix(1, 4) = Adodc1.Recordset!Precio
MSFlexGrid1.TextMatrix(1, 5) = Adodc1.Recordset!PrecioVenta
MSFlexGrid1.TextMatrix(1, 6) = Adodc1.Recordset!CodigoProveedor
Else
Adodc1.Recordset.MoveNext
c = c + 1
End If
'm = MsgBox("!EL ARTICULO YA NO SE ENCUENTRA EN INVENTARIO!", vbCritical)
'm = MsgBox("HACER UN PEDIDO AL PROVEEDOR", vbYesNo)
'If m = 6 Then
'Form6.Show

Wend

End Function
Private Sub Buscar_click()
Call buscarInv
End Sub
