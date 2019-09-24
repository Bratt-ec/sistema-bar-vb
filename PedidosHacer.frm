VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   Caption         =   "Pedidos para hacer"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14445
   Picture         =   "PedidosHacer.frx":0000
   ScaleHeight     =   8805
   ScaleWidth      =   14445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
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
      Left            =   9840
      MaskColor       =   &H00E0E0E0&
      Picture         =   "PedidosHacer.frx":50C022
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "PedidosHacer.frx":519301
      Left            =   1080
      List            =   "PedidosHacer.frx":51932C
      TabIndex        =   20
      Text            =   "Combo2"
      Top             =   3000
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "PedidosHacer.frx":519403
      Left            =   1080
      List            =   "PedidosHacer.frx":51941F
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   2520
      Width           =   2655
   End
   Begin VB.ComboBox Proveedores 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "PedidosHacer.frx":5194B4
      Left            =   1080
      List            =   "PedidosHacer.frx":5194BE
      TabIndex        =   18
      Text            =   "Proveedores"
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hacer Pedido"
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
      Left            =   9720
      Picture         =   "PedidosHacer.frx":5194EB
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      TabIndex        =   15
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      TabIndex        =   12
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      TabIndex        =   11
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AÑADIR"
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
      Left            =   9720
      Picture         =   "PedidosHacer.frx":5267CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   8
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "Proveedor"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
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
      Height          =   3975
      Left            =   840
      TabIndex        =   0
      Top             =   4800
      Width           =   9735
      _ExtentX        =   17171
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
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
      Left            =   10920
      TabIndex        =   16
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUBTOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10800
      TabIndex        =   14
      Top             =   6720
      Width           =   1395
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
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
      Left            =   10800
      TabIndex        =   13
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Unitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1740
      Left            =   3720
      Picture         =   "PedidosHacer.frx":533AA9
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c%, f%, a%, valor As Double, iva As Double, total As Double, subtotal As Double

'Cerveza CLUB
'Cerveza Pilsener
'Cerveza Pilsener Light
'Cerveza Cristal
'Cerveza CLUB ROJA
'Cerveza CLUB NEGRA
'Cerveza Miller
'Cerveza Brahama

Private Sub Combo1_Click()
If Combo1.Text = "Cerveza CLUB" Then
Text2.Text = "789"
Text3.Text = Combo1.Text
Text6.Text = "1"
End If
If Combo1.Text = "Cerveza Pilsener" Then
Text2.Text = "788"
Text3.Text = Combo1.Text
Text6.Text = "1"
End If
If Combo1.Text = "Cerveza Pilsener Light" Then
Text2.Text = "689"
Text3.Text = Combo1.Text
Text6.Text = "0,75"
End If
If Combo1.Text = "Cerveza Cristal" Then
Text2.Text = "779"
Text3.Text = Combo1.Text
Text6.Text = "1,75"
End If
If Combo1.Text = "Cerveza CLUB ROJA" Then
Text2.Text = "712"
Text3.Text = Combo1.Text
Text6.Text = "1,10"
End If
If Combo1.Text = "Cerveza CLUB NEGRA" Then
Text2.Text = "289"
Text3.Text = Combo1.Text
Text6.Text = "1,15"
End If
If Combo1.Text = "Cerveza Miller" Then
Text2.Text = "686"
Text3.Text = Combo1.Text
Text6.Text = "1,25"
End If
If Combo1.Text = "Cerveza Brahama" Then
Text2.Text = "753"
Text3.Text = Combo1.Text
Text6.Text = "1,15"
End If
End Sub


'Jhony Walker Red
'Jhony Walker Black
'Jhony Walker Blue
'Absolut Vodka
'Vodka Ruskaya
'Ron 100 Fuegos
'Ron del Rio
'Cristal Seco
'Zhumir Seco
'Zhumir Durazno
'Zhumir Tropical
'Tequila Cuervo
'Tequila El Charro

Private Sub Combo2_Click()
If Combo2.Text = "Jhony Walker Red" Then
Text2.Text = "714"
Text3.Text = Combo2.Text
Text6.Text = "9,75"
End If
If Combo2.Text = "Jhony Walker Black" Then
Text2.Text = "751"
Text3.Text = Combo2.Text
Text6.Text = "11,8"
End If
If Combo2.Text = "Jhony Walker Blue" Then
Text2.Text = "777"
Text3.Text = Combo2.Text
Text6.Text = "102,75"
End If
If Combo2.Text = "Absolut Vodka" Then
Text2.Text = "701"
Text3.Text = Combo2.Text
Text6.Text = "9,5"
End If
If Combo2.Text = "Vodka Ruskaya" Then
Text2.Text = "212"
Text3.Text = Combo2.Text
Text6.Text = "10,10"
End If
If Combo2.Text = "Ron 100 Fuegos" Then
Text2.Text = "119"
Text3.Text = Combo2.Text
Text6.Text = "10,15"
End If
If Combo2.Text = "Ron del Rio" Then
Text2.Text = "258"
Text3.Text = Combo2.Text
Text6.Text = "4,25"
End If
If Combo2.Text = "Cristal Seco" Then
Text2.Text = "553"
Text3.Text = Combo2.Text
Text6.Text = "4,15"
End If
If Combo2.Text = "Zhumir Seco" Then
Text2.Text = "353"
Text3.Text = Combo2.Text
Text6.Text = "8,15"
End If
If Combo2.Text = "Zhumir Durazno" Then
Text2.Text = "322"
Text3.Text = Combo2.Text
Text6.Text = "7,15"
End If
If Combo2.Text = "Zhumir Tropical" Then
Text2.Text = "222"
Text3.Text = Combo2.Text
Text6.Text = "9,55"
End If
If Combo2.Text = "Tequila Cuervo" Then
Text2.Text = "992"
Text3.Text = Combo2.Text
Text6.Text = "17,15"
End If
If Combo2.Text = "Tequila El Charro" Then
Text2.Text = "992"
Text3.Text = Combo2.Text
Text6.Text = "18,55"
End If
End Sub

Private Sub Command1_Click()

MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + c, 1) = Text2.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + c, 2) = Text3.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + c, 3) = Proveedores.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + c, 4) = Val(Text4.Text) 'cantidad
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + c, 5) = Val(Text6.Text) 'valoru
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + c, 6) = (CDbl(Text4.Text) * CDbl(Text6.Text))
subtotal = (CDbl(Text4.Text) * (Val(Text6.Text))) + valor
Text8.Text = subtotal
valor = CDbl(Text8.Text)
iva = Format(CDbl(valor) * 14 / 100, "0.00")
Text7.Text = iva
total = Format(CDbl(valor + CDbl(Text7.Text)), "0.00")
Text9.Text = total
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
'Text5.Text = " "
Text6.Text = " "
'Text7.Text = " "
'Text8.Text = " "
'Text9.Text = " "
Text2.SetFocus
c = c + 1
End Sub


Private Sub Command2_Click()
Dim d As Integer, i As Integer
d = 1
For i = 1 To MSFlexGrid1.Rows - 1
 If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 1) <> " " Then
 Adodc1.Recordset.AddNew
 Adodc1.Recordset!CodigoPedido = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 1)
 Adodc1.Recordset!Nombre_Producto = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 2)
 Adodc1.Recordset!Cantidad_Botellas = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 4)
 Adodc1.Recordset!Proveedor = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 3)
 Adodc1.Recordset!ValorUnitario = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 5)
 Adodc1.Recordset!TotalPago = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 6)
 Adodc1.Recordset.Update
 d = d + 1
 Else
  i = MSFlexGrid1.Rows - 1
 End If
Next i

End Sub

Private Sub Command3_Click()
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
Call GridInventario
Combo1.Visible = False
Combo2.Visible = False
End Sub
Function GridInventario()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 80
MSFlexGrid1.FormatString = "#|< CodigoPedido |< Nombre Producto |< Nombre Proveedor |< Cantidad de botellas |<Precio Unitario |< Total Pago"
MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1000
MSFlexGrid1.ColWidth(2) = 2500
MSFlexGrid1.ColWidth(3) = 2500
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 1000
MSFlexGrid1.ColWidth(6) = 1000
End Function
Function datospedidos()
c = 0
f = 1
a = Adodc1.Recordset.RecordCount
Adodc1.Recordset.MoveFirst
While (c <> a)
MSFlexGrid1.TextMatrix(f, 1) = Text2.Text
MSFlexGrid1.TextMatrix(f, 2) = Text3.Text
MSFlexGrid1.TextMatrix(f, 3) = Text4.Text
MSFlexGrid1.TextMatrix(f, 4) = Text5.Text
MSFlexGrid1.TextMatrix(f, 5) = Text6.Text
MSFlexGrid1.TextMatrix(f, 6) = Text7.Text
Adodc1.Recordset.MoveNext
c = c + 1
f = f + 1
Wend
End Function


Private Sub Proveedores_Click()
If Proveedores.Text = "Cerveceria Nacional" Then
Combo1.Visible = True
Combo2.Visible = False
Else
Combo1.Visible = False
Combo2.Visible = True
End If
End Sub
