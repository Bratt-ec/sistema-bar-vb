VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "Factura"
   ClientHeight    =   8850
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9570
   Icon            =   "Factura.frx":0000
   LinkTopic       =   "Form3"
   Picture         =   "Factura.frx":9ED32
   ScaleHeight     =   8850
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Nuevo 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   40
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Guardar2 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   39
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Eliminar 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   38
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Modificar 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   37
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Deshacer 
      Caption         =   "Deshacer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   36
      Top             =   1320
      Width           =   1215
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
      ItemData        =   "Factura.frx":1A71D5
      Left            =   4320
      List            =   "Factura.frx":1A7218
      TabIndex        =   32
      Text            =   "BEBIDAS"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text16 
      DataField       =   "Numero_Factura"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   360
      TabIndex        =   31
      Text            =   "Text16"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      DataField       =   "Codigo_Botella"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   480
      TabIndex        =   30
      Text            =   "Text15"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Numero_Factura"
      DataSource      =   "Adodc2"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Adodc2"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataField       =   "Cedula_Cliente"
      DataSource      =   "Adodc2"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      DataField       =   "Ruc_Cliente"
      DataSource      =   "Adodc2"
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
      Left            =   2160
      TabIndex        =   14
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      DataField       =   "Ruc_Empresa"
      DataSource      =   "Adodc2"
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
      Left            =   6000
      TabIndex        =   13
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      DataField       =   "Direccion"
      DataSource      =   "Adodc2"
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
      Left            =   6000
      TabIndex        =   12
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text7 
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
      Left            =   480
      TabIndex        =   11
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text8 
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
      Left            =   2520
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Text9 
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
      Left            =   4680
      TabIndex        =   9
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox Text10 
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
      Left            =   6840
      TabIndex        =   8
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton guardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   6
      Top             =   7200
      Width           =   2055
   End
   Begin VB.TextBox Text11 
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
      Left            =   7200
      TabIndex        =   5
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Text12 
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
      Left            =   7200
      TabIndex        =   4
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text13 
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
      Left            =   7200
      TabIndex        =   3
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox Text14 
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
      Left            =   7200
      TabIndex        =   2
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton añadir 
      Caption         =   "Añadir"
      DisabledPicture =   "Factura.frx":1A7380
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Picture         =   "Factura.frx":1B465F
      TabIndex        =   1
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton imprimirl 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Top             =   8040
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   8040
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "factura"
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
      Height          =   1575
      Left            =   0
      TabIndex        =   7
      Top             =   4680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   5
      Cols            =   5
      ScrollBars      =   2
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   240
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
      RecordSource    =   "DatosFactura"
      Caption         =   "Adodc2"
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   720
      TabIndex        =   35
      Top             =   3480
      Width           =   1110
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Bebida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2520
      TabIndex        =   34
      Top             =   3480
      Width           =   1995
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Unitario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4800
      TabIndex        =   33
      Top             =   3480
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   29
      Top             =   840
      Width           =   1395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   28
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cedula Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   26
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4200
      TabIndex        =   24
      Top             =   840
      Width           =   1755
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Bebidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   6840
      TabIndex        =   23
      Top             =   3120
      Width           =   1845
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   22
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label12 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6120
      TabIndex        =   21
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5640
      TabIndex        =   20
      Top             =   7680
      Width           =   1365
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total a pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5400
      TabIndex        =   19
      Top             =   8280
      Width           =   1635
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FACTURA"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   3480
      TabIndex        =   18
      Top             =   0
      Width           =   2265
   End
   Begin VB.Menu regresar 
      Caption         =   "Regresar"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valor As Currency, c%, total As Currency, subtotal As Currency, iva As Currency

Function Facturaexcel()
Dim objeto, i, NUM
Dim cadena As String
cadena = "I:\Poyecto Jampi\factura.xlsx"
Set objeto = CreateObject("excel.application")
objeto.Visible = True
objeto.Workbooks.Open FileName:=cadena
NUM = 9
For i = 1 To MSFlexGrid1.Rows - 1
    If Trim(MSFlexGrid1.TextMatrix(i, 1)) <> "" Then
    objeto.RANGE("A" + Trim(Str(NUM))).Value = Trim((MSFlexGrid1.TextMatrix(i, 1)))
    objeto.RANGE("B" + Trim(Str(NUM))).Value = Trim((MSFlexGrid1.TextMatrix(i, 2)))
    objeto.RANGE("C" + Trim(Str(NUM))).Value = Trim((MSFlexGrid1.TextMatrix(i, 3)))
    objeto.RANGE("D" + Trim(Str(NUM))).Value = Trim((MSFlexGrid1.TextMatrix(i, 4)))
    objeto.RANGE("E" + Trim(Str(NUM))).Value = Trim((MSFlexGrid1.TextMatrix(i, 5)))
    objeto.RANGE("E" + Trim(Str(NUM))).Value = Trim((MSFlexGrid1.TextMatrix(i, 5)))
    '**********INSETAR FILA EXCEL**********
    objeto.RANGE(LTrim(Str(NUM + 1)) & ":" & LTrim(Str(NUM + 1))).Select
    objeto.SELECTION.Insert
    '*******'
    NUM = NUM + 1
    Else
    i = MSFlexGrid1.Rows - 1
    End If
c = 1
objeto.RANGE("E" + Trim(Str(20 + c))).Value = Text11.Text
objeto.RANGE("E" + Trim(Str(21 + c))).Value = Text12.Text
objeto.RANGE("E" + Trim(Str(23 + c))).Value = Text14.Text
objeto.RANGE("B" + Trim(Str(3 + c))).Value = Text1.Text
objeto.RANGE("B" + Trim(Str(4 + c))).Value = Text2.Text
objeto.RANGE("B" + Trim(Str(5 + c))).Value = Text3.Text
objeto.RANGE("B" + Trim(Str(6 + c))).Value = Text4.Text
objeto.RANGE("D" + Trim(Str(3 + c))).Value = Text1.Text
objeto.RANGE("D" + Trim(Str(4 + c))).Value = Text1.Text
c = c + 1
Next i
Set objeto = Nothing
End Function

Private Sub excel_Click()
Call Facturaexcel
End Sub

Function Grid()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 80
MSFlexGrid1.FormatString = "#|< Codigo |< Descripcion |< Precio Unitario |< Cantidad |< Total"
MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 3500
MSFlexGrid1.ColWidth(3) = 1500
MSFlexGrid1.ColWidth(4) = 1500
MSFlexGrid1.ColWidth(5) = 1000
End Function
Private Sub añadir_Click()
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + c, 1) = Text7.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + c, 2) = Text8.Text
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + c, 3) = CDbl(Text9.Text)
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + c, 4) = CDbl(Text10.Text)
MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + c, 5) = ((Text9.Text) * (Text10.Text))
subtotal = (Val(Text9.Text) * (Val(Text10))) + valor
Text11.Text = subtotal
valor = Val(Text11.Text)
iva = Format(Val(valor) * 14 / 100, "0.00")
Text12.Text = iva
total = Format(Val(valor + Val(Text12)), "0.00")
Text14.Text = total
Text7.Text = " "
Text8.Text = " "
Text9.Text = " "
Text10.Text = " "
Text11.SetFocus
c = c + 1
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "Cerveza CLUB" Then
Text7.Text = "001"
Text8.Text = Combo1.Text
Text9.Text = "1,75"
End If
If Combo1.Text = "Cerveza Pilsener" Then
Text7.Text = "002"
Text8.Text = Combo1.Text
Text9.Text = "1,50"
End If
If Combo1.Text = "Cerveza Pilsener Light" Then
Text7.Text = "003"
Text8.Text = Combo1.Text
Text9.Text = "1,75"
End If
If Combo1.Text = "Cerveza Cristal" Then
Text7.Text = "004"
Text8.Text = Combo1.Text
Text9.Text = "2,75"
End If
If Combo1.Text = "Cerveza CLUB ROJA" Then
Text7.Text = "005"
Text8.Text = Combo1.Text
Text9.Text = "2,10"
End If
If Combo1.Text = "Cerveza CLUB NEGRA" Then
Text7.Text = "006"
Text8.Text = Combo1.Text
Text9.Text = "2,15"
End If
If Combo1.Text = "Cerveza Miller" Then
Text7.Text = "007"
Text8.Text = Combo1.Text
Text9.Text = "2,25"
End If
If Combo1.Text = "Cerveza Brahama" Then
Text7.Text = "008"
Text8.Text = Combo1.Text
Text9.Text = "2,15"
End If
If Combo1.Text = "Jhony Walker Red" Then
Text7.Text = "009"
Text8.Text = Combo1.Text
Text9.Text = "10,75"
End If
If Combo1.Text = "Jhony Walker Black" Then
Text7.Text = "010"
Text8.Text = Combo1.Text
Text9.Text = "12,8"
End If
If Combo1.Text = "Jhony Walker Blue" Then
Text7.Text = "011"
Text8.Text = Combo1.Text
Text9.Text = "103,75"
End If
If Combo1.Text = "Absolut Vodka" Then
Text7.Text = "012"
Text8.Text = Combo1.Text
Text9.Text = "10,5"
End If
If Combo1.Text = "Vodka Ruskaya" Then
Text7.Text = "013"
Text8.Text = Combo1.Text
Text9.Text = "11,10"
End If
If Combo1.Text = "Ron 100 Fuegos" Then
Text7.Text = "014"
Text8.Text = Combo1.Text
Text9.Text = "11,15"
End If
If Combo1.Text = "Ron del Rio" Then
Text7.Text = "015"
Text8.Text = Combo1.Text
Text9.Text = "5,25"
End If
If Combo1.Text = "Cristal Seco" Then
Text7.Text = "016"
Text8.Text = Combo1.Text
Text9.Text = "4,15"
End If
If Combo1.Text = "Zhumir Seco" Then
Text7.Text = "017"
Text8.Text = Combo1.Text
Text9.Text = "8,15"
End If
If Combo1.Text = "Zhumir Durazno" Then
Text7.Text = "018"
Text8.Text = Combo1.Text
Text9.Text = "7,15"
End If
If Combo1.Text = "Zhumir Tropical" Then
Text7.Text = "019"
Text8.Text = Combo1.Text
Text9.Text = "9,55"
End If
If Combo1.Text = "Tequila Cuervo" Then
Text7.Text = "020"
Text8.Text = Combo1.Text
Text9.Text = "17,15"
End If
If Combo1.Text = "Tequila El Charro" Then
Text7.Text = "021"
Text8.Text = Combo1.Text
Text9.Text = "18, 55"
End If
End Sub


Private Sub Deshacer_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Nuevo.Enabled = True
Eliminar.Enabled = True
Modificar.Enabled = True
Guardar2.Enabled = False
Deshacer.Enabled = False
End Sub

Private Sub Eliminar_Click()
Adodc2.Recordset.Delete
m = MsgBox("Registro Eliminado", vbInformation)
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Nuevo.Enabled = False
Eliminar.Enabled = False
Modificar.Enabled = False
Guardar2.Enabled = True
Deshacer.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
End Sub

Private Sub Form_Load()
Call Grid
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
End Sub

Private Sub guardar_Click()
Call guardarfactura
End Sub
Function guardarfactura()
Dim d As Integer, i As Integer
d = 1
For i = 1 To MSFlexGrid1.Rows - 1
 If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 1) <> " " Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.AddNew
Adodc1.Recordset!Codigo_Botella = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 1)
Adodc1.Recordset!Descripcion = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 2)
Adodc1.Recordset!Precio_Unit = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 3)
Adodc1.Recordset!Cantidad = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 4)
Adodc1.Recordset!Total_Pagar = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + d, 5)
Adodc1.Recordset.Update
 d = d + 1
 Else
  i = MSFlexGrid1.Rows - 1
 End If
Next i
End Function

Private Sub Guardar2_Click()
If Text1.Text <> "" Then
Adodc2.Recordset.Update
Nuevo.Enabled = True
Eliminar.Enabled = True
Modificar.Enabled = True
Guardar2.Enabled = False
Deshacer.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
End If
End Sub

Private Sub imprimirl_Click()
'Call Facturaexcel
Form3.PrintForm
End Sub

Private Sub Modificar_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Nuevo.Enabled = False
Eliminar.Enabled = False
Modificar.Enabled = False
Guardar2.Enabled = True
Deshacer.Enabled = True
End Sub

Private Sub Nuevo_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.AddNew
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Nuevo.Enabled = False
Eliminar.Enabled = False
Modificar.Enabled = False
Guardar2.Enabled = True
Deshacer.Enabled = True
End Sub

Private Sub regresar_Click()
Form2.Show
Unload Me
End Sub
