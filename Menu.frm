VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Menu"
   ClientHeight    =   2730
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9135
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "Menu.frx":9ED32
   ScaleHeight     =   2730
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   1740
      Left            =   1320
      Picture         =   "Menu.frx":1A71D5
      Top             =   360
      Width           =   6000
   End
   Begin VB.Menu Ventas 
      Caption         =   "Vender"
      Begin VB.Menu Factura 
         Caption         =   "Factura"
      End
   End
   Begin VB.Menu Bebidas 
      Caption         =   "Bebidas"
      Begin VB.Menu Inventario 
         Caption         =   "Inventario"
      End
   End
   Begin VB.Menu Pedidos 
      Caption         =   "Pedidos"
      Begin VB.Menu Pedidos2 
         Caption         =   "Pedidos a hacer"
      End
      Begin VB.Menu Pedidos_recibir 
         Caption         =   "Pedidos a recibir"
      End
   End
   Begin VB.Menu Clientes 
      Caption         =   "Clientes"
      Begin VB.Menu Lista_Clientes 
         Caption         =   "Lista de Clientes"
      End
   End
   Begin VB.Menu Opciones 
      Caption         =   "Opciones"
      Begin VB.Menu Instrucciones 
         Caption         =   "Instrucciones"
      End
      Begin VB.Menu Cerrar_Sesion 
         Caption         =   "Cerrar"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cerrar_Sesion_Click()
Form1.Show
Unload Me
End Sub

Private Sub Factura_Click()
Form3.Show
Unload Me
End Sub

Private Sub Instrucciones_Click()
Form8.Show
Unload Me
End Sub

Private Sub Inventario_Click()
Form4.Show
Unload Me
End Sub

Private Sub Lista_Clientes_Click()
Form5.Show
Unload Me
End Sub

Private Sub Pedidos_recibir_Click()
Form7.Show
Unload Me
End Sub

Private Sub Pedidos2_Click()
Form6.Show
Unload Me
End Sub
