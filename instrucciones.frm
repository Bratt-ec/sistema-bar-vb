VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Instrucciones"
   ClientHeight    =   7800
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6225
   LinkTopic       =   "Form8"
   ScaleHeight     =   7800
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "INSTRUCCIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   $"instrucciones.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7335
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Menu Opciones 
      Caption         =   "Opciones"
      Begin VB.Menu regresar 
         Caption         =   "Regresar"
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Command1.Enabled = False
End Sub

Private Sub regresar_Click()
Form2.Show
Unload Me
End Sub

Private Sub Salir_Click()
Form1.Show
Unload Me
End Sub
