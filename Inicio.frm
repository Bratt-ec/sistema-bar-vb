VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Inicio"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6975
   Icon            =   "Inicio.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   4170
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Height          =   3015
      Left            =   360
      Picture         =   "Inicio.frx":9ED32
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Height          =   3015
      Left            =   3720
      Picture         =   "Inicio.frx":A234D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
Form1.Show
Unload Me
End Sub


Private Sub Command1_Click()
End
End Sub
