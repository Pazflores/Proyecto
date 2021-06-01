VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Inventario"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12870
   LinkTopic       =   "Form4"
   ScaleHeight     =   7935
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton R 
      Caption         =   "Regresar al menú"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10800
      TabIndex        =   3
      Top             =   7080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub R_Click()
Form5.Show
Me.Hide
End Sub
