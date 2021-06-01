VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Factura"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765
   LinkTopic       =   "Form3"
   ScaleHeight     =   9600
   ScaleWidth      =   9765
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
      Left            =   240
      TabIndex        =   14
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   8880
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2655
      Left            =   240
      TabIndex        =   12
      Top             =   3960
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   10
      Cols            =   5
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Añadir"
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pagar"
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Finalizar "
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   8880
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Precio "
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Describción"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Factura "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub R_Click()
Form5.Show
Me.Hide
End Sub
