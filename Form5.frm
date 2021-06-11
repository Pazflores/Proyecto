VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Menú"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14745
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   14745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsalir 
      Caption         =   "SALIR"
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
      Left            =   6600
      TabIndex        =   4
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdinventario 
      Caption         =   "INVENTARIO"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   3
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdfactura 
      Caption         =   "FACTURA"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11760
      TabIndex        =   2
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdproductos 
      Caption         =   "PRODUCTOS"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   1
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdclientes 
      Caption         =   "CLIENTES"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MENÚ"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1335
      Left            =   2880
      Top             =   360
      Width           =   9255
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclientes_Click()
Form2.Show
Me.Hide
End Sub

Private Sub cmdfactura_Click()
Form3.Show
Me.Hide
End Sub

Private Sub cmdInventario_Click()
Form4.Show
Me.Hide
End Sub

Private Sub cmdproductos_Click()
Form6.Show
Me.Hide
End Sub

Private Sub cmdsalir_Click()
If MsgBox("Esta seguro que desea cerrar el formulario?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub

