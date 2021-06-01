VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Menú"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13950
   LinkTopic       =   "Form5"
   ScaleHeight     =   7980
   ScaleWidth      =   13950
   StartUpPosition =   1  'CenterOwner
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
      Left            =   6120
      TabIndex        =   5
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
      Left            =   7680
      TabIndex        =   4
      Top             =   4920
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
      Left            =   11040
      TabIndex        =   3
      Top             =   4920
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
      TabIndex        =   2
      Top             =   4920
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
      Left            =   840
      TabIndex        =   1
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Image Image4 
      Height          =   1920
      Left            =   11040
      Picture         =   "Form5.frx":0000
      Top             =   2640
      Width           =   1920
   End
   Begin VB.Image Image3 
      Height          =   1920
      Left            =   7800
      Picture         =   "Form5.frx":1542
      Top             =   2640
      Width           =   1920
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   4320
      Picture         =   "Form5.frx":2C14
      Top             =   2640
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   840
      Picture         =   "Form5.frx":3E82
      Top             =   2640
      Width           =   1920
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Papeleria Virgen del Cisne"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   0
      Top             =   360
      Width           =   7935
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

Private Sub cmdinventario_Click()
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

