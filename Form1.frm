VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   11655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdingresar 
      Caption         =   "INGRESAR"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      MaskColor       =   &H8000000F&
      TabIndex        =   6
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   6960
      MaskColor       =   &H8000000F&
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox txtusuario 
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   5880
      TabIndex        =   4
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox txtcontraseña 
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   5880
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4200
      Width           =   4455
   End
   Begin VB.Image Image2 
      Height          =   1680
      Left            =   2040
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1770
      Left            =   2040
      Picture         =   "Form1.frx":3F89
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "BIENVENIDO "
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
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdingresar_Click()
If txtusuario.Text = "Papeleria" And txtcontraseña.Text = "virgendelcisne" Then
  Form2.Show
  Me.Hide
Else
MsgBox "Los datos ingresados no son correctos"
txtcontraseña.Text = ""
txtusuario.Text = ""
txtusuario.SetFocus
End If
End Sub
Private Sub Command1_Click()
If MsgBox("Esta seguro que desea cerrar el formulario?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub

