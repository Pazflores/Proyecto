VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdingresar 
      Caption         =   "INGRESAR"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SALIR"
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtusuario 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2040
      Width           =   4455
   End
   Begin VB.TextBox txtcontraseña 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label Label7 
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "BIENVENIDO "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdingresar_Click()
If txtusuario.Text = "papeleria" And txtcontraseña.Text = "virgendelcisne" Then
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

