VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Loguin"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   14715
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4080
      MaskColor       =   &H8000000F&
      TabIndex        =   5
      Top             =   7200
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
      Left            =   9120
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   7200
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
      Left            =   5760
      TabIndex        =   3
      Top             =   4320
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
      Left            =   5760
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5880
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PAPELERÍA VIRGEN DEL CISNE"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2040
      TabIndex        =   6
      Top             =   600
      Width           =   11295
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1335
      Left            =   1560
      Top             =   360
      Width           =   11895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
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
      Left            =   5760
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   5760
      TabIndex        =   0
      Top             =   5280
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdingresar_Click()
'Contraseña por defecto'
If txtusuario.Text = "Papeleria" And txtcontraseña.Text = "virgendelcisne" Then
  Form5.Show
  Me.Hide
Else
MsgBox "Los datos ingresados no son correctos"
txtcontraseña.Text = ""
txtusuario.Text = ""
txtusuario.SetFocus
End If
End Sub
Private Sub Command1_Click()
'Mensaje informativo pra cerrar el formulario'
If MsgBox("Esta seguro que desea cerrar el formulario?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub
