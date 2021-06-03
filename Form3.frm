VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14025
   BeginProperty Font 
      Name            =   "Myanmar Text"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Left            =   120
      TabIndex        =   30
      Top             =   9480
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11880
      TabIndex        =   28
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   25
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11280
      TabIndex        =   14
      Top             =   9840
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11280
      TabIndex        =   13
      Top             =   9000
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8760
      TabIndex        =   12
      Top             =   9000
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   8640
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Lista 
      Height          =   2895
      Left            =   960
      TabIndex        =   10
      Top             =   5520
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   5106
      _Version        =   393216
      Rows            =   10
      Cols            =   5
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Finalizar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   8640
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      TabIndex        =   6
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      TabIndex        =   4
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      TabIndex        =   2
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Label Label17 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   29
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   10560
      TabIndex        =   27
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label15 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Top             =   4680
      Width           =   5415
   End
   Begin VB.Label Label14 
      Caption         =   "Correo"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label13 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   3600
      Width           =   5415
   End
   Begin VB.Label Label12 
      Caption         =   "Celular"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label11 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   2400
      Width           =   5415
   End
   Begin VB.Label Label10 
      Caption         =   "Datos Del Producto"
      Height          =   375
      Left            =   8040
      TabIndex        =   19
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label Label9 
      Caption         =   "Datos Del Cliente"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label Label8 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   17
      Top             =   9600
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Iva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   16
      Top             =   8760
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   15
      Top             =   8760
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Precio "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Factura "
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
Form5.Show
Me.Hide
End Sub

Private Sub Command4_Click()
tot = tot - a
Text5.Text = tot
fila = fila - 1
Lista.Col = 1
Lista.Row = fila
Lista.Text = ""
Lista.Col = 2
Lista.Row = fila
Lista.Text = ""
Lista.Col = 3
Lista.Row = fila
Lista.Text = ""
Lista.Col = 4
Lista.Row = fila
Lista.Text = ""
End Sub

Private Sub Form_Load()
Lista.ColWidth(0) = 10
Lista.ColWidth(1) = 3000
Lista.ColAlignment(1) = 5
Lista.Col = 1
Lista.Row = 0
Lista.Text = "Producto"
Lista.ColWidth(2) = 3000
Lista.ColAlignment(2) = 5
Lista.Col = 2
Lista.Row = 0
Lista.Text = "Cantidad"
Lista.ColWidth(3) = 3000
Lista.ColAlignment(3) = 5
Lista.Col = 3
Lista.Row = 0
Lista.Text = "Precio"
Lista.ColWidth(4) = 3000
Lista.ColAlignment(4) = 5
Lista.Col = 4
Lista.Row = 0
Lista.Text = "Total unico"
fila = 1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Lista.Col = 1
    Lista.Row = fila
    Lista.Text = Text2.Text
    Lista.Col = 2
    Lista.Row = fila
    Lista.Text = Text3.Text
    Lista.Col = 3
    Lista.Row = fila
    Lista.Text = Text1.Text
    a = Val(Text2.Text) * Val(Text3.Text)
    Lista.Col = 4
    Lista.Row = fila
    Lista.Text = a
    tot = tot + a
    Text5.Text = tot
    Text6.Text = tot * 0.12
    Text7 = tot + Val(Text6.Text)
    fila = fila + 1
    Text2.Text = ""
    Text3.Text = ""
    Text1.Text = ""
    Text2.SetFocus
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label17.Caption = Text10.Text
    Text10.Text = ""
    Text2.SetFocus
End If
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label11.Caption = Text4.Text
    Text4.Text = ""
    Text8.SetFocus
End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label13.Caption = Text8.Text
    Text8.Text = ""
    Text9.SetFocus
End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label15.Caption = Text9.Text
    Text9.Text = ""
    Text10.SetFocus
End If
End Sub
