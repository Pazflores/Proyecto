VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   10215
   ClientLeft      =   3210
   ClientTop       =   885
   ClientWidth     =   14535
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
   ScaleWidth      =   14535
   Begin MSFlexGridLib.MSFlexGrid Lista 
      Height          =   2655
      Left            =   360
      TabIndex        =   35
      Top             =   6360
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   20
      Cols            =   5
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Regresar al Menú"
      Height          =   495
      Left            =   0
      TabIndex        =   34
      Top             =   9480
      Width           =   1935
   End
   Begin VB.TextBox Text12 
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
      TabIndex        =   31
      Top             =   5280
      Width           =   3975
   End
   Begin VB.TextBox Text11 
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
      TabIndex        =   28
      Top             =   960
      Width           =   3975
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
      Left            =   11640
      TabIndex        =   26
      Top             =   1440
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
      TabIndex        =   23
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
      TabIndex        =   20
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
      Left            =   10920
      TabIndex        =   12
      Top             =   9600
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
      Left            =   8400
      TabIndex        =   11
      Top             =   9600
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
      Left            =   5880
      TabIndex        =   10
      Top             =   9600
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Borrar ultima fila"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   9480
      Width           =   2895
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
      Top             =   2040
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
      Top             =   4080
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
      Top             =   3480
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
      Top             =   4680
      Width           =   3975
   End
   Begin VB.Label Label21 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label20 
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
      TabIndex        =   32
      Top             =   5760
      Width           =   5415
   End
   Begin VB.Label Label19 
      Caption         =   "RUC"
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label18 
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
      TabIndex        =   29
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label17 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   27
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   10800
      TabIndex        =   25
      Top             =   1800
      Width           =   735
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
      TabIndex        =   24
      Top             =   4680
      Width           =   5415
   End
   Begin VB.Label Label14 
      Caption         =   "Correo"
      Height          =   255
      Left            =   360
      TabIndex        =   22
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
      TabIndex        =   21
      Top             =   3600
      Width           =   5415
   End
   Begin VB.Label Label12 
      Caption         =   "Celular"
      Height          =   255
      Left            =   360
      TabIndex        =   19
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
      TabIndex        =   18
      Top             =   2520
      Width           =   5415
   End
   Begin VB.Label Label10 
      Caption         =   "Datos Del Producto"
      Height          =   375
      Left            =   8040
      TabIndex        =   17
      Top             =   2400
      Width           =   5415
   End
   Begin VB.Label Label9 
      Caption         =   "Datos Del Cliente"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   240
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
      Left            =   10920
      TabIndex        =   15
      Top             =   9360
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
      Left            =   8400
      TabIndex        =   14
      Top             =   9360
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
      Left            =   5880
      TabIndex        =   13
      Top             =   9360
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
      Top             =   2040
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
      Top             =   4080
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
      Top             =   3480
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
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Virgen de Cisne"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7080
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
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
Lista.Text = "Precio"
Lista.ColWidth(3) = 3000
Lista.ColAlignment(3) = 5
Lista.Col = 3
Lista.Row = 0
Lista.Text = "Cantidad"
Lista.ColWidth(4) = 3000
Lista.ColAlignment(4) = 5
Lista.Col = 4
Lista.Row = 0
Lista.Text = "Total Unico"
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
    a = Val(Text3.Text) * Val(Text1.Text)
    Lista.Col = 4
    Lista.Row = fila
    Lista.Text = a
    tot = tot + a
    Text5.Text = tot
    fila = fila + 1
    Text2.Text = ""
    Text3.Text = ""
    Text1.Text = ""
    Text2.SetFocus
    Text6 = tot * 0.12
    Text7.Text = tot + Val(Text6.Text)
End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label17.Caption = Text10.Text
    Text10.Text = ""
    Text2.SetFocus
End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label18.Caption = Text11.Text
    Text11.Text = ""
    Text4.SetFocus
End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label20.Caption = Text12.Text
    Text12.Text = ""
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
    Text12.SetFocus
End If
End Sub
