VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19005
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
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   9135
   ScaleWidth      =   19005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Crear Factura"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   48
      Top             =   7320
      Width           =   3015
   End
   Begin VB.PictureBox Adodc3 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   17520
      ScaleHeight     =   270
      ScaleWidth      =   1140
      TabIndex        =   50
      Top             =   6360
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ComboBox Combo2 
      Height          =   420
      Left            =   8400
      TabIndex        =   47
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text14 
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
      Left            =   7800
      TabIndex        =   45
      Top             =   2880
      Width           =   2775
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
      Left            =   1560
      TabIndex        =   44
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar Producto"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   43
      Top             =   5280
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      Height          =   420
      ItemData        =   "Form3.frx":DC12
      Left            =   2760
      List            =   "Form3.frx":DC14
      TabIndex        =   42
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text13 
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
      Left            =   2280
      TabIndex        =   40
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Buscar Cliente"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11640
      TabIndex        =   38
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   18000
      Top             =   7080
   End
   Begin VB.PictureBox Adodc2 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   17520
      ScaleHeight     =   315
      ScaleWidth      =   1140
      TabIndex        =   51
      Top             =   5520
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   17520
      ScaleHeight     =   315
      ScaleWidth      =   1140
      TabIndex        =   52
      Top             =   4680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Regresar al Men?"
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
      Left            =   14280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   34
      Top             =   360
      Width           =   2055
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
      Left            =   12600
      TabIndex        =   31
      Top             =   3600
      Width           =   2895
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
      Left            =   7680
      TabIndex        =   28
      Top             =   3360
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid Lista 
      Height          =   2535
      Left            =   4440
      TabIndex        =   27
      Top             =   5280
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   20
      Cols            =   5
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
      Left            =   12600
      TabIndex        =   23
      Top             =   2760
      Width           =   2895
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
      Left            =   12600
      TabIndex        =   20
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
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
      Left            =   12000
      TabIndex        =   12
      Top             =   8520
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
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
      Left            =   8640
      TabIndex        =   11
      Top             =   8520
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      DataField       =   "PrecioUnitario"
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
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
      Left            =   5640
      TabIndex        =   10
      Top             =   8520
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      Top             =   6720
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
      Left            =   7680
      TabIndex        =   8
      Top             =   4320
      Width           =   2895
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
      Left            =   2280
      TabIndex        =   6
      Top             =   4080
      Width           =   2295
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
      Left            =   2280
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc3"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   13440
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label Label28 
      Caption         =   "Label28"
      Height          =   735
      Left            =   17160
      TabIndex        =   49
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar Dato:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   46
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   41
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de busqueda:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   39
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
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
      Left            =   8760
      TabIndex        =   37
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   36
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   35
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Direcci?n:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11160
      TabIndex        =   33
      Top             =   3600
      Width           =   1215
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
      Left            =   11160
      TabIndex        =   32
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "C?dula:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   30
      Top             =   3360
      Width           =   1215
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
      Left            =   6360
      TabIndex        =   29
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   26
      Top             =   1560
      Width           =   1395
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   6360
      TabIndex        =   25
      Top             =   1560
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
      Left            =   11160
      TabIndex        =   24
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Correo:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11160
      TabIndex        =   22
      Top             =   2760
      Width           =   1215
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
      Left            =   11160
      TabIndex        =   21
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Celular:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11160
      TabIndex        =   19
      Top             =   1800
      Width           =   1215
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
      Left            =   6360
      TabIndex        =   18
      Top             =   4680
      Width           =   4215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Del Producto:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Del Cliente:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   15
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Iva:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   14
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio "
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FACTURA"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   6840
      TabIndex        =   0
      Top             =   360
      Width           =   3765
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

Private Sub Command2_Click()
Label28.Caption = Text3.Text * Text1.Text
End Sub

Private Sub Command3_Click()
Productos
If Text10.Text = "" Then
    MsgBox "Llenar la opcion buscar", vbCritical, "Llenar casilla "
ElseIf Combo1.Text = "" Then
    MsgBox "Porfavor seleccione una de las opciones para continuar", vbCritical, "Elija una opci?n"
Else
    RsProductos.MoveFirst
    RsProductos.Find "C?digo='" & Text10.Text & "'"
    If RsProductos.EOF Then
    MsgBox "No se encontro"
    End If
    
    If RsProductos.EOF = False And RsProductos.BOF = False Then
        Text2.Text = RsProductos.Fields(1).Value
        Text3.Text = RsProductos.Fields(2).Value
        Text13.Text = RsProductos.Fields(0).Value
        Text10.Text = ""
        Text10.SetFocus
    Else
        MsgBox "No se ha podido encontrar el archivo deseado", vbCritical, "Archivo no encontrado"
        RsProductos.MoveFirst
    End If
End If
End Sub

Private Sub Command4_Click()
If lista.Row <= 0 Then
MsgBox "Debe Seleccionar una fila"
    ElseIf lista.Row = 1 Then
    lista.Clear
Else
lista.RemoveItem (lista.Row)
Text7.Text = ""
Text6.Text = ""
Text5.Text = ""
End If
End Sub

Private Sub Command5_Click()
Cliente
If Text14.Text = "" Then
    MsgBox "Llenar la opcion buscar", vbCritical, "Llenar casilla "
ElseIf Combo2.Text = "" Then
    MsgBox "Porfavor seleccione una de las opciones para continuar", vbCritical, "Elija una opci?n"
Else
    RsCliente.MoveFirst
    RsCliente.Find "C?dula='" & Text14.Text & "'"
    If RsCliente.EOF Then
    MsgBox "No se encontro"
    End If
    
    If RsCliente.EOF = False And RsCliente.BOF = False Then
        Text11.Text = RsCliente.Fields(0).Value
        Text4.Text = RsCliente.Fields(1).Value
        Text8.Text = RsCliente.Fields(3).Value
        Text9.Text = RsCliente.Fields(4).Value
        Text12.Text = RsCliente.Fields(2).Value
        Text14.Text = ""
        Text14.SetFocus
       
    Else
        MsgBox "No se ha podido encontrar el archivo deseado", vbCritical, "Archivo no encontrado"
        RsCliente.MoveFirst
    End If
End If
Label18.Caption = Text11.Text
Label11.Caption = Text4.Text
Label13.Caption = Text8.Text
Label15.Caption = Text9.Text
Label20.Caption = Text12.Text
Text11.Text = ""
Text4.Text = ""
Text8.Text = ""
Text9.Text = ""
Text12.Text = ""
End Sub

Private Sub Command6_Click()
With DataReport1
        .Sections("Secci?n2").Controls("Etiqueta1").Caption = Label18.Caption
        .Sections("Secci?n2").Controls("Etiqueta2").Caption = Label11.Caption
        .Sections("Secci?n2").Controls("Etiqueta3").Caption = Label13.Caption
        .Sections("Secci?n2").Controls("Etiqueta4").Caption = Label15.Caption
        .Sections("Secci?n2").Controls("Etiqueta5").Caption = Label20.Caption
        .Sections("Secci?n4").Controls("Etiqueta18").Caption = Label17.Caption
        .Sections("Secci?n4").Controls("Etiqueta17").Caption = Label22.Caption
        .Sections("Secci?n3").Controls("Etiqueta19").Caption = Text7.Text
        .Sections("Secci?n3").Controls("Etiqueta22").Caption = Text6.Text
        .Sections("Secci?n3").Controls("Etiqueta23").Caption = Text5.Text
    End With

Temporal
Set DataReport1.DataSource = Rstemporal
DataReport1.Show
End Sub

Private Sub Form_Load()
Combo2.AddItem ("C?dula")
Combo1.AddItem ("id")
Text1.Text = ""
Text5.Text = ""
lista.ColWidth(0) = 10
lista.ColWidth(1) = 3000
lista.ColAlignment(1) = 5
lista.Col = 1
lista.Row = 0
lista.Text = "Producto"
lista.ColWidth(2) = 3000
lista.ColAlignment(2) = 5
lista.Col = 2
lista.Row = 0
lista.Text = "Precio"
lista.ColWidth(3) = 3000
lista.ColAlignment(3) = 5
lista.Col = 3
lista.Row = 0
lista.Text = "Cantidad"
lista.ColWidth(4) = 3000
lista.ColAlignment(4) = 5
lista.Col = 4
lista.Row = 0
lista.Text = "Total Unico"
fila = 1
Temporal
    With Rstemporal
        For i = 1 To .RecordCount
            .Delete
            .MoveNext
        Next i
    End With
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label28.Caption = Text3.Text * Text1.Text
    lista.Col = 1
    lista.Row = fila
    lista.Text = Text2.Text
    lista.Col = 2
    lista.Row = fila
    lista.Text = Text3.Text
    lista.Col = 3
    lista.Row = fila
    lista.Text = Text1.Text
    lista.Col = 4
    lista.Row = fila
    lista.Text = Label28.Caption
    tot = tot + Label28.Caption
    Text5.Text = tot
    fila = fila + 1
    DetalleFactura
    With Rsdefac
        .AddNew
        !Cantidad = Text1.Text
        !Descripci?n = Text2.Text
        !PrecioUnitario = Text5.Text
        !Total = Text7.Text
        !C?digo_producto = Text13.Text
        .UpdateBatch
    End With
    Temporal
    With Rstemporal
        .AddNew
        !Cantidad = Text1.Text
        !Descripci?n = Text2.Text
        !PrecioUnitario = Text5.Text
        !Total = Text7.Text
        !C?digo_producto = Text13.Text
        .UpdateBatch
    End With
    Text2.Text = ""
    Text3.Text = ""
    Text1.Text = ""
    Text13.Text = ""
    Text2.SetFocus
    Text6 = tot * 0.12
    Text7.Text = tot + Text6.Text
End If
End Sub


Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text13.SetFocus
End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
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

Private Sub Text80_Change()
   
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label15.Caption = Text9.Text
    Text9.Text = ""
    Text12.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Label17.Caption = Date
Label22.Caption = Time
End Sub
