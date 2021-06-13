VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18390
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
   ScaleHeight     =   9255
   ScaleWidth      =   18390
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
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   49
      Top             =   8520
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   16800
      Top             =   6480
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form3.frx":DC12
      OLEDBString     =   $"Form3.frx":DC9E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "DetalleFactura"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox Combo2 
      Height          =   420
      Left            =   8400
      TabIndex        =   48
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
      TabIndex        =   46
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
      TabIndex        =   45
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
      TabIndex        =   44
      Top             =   5280
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      Height          =   420
      ItemData        =   "Form3.frx":DD2A
      Left            =   2760
      List            =   "Form3.frx":DD2C
      TabIndex        =   43
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
      TabIndex        =   41
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
      TabIndex        =   39
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   17280
      Top             =   7200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar Información"
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
      TabIndex        =   36
      Top             =   7920
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   16800
      Top             =   5520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form3.frx":DD2E
      OLEDBString     =   $"Form3.frx":DDB7
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Productos"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   16800
      Top             =   4560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form3.frx":DE40
      OLEDBString     =   $"Form3.frx":DEC9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Cliente"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myanmar Text"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Regresar al Menú"
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
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   34
      Top             =   6240
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
      Left            =   7800
      TabIndex        =   28
      Top             =   3360
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid Lista 
      Height          =   2535
      Left            =   4680
      TabIndex        =   27
      Top             =   5280
      Width           =   11655
      _ExtentX        =   20558
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
      DataField       =   "Total"
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
      Left            =   12000
      TabIndex        =   12
      Top             =   8520
      Width           =   2175
   End
   Begin VB.TextBox Text6 
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
      Left            =   8640
      TabIndex        =   11
      Top             =   8520
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      DataField       =   "PrecioUnitario"
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
      Left            =   5640
      TabIndex        =   10
      Top             =   8520
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Borrar ultima fila"
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
      Top             =   7320
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
      DataField       =   "Cantidad"
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
      TabIndex        =   47
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
      TabIndex        =   42
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
      TabIndex        =   40
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
      TabIndex        =   38
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
      TabIndex        =   37
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
      Caption         =   "Dirección:"
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
      Caption         =   "Cédula:"
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
Adodc3.Recordset.AddNew
Adodc3.Recordset.Update
Adodc3.Refresh
End Sub

Private Sub Command3_Click()
Productos
If Text10.Text = "" Then
    MsgBox "Llenar la opcion buscar", vbCritical, "Llenar casilla "
ElseIf Combo1.Text = "" Then
    MsgBox "Porfavor seleccione una de las opciones para continuar", vbCritical, "Elija una opción"
Else
    RsProductos.MoveFirst
    RsProductos.Find "Código='" & Text10.Text & "'"
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

Private Sub Command5_Click()
Cliente
If Text14.Text = "" Then
    MsgBox "Llenar la opcion buscar", vbCritical, "Llenar casilla "
ElseIf Combo2.Text = "" Then
    MsgBox "Porfavor seleccione una de las opciones para continuar", vbCritical, "Elija una opción"
Else
    RsCliente.MoveFirst
    RsCliente.Find "Cédula='" & Text14.Text & "'"
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
DataReport1.Show
End Sub

Private Sub Form_Load()
Combo2.AddItem ("Cédula")
Combo1.AddItem ("id")
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
    Text13.Text = ""
    Text2.SetFocus
    Text6 = tot * 0.12
    Text7.Text = tot + Val(Text6.Text)
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
