VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Inventario"
   ClientHeight    =   6585
   ClientLeft      =   4995
   ClientTop       =   2670
   ClientWidth     =   10590
   LinkTopic       =   "Form4"
   ScaleHeight     =   6585
   ScaleWidth      =   10590
   Begin VB.TextBox txtBuscar 
      Height          =   405
      Left            =   4320
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form4.frx":0000
      Left            =   1920
      List            =   "Form4.frx":0010
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton RM 
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
      Height          =   735
      Left            =   7800
      TabIndex        =   2
      Top             =   4920
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":0038
      Height          =   2535
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4471
      _Version        =   393216
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   600
      Top             =   5040
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   $"Form4.frx":004D
      OLEDBString     =   $"Form4.frx":00D5
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Productos"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar por: "
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Buscar As String

Private Sub Form_Load()
    Adodc1.Visible = False
    FormatoDataGrid
    
    main
    Productos
    Adodc1.CursorLocation = adUseClient
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Karen\Desktop\Papeleria\Proyecto\Base_de_Datos.mdb;Persist Security Info=False"
    'Conectar el adodc1 con la tabla
    Adodc1.RecordSource = "Select *from Productos"
    'Adodc1.Refresh 'Actualizar los datos del adodc1
    Set DataGrid1.DataSource = RsProductos
End Sub

Private Sub R_Click()
    Form5.Show
    Me.Hide
End Sub

Sub FormatoDataGrid()
    DataGrid1.Columns(0).Width = 1000
    DataGrid1.Columns(1).Width = 2000
    DataGrid1.Columns(2).Width = 1500
    DataGrid1.Columns(3).Width = 1500
    DataGrid1.Columns(4).Width = 1500
    DataGrid1.Columns(5).Width = 1500
End Sub

Private Sub RM_Click()
    Form5.Show
    Me.Hide
End Sub

Private Sub txtBuscar_Change()
    Buscar = txtBuscar.Text & "%"
    If Combo1.Text = "Código" Then BuscarCodigo
    If Combo1.Text = "Nombre" Then BuscarNombre
    If Combo1.Text = "Color" Then BuscarColor
    If Combo1.Text = "Id_Producto" Then BuscarId_Producto
    Set DataGrid1.DataSource = RsProductos
End Sub

Sub BuscarCodigo()
    If RsProductos.State = 1 Then RsProductos.Close
    RsProductos.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RsProductos.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            
    RsProductos.Open "Select * from Productos Where Código like '%" & Buscar & "'", Base
End Sub

Sub BuscarNombre()
    If RsProductos.State = 1 Then RsProductos.Close
    RsProductos.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RsProductos.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            
    RsProductos.Open "Select * from Productos Where Nombre like '%" & Buscar & "'", Base
End Sub

Sub BuscarColor()
    If RsProductos.State = 1 Then RsProductos.Close
    RsProductos.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RsProductos.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            
    RsProductos.Open "Select * from Productos Where Color like '%" & Buscar & "'", Base
End Sub

Sub BuscarId_Producto()
    If RsProductos.State = 1 Then RsProductos.Close
    RsProductos.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RsProductos.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            
    RsProductos.Open "Select * from Productos Where Id_producto like '%" & Buscar & "'", Base
End Sub
