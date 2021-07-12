VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Inventario"
   ClientHeight    =   8430
   ClientLeft      =   4995
   ClientTop       =   2670
   ClientWidth     =   15030
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   8430
   ScaleWidth      =   15030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBuscar 
      Height          =   405
      Left            =   7560
      TabIndex        =   5
      Top             =   2880
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form4.frx":CE8B
      Left            =   4800
      List            =   "Form4.frx":CE9B
      TabIndex        =   4
      Top             =   3000
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
      Left            =   12480
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":CEC3
      Height          =   2535
      Left            =   3240
      TabIndex        =   1
      Top             =   4200
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
            LCID            =   12298
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
            LCID            =   12298
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
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12840
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11400
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar por: "
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6240
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim buscar As String



Private Sub Form_Load()
    Adodc1.Visible = False
    formatodatagrid
    
    'main
    Productos
    Adodc1.CursorLocation = adUseClient
    
    'Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Karen\Desktop\Papeleria\Proyecto\Base_de_Datos.mdb;Persist Security Info=False"
    Adodc1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\PAZ\Desktop\repositorio\Proyecto\Base_de_Datos.mdb;Persist Security Info=False"
    
    'Conectar el adodc1 con la tabla
    Adodc1.RecordSource = "Select *from Productos"
    'Adodc1.Refresh 'Actualizar los datos del adodc1
    Set DataGrid1.DataSource = RsProductos
End Sub

Private Sub R_Click()
    Form5.Show
    Me.Hide
End Sub

Sub formatodatagrid()
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
    buscar = txtBuscar.Text & "%"
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
            
    RsProductos.Open "Select * from Productos Where Código like '%" & buscar & "'", Base
End Sub

Sub BuscarNombre()
    If RsProductos.State = 1 Then RsProductos.Close
    RsProductos.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RsProductos.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            
    RsProductos.Open "Select * from Productos Where Nombre like '%" & buscar & "'", Base
End Sub

Sub BuscarColor()
    If RsProductos.State = 1 Then RsProductos.Close
    RsProductos.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RsProductos.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            
    RsProductos.Open "Select * from Productos Where Color like '%" & buscar & "'", Base
End Sub

Sub BuscarId_Producto()
    If RsProductos.State = 1 Then RsProductos.Close
    RsProductos.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RsProductos.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            
    RsProductos.Open "Select * from Productos Where Id_producto like '%" & buscar & "'", Base
End Sub
