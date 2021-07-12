VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14580
   LinkTopic       =   "Form7"
   ScaleHeight     =   8505
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid lista 
      Bindings        =   "Form7.frx":0000
      Height          =   3135
      Left            =   600
      TabIndex        =   9
      Top             =   4320
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   5530
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   10920
      Top             =   360
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      Connect         =   $"Form7.frx":0015
      OLEDBString     =   $"Form7.frx":00A1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Factura"
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
   Begin VB.Frame Frame1 
      Caption         =   "Forma de búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   13575
      Begin VB.CommandButton Command2 
         Caption         =   "Hacer el Reporte"
         Height          =   375
         Left            =   8760
         TabIndex        =   11
         Top             =   1440
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Mostrar Resultados "
         Height          =   375
         Left            =   8760
         TabIndex        =   10
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   3135
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Anuladas"
         Height          =   195
         Left            =   6000
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Número de ventas"
         Height          =   195
         Left            =   10920
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fecha "
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Segundo Dato"
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Primer Dato"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   3135
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Búsqueda de Facturas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bus As String

Private Sub Command1_Click()
With Rsfactura
b = "#" & Text2.Text & "#"
'a = "#" & Text1.Text & "#"
    If Rsfactura.State = 1 Then Rsfactura.Close
        If Option1.Value = True Then bus = "#" & Text1.Text & "#": .Open "select * From Factura Where ((fecha)>=" & a & ") AND ((fecha)<=" & b & ")", Base, adOpenStatic, adLockBatchOptimistic
        If Option2.Value = True Then bus = "'" & Text1.Text & "'": .Open "select * From Factura Where ((Total)>=" & bus & ")", Base, adOpenStatic, adLockBatchOptimistic
        DataReport2.Show
        DataField = "vacio"
        If Option3.Value = True Then .Open "select * From Factura Where ((Anuladas)= true )", Base, adOpenStatic, adLockBatchOptimistic
            If .EOF Or .BOF Then Exit Sub
End With
'Set lista.DataSource = Rsfactura
'Set DataReport2.DataSource = Rsfactura
End Sub

Private Sub Command2_Click()
With DataReport2
    Set lista.DataSource = Rsfactura
    DataReport2.Show
End With
End Sub
