VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14310
   LinkTopic       =   "Form2"
   ScaleHeight     =   8025
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtapellido 
      DataField       =   "Apellido"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   16
      Top             =   3360
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   9240
      Top             =   4320
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   $"Form2.frx":0000
      OLEDBString     =   $"Form2.frx":0089
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Cliente"
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
   Begin VB.TextBox txttelefono 
      DataField       =   "Teléfono"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   5880
      Width           =   2535
   End
   Begin VB.TextBox txtdireccion 
      DataField       =   "Dirección"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox txtcorreo 
      DataField       =   "Correo"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox txtnombre 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtcedula 
      DataField       =   "Cédula"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton cmdanterior 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   9600
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdsiguiente 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   9600
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   8400
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdnuevo 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Apellido"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Teléfono"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Dirección "
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Correo"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Cédula"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Registro de Clientes"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdanterior_Click()
On Error Resume Next
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveNext
End If
End Sub
Private Sub cmdeliminar_Click()
On Error GoTo salida
Adodc1.Recordset.Delete
MsgBox "Se eliminaron los dato correctamente", vbInformation, "Sistema de Registro"
Adodc1.Recordset.AddNew
Exit Sub
salida:
MsgBox "Los campos estan vacios, busque los datos a eliminar", vbCritical, "Sistema de Registro"
End Sub

Private Sub cmdguardar_Click()
On Error GoTo salida
Adodc1.Recordset.Update
MsgBox "Se guardaron los datos correctamente, se han corrido al registro anterior", vbInformation, "Sistema de Registro"
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
End If
Exit Sub
salida:
MsgBox "Los campos estan vacios no se pueden guardar hasta llenarlos", vbInformation, "Sistema de Registro"
End Sub

Private Sub cmdnuevo_Click()
On Error GoTo salida
Adodc1.Recordset.AddNew
MsgBox "Llene los campos para ingresar un nuevo registro", vbInformation, "Sistema de Registro"
Exit Sub
salida:
MsgBox "Dar clic dos veces en Nuevo para registrar", vbCritical, "Sistema de Registro"
End Sub
Private Sub cmdsiguiente_Click()
On Error Resume Next
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MovePrevious
End If
End Sub
