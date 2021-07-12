VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Registro de Clientes "
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15525
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   15525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "Buscar"
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
      Left            =   11760
      TabIndex        =   19
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton R 
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
      Left            =   12600
      TabIndex        =   18
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtapellido 
      DataField       =   "Apellido"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   16
      Top             =   3960
      Width           =   2895
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   1755
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txttelefono 
      DataField       =   "Teléfono"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   14
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox txtdireccion 
      DataField       =   "Dirección"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   5160
      Width           =   3015
   End
   Begin VB.TextBox txtcorreo 
      DataField       =   "Correo"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox txtnombre 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   7
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtcedula 
      DataField       =   "Cédula"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton cmdanterior 
      Caption         =   "Anterior"
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
      Left            =   4680
      TabIndex        =   5
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdsiguiente 
      Caption         =   "Siguiente"
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
      TabIndex        =   4
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "Eliminar"
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
      Left            =   8640
      TabIndex        =   3
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Guardar"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdnuevo 
      Caption         =   "Nuevo"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido:"
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
      Left            =   9240
      TabIndex        =   17
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
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
      Left            =   9240
      TabIndex        =   15
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección :"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Correo:"
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
      Left            =   3240
      TabIndex        =   12
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula:"
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
      Left            =   3240
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
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
      Left            =   9240
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Clientes"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   0
      Top             =   1200
      Width           =   6255
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

Private Sub cmdbuscar_Click()
On Error GoTo salida
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
End If
'Igualar la variable busqueda al input'
Dim Busqueda As String
Busqueda = InputBox("Ingrese el número de cédula que desea Buscar", "Sistema de Registro")
'Realizamos la busqueda usando el metodo find'
Adodc1.Recordset.Find "Cédula='" & Trim(Busqueda) & "'"
'Si encuentra resultados que nos muestre en un msgbox'
If Adodc1.Recordset.EOF Then
MsgBox "Saliendo de busqueda cédula no encontrada", vbCritical, "Sistema de Registro"
Exit Sub
End If
'Y si encontró resultados mostrar la descripción del cliente en un textbox'
txtcedula.Text = Adodc1.Recordset.Fields(0).Value
txtnombre.Text = Adodc1.Recordset.Fields(1).Value
txtapellido.Text = Adodc1.Recordset.Fields(2).Value
txtdireccion.Text = Adodc1.Recordset.Fields(3).Value
txttelefono.Text = Adodc1.Recordset.Fields(4).Value
txtcorreo.Text = Adodc1.Recordset.Fields(5).Value
Exit Sub
salida:
End Sub

Private Sub cmdeliminar_Click()
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveLast
End If
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
End Sub
Private Sub cmdsiguiente_Click()
On Error Resume Next
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MovePrevious
End If
End Sub



Private Sub R_Click()
Form5.Show
Me.Hide
End Sub
