VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Registro de ProductosRegistro de Productos"
   ClientHeight    =   9225
   ClientLeft      =   3285
   ClientTop       =   1200
   ClientWidth     =   13635
   LinkTopic       =   "Form6"
   ScaleHeight     =   9225
   ScaleWidth      =   13635
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
      Left            =   5760
      TabIndex        =   20
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmdInventario 
      Caption         =   "Inventario"
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
      Left            =   9600
      TabIndex        =   19
      Top             =   8400
      Width           =   1815
   End
   Begin VB.TextBox txtId_Producto 
      DataField       =   "Id_producto"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   9600
      TabIndex        =   17
      Top             =   4920
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11280
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Connect         =   $"Form6.frx":0000
      OLEDBString     =   $"Form6.frx":0089
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
      Left            =   1080
      TabIndex        =   16
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
      Left            =   5520
      TabIndex        =   15
      Top             =   6000
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
      Left            =   9960
      TabIndex        =   14
      Top             =   6120
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
      Left            =   7800
      TabIndex        =   13
      Top             =   7440
      Width           =   1815
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
      Left            =   3360
      TabIndex        =   12
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox txtcolor 
      DataField       =   "Color"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   5040
      Width           =   3855
   End
   Begin VB.TextBox txtstock 
      DataField       =   "Stock"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   9600
      TabIndex        =   9
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox txtnombre 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   3960
      Width           =   3735
   End
   Begin VB.TextBox txtprecio 
      DataField       =   "Precio"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txtcodigo 
      DataField       =   "Código"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   2640
      Width           =   3735
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
      Left            =   11520
      TabIndex        =   1
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Id_producto"
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
      Left            =   7200
      TabIndex        =   18
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   10080
      Picture         =   "Form6.frx":0112
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   960
   End
   Begin VB.Image Image3 
      Height          =   840
      Left            =   2160
      Picture         =   "Form6.frx":2FDC
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   1320
      Picture         =   "Form6.frx":64C7
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1800
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Color"
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
      Left            =   600
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre"
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
      Left            =   600
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Precio"
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
      Left            =   7200
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stock"
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
      Left            =   7200
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Código"
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
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Productos"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   7095
   End
End
Attribute VB_Name = "Form6"
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
Busqueda = InputBox("Ingrese el número de código que desea Buscar", "Sistema de Registro")
'Realizamos la busqueda usando el metodo find'
Adodc1.Recordset.Find "Código='" & Trim(Busqueda) & "'"
'Si encuentra resultados que nos muestre en un msgbox'
If Adodc1.Recordset.EOF Then
MsgBox "Saliendo de busqueda cédula no encontrada", vbCritical, "Sistema de Registro"
Exit Sub
End If
'Y si encontró resultados mostrar la descripción del cliente en un textbox'
txtcodigo.Text = Adodc1.Recordset.Fields(0).Value
txtnombre.Text = Adodc1.Recordset.Fields(1).Value
txtprecio.Text = Adodc1.Recordset.Fields(2).Value
txtstock.Text = Adodc1.Recordset.Fields(3).Value
txtcolor.Text = Adodc1.Recordset.Fields(4).Value
txtId_Producto.Text = Adodc1.Recordset.Fields(5).Value
Exit Sub
salida:
End Sub

Private Sub cmdeliminar_Click()
On Error GoTo salida
Adodc1.Recordset.Delete
MsgBox "Se eliminaron los datos correctamente", vbInformation, "Sistema de productos"
Adodc1.Recordset.AddNew
Adodc1.Refresh
Exit Sub
salida:
MsgBox "Los campos estan vacios busque datos a eliminar", vbCritical, "Ssistema de productos"
End Sub

Private Sub cmdguardar_Click()
On Error GoTo salida
Adodc1.Recordset.Update
    With RsProductos
        .Requery 'Actualizar la tabla y ubicarnos en el primer registro
        .AddNew 'Adicionar un nuevo item
        
        'Paso los valores de la cajas de texto del formulario a la BD
        !Código = txtcodigo.Text
        !Nombre = txtnombre.Text
        !Precio = txtprecio.Text
        !Stock = txtstock.Text
        !Color = txtcolor.Text
        !Id_producto = txtId_Producto.Text
        
        .UpdateBatch 'Grabar en la BD
    
    End With
'MsgBox "Se guardaron los datos correctamente al registro anterior", vbInformation, "Sistema de productos"
'Adodc1.Recordset.MovePrevious
'If Adodc1.Recordset.BOF Then
'End If
'Exit Sub
salida:
'MsgBox "Los campos estan vacios no se puede guardar hasta llenarlos", vbInformation, "Sistema de productos"
    
MsgBox "El registro fue guardado correctamente", vbInformation
LimpiarCajas
End Sub

Private Sub cmdInventario_Click()
    Form4.Show
    Me.Hide
End Sub

Private Sub cmdnuevo_Click()
On Error GoTo salida
Adodc1.Recordset.AddNew
MsgBox "Clic a lado del codigo para agregar un nuevo registro", vbInformation, "Sistema de productos"
    txtcodigo.SetFocus
Exit Sub
salida:
MsgBox "Dando clic dos veces en nuevo tienes que registar", vbCritical, "Sistema de productos"
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

Private Sub LimpiarCajas()
    txtcodigo.Text = ""
    txtnombre.Text = ""
    txtprecio.Text = ""
    txtstock.Text = ""
    txtcolor.Text = ""
    txtId_Producto.Text = ""
End Sub
