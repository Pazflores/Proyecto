Attribute VB_Name = "Module1"
'Variables para conexion a la base de datos
Global Base As New ADODB.Connection
'Variable para acceder a la tabla Usuario
Global RsCliente As New ADODB.Recordset
'Global RsDetalleFactura As New ADODB.Recordset
'Global RsFactura As New ADODB.Recordset
Global RsProductos As New ADODB.Recordset
'Global RsTipodeProducto As New ADODB.Recordset

Sub main()
    With Base
        .CursorLocation = adUseClient 'Vamos a ser clientes de la base de datos
        'Conexion a la base de datos
        '.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Usuario\Desktop\Proyectopap\Proyecto\Base_de_Datos.mdb;Persist Security Info=False"
        .Open " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\PAZ\Desktop\repositorio\Proyecto\Base_de_Datos.mdb;Persist Security Info=False "
        '.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Karen\Desktop\Papeleria\Proyecto\Base_de_Datos.mdb;Persist Security Info=False"
        Form1.Show
    End With
End Sub

'Procedimiento para manejar la tabla Productos
Sub Productos()
    With RsProductos
        If .State = 1 Then .Close
            .Source = "Productos"
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open "select * from Productos", Base
    End With
End Sub

Sub Cliente()
    With RsCliente
        If .State = 1 Then .Close
            .Source = "Cliente"
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open "select * from Cliente", Base
    End With
End Sub
