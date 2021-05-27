Attribute VB_Name = "Module1"
Sub main()
    With Base
        .CursorLocation = adUseClient 'Vamos a ser clientes de la base de datos
        'Conexion a la base de datos
        .Open " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\PAZ\Desktop\repositorio\Proyecto\Base_de_Datos.mdb;Persist Security Info=False "
    End With
End Sub
