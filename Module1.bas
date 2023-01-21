Attribute VB_Name = "Module1"
Public cnn As New ADODB.Connection
Public Rs As New ADODB.Recordset

Public sID As String
Public sNombre As String
Public sApellido As String
Public sTelefono As String
Public sDireccion As String
Public Modo As String



Public Sub IniciarConexion()

    With cnn
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
              App.Path & "\Base.mdb" & ";Persist Security Info=False"
    End With
End Sub

Public Sub RefrescarLista()

Dim SubItems As ListItem
Rs.Requery
frmMain.LV.ListItems.Clear
If Rs.RecordCount > 0 Then
Rs.MoveFirst
    Do Until Rs.EOF = True
        sID = Rs("ID")
        sNombre = Rs("Nombre")
        sApellido = Rs("Apellido")
        sTelefono = Rs("Telefono")
        sDireccion = Rs("Direccion")
        Set SubItems = frmMain.LV.ListItems.Add(, , sID)
        SubItems.SubItems(1) = sNombre
        SubItems.SubItems(2) = sApellido
        SubItems.SubItems(3) = sTelefono
        SubItems.SubItems(4) = sDireccion
        Rs.MoveNext
    Loop

Rs.Requery
End If

End Sub

Sub Agregar()

frmDatos.Caption = "Agregar Entrada"
Modo = "Agregar"
frmDatos.Show vbModal

End Sub

Sub Editar()

If frmMain.LV.ListItems.Count = 0 Then
MsgBox "No existen registros en la base de datos", vbOKOnly + vbInformation, "Error"
Exit Sub
End If
               
    frmDatos.txtNombre = frmMain.LV.SelectedItem.SubItems(1)
    frmDatos.txtApellido = frmMain.LV.SelectedItem.SubItems(2)
    frmDatos.txtTelefono = frmMain.LV.SelectedItem.SubItems(3)
    frmDatos.txtDireccion = frmMain.LV.SelectedItem.SubItems(4)
    frmDatos.txtId = frmMain.LV.SelectedItem

frmDatos.Caption = "Editar Entrada"
Modo = "Editar"
frmDatos.Show vbModal
End Sub

Function BuscarCaracter(Cadena As String)

Dim i As Integer
i = 1
    For i = 1 To Len(Cadena)
    
    If Mid(Cadena, i) = "'" Then
        MsgBox "Caracter ' invalido", vbOKOnly + vbExclamation, "Error"
        BuscarCaracter = 1
    End If
    
    Next i

End Function

