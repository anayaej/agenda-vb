Attribute VB_Name = "Module1"
Public cnn As New ADODB.Connection
Public cnneventos As New ADODB.Connection

Public Rs As New ADODB.Recordset
Public Rs2 As New ADODB.Recordset

Public sID As String
Public sNombre As String
Public sApellido As String
Public sTelefono As String
Public sDireccion As String
Public sCelular As String
Public sEmail As String
Public sTrabajo As String
Public sEmpresa As String
Public sPuesto As String

Public evModo As String
Public Modo As String




Public Sub IniciarConexion()

    With cnn
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
              App.Path & "\Datos.dat" & ";Persist Security Info=False"
    End With

    
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

Public Sub RefrescarListaEventos()
Dim restante As String
    
    If Rs2.RecordCount > 0 Then
    Rs2.MoveFirst
    Do Until Rs2.EOF = True
    Set evento = frmEventos.LVEventos.ListItems.Add(, , Rs2("ID"))
    evento.SubItems(1) = Rs2("Fecha")
    restante = Int(1 + Rs2("Fecha") - Now)
    Select Case restante
        Case Is < 0: evento.SubItems(2) = "Caducado"
        Case Is = 0: evento.SubItems(2) = "Hoy"
        Case Is = 1: evento.SubItems(2) = "Mañana"
        Case Is > 0: evento.SubItems(2) = restante
    End Select
    evento.SubItems(3) = Rs2("Nombre")
    evento.SubItems(4) = Rs2("Descripcion")
    evento.SubItems(5) = Rs2("NotasAdicionales")
    evento.SubItems(6) = Rs2("Estado")
    Rs2.MoveNext
    Loop
    End If
End Sub


Public Sub RefrescarLista()

Dim SubItems As ListItem
If Rs.State = 0 Then Rs.Open "SELECT * FROM Datos", cnn

frmMain.LV.ListItems.Clear
If Rs.RecordCount > 0 Then
Rs.MoveFirst
    Do Until Rs.EOF = True
        sID = Rs("ID")
        sNombre = Rs("Nombre")
        sApellido = Rs("Apellido")
        sTelefono = Rs("Thogar")
        sDireccion = Rs("Direccion")
        sCelular = Rs("TCelular")
        sTrabajo = Rs("TTrabajo")
        sEmail = Rs("Email")
        sEmpresa = Rs("Empresa")
        sPuesto = Rs("Puesto")
              
        
        Set SubItems = frmMain.LV.ListItems.Add(, , sID)
        SubItems.SubItems(1) = sNombre
        SubItems.SubItems(2) = sApellido
        SubItems.SubItems(3) = sDireccion
        SubItems.SubItems(4) = sTelefono
        SubItems.SubItems(5) = sCelular
        SubItems.SubItems(6) = sTrabajo
        SubItems.SubItems(7) = sEmail
        SubItems.SubItems(8) = sEmpresa
        SubItems.SubItems(9) = sPuesto
        Rs.MoveNext
    Loop


End If

End Sub

