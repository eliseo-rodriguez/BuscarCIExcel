Public Function BuscarCI(strCI As String) As String
On Error GoTo Error:
    Dim PathAccessFile As String
    Dim sSQL As String
    Dim Cnx As ADODB.Connection
    Dim rs As New ADODB.Recordset
   
    If strCI = "" Then
        Exit Function
    End If
                     
    PathAccessFile = "C:\Users\irama\Documents\Irama Iva\ELECT\ELECT_\2024\REP2024.accdb"
    Set Cnx = CreateObject("ADODB.connection")
   
    If Err.Number <> 0 Then
        Exit Function
    End If
   
    Cnx.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + PathAccessFile + ";Persist Security Info=False;"
   
    Set rs = CreateObject("ADODB.Recordset")
   
    If Err.Number <> 0 Then
        Exit Function
    End If
   
    sSQL = "SELECT NAC, CEDULA,[P_NOMBRE], [S_NOMBRE], [P_APELLIDO], [S_APELLIDO], [P_NOMBRE]+ ' '  +[S_NOMBRE] +' ' +[P_APELLIDO]+ ' '+[S_APELLIDO] AS NOMBRE "
    sSQL = sSQL + " FROM [REP] WHERE CEDULA='" + strCI + "'"

    rs.Open sSQL, Cnx
   
    If (IsNull(rs)) Then
      BuscarCI = "No existe"
      Exit Function
    End If
   
    If (rs.EOF And rs.BOF) Then
      BuscarCI = "No existe"
      Exit Function
    End If
   
    Dim lxNombre As String
    lxNombre = rs.Fields("NOMBRE")
    lxNombre = Replace(lxNombre, """", "")
   
    BuscarCI = lxNombre
   
    Exit Function
Error:
    BuscarCI = Err.Description
    Exit Function
End Function

Public Function BuscarNombresCI(strCI As String) As String
On Error GoTo Error:
    Dim PathAccessFile As String
    Dim sSQL As String
    Dim Cnx As ADODB.Connection
    Dim rs As New ADODB.Recordset
   
    If strCI = "" Then
        Exit Function
    End If
   
    PathAccessFile = "C:\Users\irama\Documents\Irama Iva\ELECT\ELECT_\2024\REP2024.accdb"
    Set Cnx = CreateObject("ADODB.connection")
   
    If Err.Number <> 0 Then
        Exit Function
    End If
   
    Cnx.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + PathAccessFile + ";Persist Security Info=False;"
   
    Set rs = CreateObject("ADODB.Recordset")
   
    If Err.Number <> 0 Then
        Exit Function
    End If
   
    sSQL = "SELECT NAC, CEDULA,[P_NOMBRE], [S_NOMBRE], [P_APELLIDO], [S_APELLIDO], [P_NOMBRE]+ ' '  +[S_NOMBRE] AS NOMBRE "
    sSQL = sSQL + " FROM [REP] WHERE CEDULA='" + strCI + "'"

    rs.Open sSQL, Cnx
   
    If (IsNull(rs)) Then
      BuscarNombresCI = "No existe"
      Exit Function
    End If
   
    If (rs.EOF And rs.BOF) Then
      BuscarNombresCI = "No existe"
      Exit Function
    End If
   
    Dim lxNombre As String
    lxNombre = rs.Fields("NOMBRE")
    lxNombre = Replace(lxNombre, """", "")
   
    BuscarNombresCI = lxNombre
   
    Exit Function
Error:
    BuscarNombresCI = Err.Description
    Exit Function
End Function

Public Function BuscarApellidosCI(strCI As String) As String
On Error GoTo Error:
    Dim PathAccessFile As String
    Dim sSQL As String
    Dim Cnx As ADODB.Connection
    Dim rs As New ADODB.Recordset
   
    If strCI = "" Then
        Exit Function
    End If
   
    PathAccessFile = "C:\Users\irama\Documents\Irama Iva\ELECT\ELECT_\2024\REP2024.accdb"
    Set Cnx = CreateObject("ADODB.connection")
   
    If Err.Number <> 0 Then
        Exit Function
    End If
   
    Cnx.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + PathAccessFile + ";Persist Security Info=False;"
   
    Set rs = CreateObject("ADODB.Recordset")
   
    If Err.Number <> 0 Then
        Exit Function
    End If
   
    sSQL = "SELECT NAC, CEDULA,[P_NOMBRE], [S_NOMBRE], [P_APELLIDO], [S_APELLIDO], [P_APELLIDO]+ ' '  +[S_APELLIDO] AS APELLIDOS "
    sSQL = sSQL + " FROM [REP] WHERE CEDULA='" + strCI + "'"

    rs.Open sSQL, Cnx
   
    If (IsNull(rs)) Then
      BuscarApellidosCI = "No existe"
      Exit Function
    End If
   
    If (rs.EOF And rs.BOF) Then
      BuscarApellidosCI = "No existe"
      Exit Function
    End If
   
    Dim lxApellidos As String
    lxApellidos = rs.Fields("APELLIDOS")
    lxApellidos = Replace(lxApellidos, """", "")
   
    BuscarApellidosCI = lxApellidos
   
    Exit Function
Error:
    BuscarApellidosCI = Err.Description
    Exit Function
End Function

Public Function BuscarDataAdicionalCI(strCI As String, campo As String) As String
On Error GoTo Error:
    Dim PathAccessFile As String
    Dim sSQL As String
    Dim Cnx As ADODB.Connection
    Dim rs As New ADODB.Recordset
   
    If strCI = "" Then
        Exit Function
    End If
   
    PathAccessFile = "C:\Users\irama\Documents\Irama Iva\ELECT\ELECT_\2024\REP2024.accdb"
    Set Cnx = CreateObject("ADODB.connection")
   
    If Err.Number <> 0 Then
        Exit Function
    End If
   
    Cnx.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + PathAccessFile + ";Persist Security Info=False;"
   
    Set rs = CreateObject("ADODB.Recordset")
   
    If Err.Number <> 0 Then
        Exit Function
    End If
   
    sSQL = "SELECT NAC, CEDULA, [P_NOMBRE], [S_NOMBRE], [P_APELLIDO], [S_APELLIDO], [P_NOMBRE]+ ' '  +[S_NOMBRE] +' ' +[P_APELLIDO]+ ' '+[S_APELLIDO] AS NOMBRE "
    sSQL = sSQL + ", [" + campo + "] "
    sSQL = sSQL + " FROM [REP] WHERE CEDULA='" + strCI + "'"

    rs.Open sSQL, Cnx
   
    If (IsNull(rs)) Then
      BuscarDataAdicionalCI = "No existe"
      Exit Function
    End If
   
    If (rs.EOF And rs.BOF) Then
      BuscarDataAdicionalCI = "No existe"
      Exit Function
    End If
   
    Dim lxCampo As String
    lxNombre = rs.Fields(campo)
    lxNombre = Replace(lxNombre, """", "")
   
    BuscarDataAdicionalCI = lxNombre
   
    Exit Function
Error:
    BuscarDataAdicionalCI = Err.Description
    Exit Function
End Function