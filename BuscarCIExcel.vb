Public Function BuscarCI(strCI As String) As String
On Error GoTo Error:
    Dim PathAccessFile As String
    Dim sSQL As String
    Dim Cnx As ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    If strCI = "" Then
        Exit Function
    End If
    
    PathAccessFile = "C:\Users\rodri\OneDrive\Documents\Rep_CI.accdb"
    Set Cnx = CreateObject("ADODB.connection")
    
    If Err.Number <> 0 Then
        Exit Function
    End If
    
    Cnx.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + PathAccessFile + ";Persist Security Info=False;"
    
    Set rs = CreateObject("ADODB.Recordset")
    
    If Err.Number <> 0 Then
        Exit Function
    End If
    
    sSQL = "SELECT Nacionalidad, CEDULA,[P NOMBRE], [S NOMBRE], [P APELLIDO], [S APELLIDO], [P NOMBRE]+ "" "" +[S NOMBRE] +"" ""+[P APELLIDO]+"" ""+[S APELLIDO] AS NOMBRE FROM [REPCI] WHERE CEDULA='" + strCI + "'"
    rs.Open sSQL, Cnx
    Dim lxNombre As String
    lxNombre = rs.Fields("NOMBRE")
    BuscarCI = lxNombre
    
    Exit Function
Error:
    BuscarCI = Err.Description
    Exit Function
End Function
