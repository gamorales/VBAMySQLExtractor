Dim oConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim objMyPivotCache As PivotCache  'Las variables para crear la tabla dinámica
Dim objMyPivotTable As PivotTable
Dim objExcelSheet As Worksheet     'Para crear una nueva hoja

Private Sub ConnectDB()
  On Error Resume Next
  Set oConn = New ADODB.Connection ''

  oConn.Open "DRIVER={MySQL ODBC 3.51 Driver};" & _
   "SERVER=192.168.0.2;" & _
   "DATABASE=binaps;" & _
   "USER=root;" & _
   "PASSWORD=;" & _
   "Option=3"
End Sub

Private Sub ConnectDBdymB2B()
  On Error Resume Next
  Set oConn = New ADODB.Connection ''

  oConn.Open "DRIVER={MySQL ODBC 3.51 Driver};" & _
   "SERVER=localhost;" & _
   "DATABASE=dymb2b;" & _
   "USER=root;" & _
   "PASSWORD=lerolero;" & _
   "Option=3"
End Sub

Private Sub cmbCampos_Change()
  On Error Resume Next
  Set rs = New ADODB.Recordset  'El recordset para ejecutar sentencias
  Dim query As String   'La consulta SQL
  ConnectDB
  
  query = "SELECT sql FROM sqlescenario WHERE nombre LIKE '" & cmbCampos.Value & "'"
  rs.Open query, oConn, adOpenStatic, adLockOptimistic  'ejecutamos
  
  rs.MoveFirst
  txtSQL.Text = rs.Fields("sql")
End Sub

Private Sub cmdCerrar_Click()
  Unload Me
End Sub

Private Sub cmdGenerarCRM_Click()
  On Error Resume Next
  Set rs = New ADODB.Recordset  'El recordset para ejecutar sentencias
  Dim query As String   'La consulta SQL
  ConnectDBdymB2B
  
  'rs.Open query, oConn, adOpenStatic, adLockOptimistic  'ejecutamos
  rs.Open txtQuery.Text, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
  
  'Crear nueva hoja
  Set objExcelSheet = ActiveWorkbook.Worksheets.Add
  objExcelSheet.Name = "Hoja_CRM"
  
  'Crear Tabla en cache
  Set objMyPivotCache = ActiveWorkbook.PivotCaches.Add(xlExternal)
  Set objMyPivotCache.Recordset = rs

  'Crear Tabla dinámica
  Set objMyPivotTable = ActiveWorkbook.Sheets("Hoja_CRM").PivotTables.Add(objMyPivotCache, Cells(2, 1))
  
  'Cerramos el formulario
  Unload Me

End Sub

Private Sub MultiPage_Change()

End Sub

Private Sub txtDesde_Change()
  txtQuery.Text = "SELECT c.case_number AS Numero, c.name AS Caso, c.date_entered AS Fecha_Creacion, " & _
          "c.date_modified AS Fecha_Modificacion, c.type AS Tipo, c.status AS Estado, c.priority AS Prioridad, " & _
          "c.created_by AS Creado, c.assigned_user_id AS Asignado " & _
          "FROM cases AS c, users AS u " & _
          "WHERE c.created_by = u.id AND c.deleted='0' AND c.date_entered BETWEEN '" & txtDesde.Text & "' AND '" & txtHasta.Text & "'"
End Sub

Private Sub txtHasta_Change()
  txtQuery.Text = "SELECT c.case_number AS Numero, c.name AS Caso, c.date_entered AS Fecha_Creacion, " & _
          "c.date_modified AS Fecha_Modificacion, c.type AS Tipo, c.status AS Estado, c.priority AS Prioridad, " & _
          "c.created_by AS Creado, c.assigned_user_id AS Asignado " & _
          "FROM cases AS c, users AS u " & _
          "WHERE c.created_by = u.id AND c.deleted='0' AND c.date_entered BETWEEN '" & txtDesde.Text & "' AND '" & txtHasta.Text & "'"
End Sub

Private Sub UserForm_Activate()
  On Error Resume Next
  Set rs = New ADODB.Recordset  'El recordset para ejecutar sentencias
  Dim query As String   'La consulta SQL
  
  'Sabemos que usuario está conectado en el PC para la consulta de los escenarios
  Dim sBuffer As String
  Dim lSize As Long
  Dim Usuario As String

  sBuffer = Space$(255)
  lSize = Len(sBuffer)
  Call GetUserName(sBuffer, lSize)
  If lSize > 0 Then
    txtUser.Text = Left$(sBuffer, lSize) & "!0''"
  Else
    txtUser.Text = vbNullString
  End If
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ''''''''''''''''''''HASTA ACA''''''''''''''''''''''''''''''
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  ConnectDB
  
  'Consulto los permisos del usuario en los escenarios
  query = "SELECT escenario FROM usuarioescenario WHERE usuario LIKE '" & txtUser.Text & "' ORDER BY escenario"
  rs.Open query, oConn, adOpenStatic, adLockOptimistic  'ejecutamos
  
  'Me voy al primer registro para comenzar a correr desde ahí por medio de un ciclo
  rs.MoveFirst
  Do While Not rs.EOF
    cmbCampos.AddItem (rs.Fields("escenario")) 'Agrego al combo los registros
    rs.MoveNext
  Loop
 
End Sub

Private Sub cmdGenerar_Click()
  'On Error Resume Next
  Set rs = New ADODB.Recordset  'El recordset para ejecutar sentencias
  ConnectDB
  
  'Abrir Recordset
  rs.Open txtSQL.Text, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText

  'Crear nueva hoja
  Set objExcelSheet = ActiveWorkbook.Worksheets.Add
  objExcelSheet.Name = "Hoja_" + cmbCampos.Value
  
  'Crear Tabla en cache
  Set objMyPivotCache = ActiveWorkbook.PivotCaches.Add(xlExternal)
  Set objMyPivotCache.Recordset = rs

  'Crear Tabla dinámica
  Set objMyPivotTable = ActiveWorkbook.Worksheets("Hoja_" + cmbCampos.Value).PivotTables.Add(objMyPivotCache, Cells(2, 1))
  
  
  'Cerramos el formulario
  Unload Me
  
End Sub
