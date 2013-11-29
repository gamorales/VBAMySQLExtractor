Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

'Ejecutamos el Formulario
Sub MacroMain()
  Load frmMySQL
  frmMySQL.Show
End Sub

Public Sub AddCustomIcon()
  Dim cbar As Object
  Dim NewBinaps As Object
  Dim NewCRM As Object
  Dim NewConfig As Object
  Dim NewMenu As Object

  Set cbar = CommandBars.Add(Name:="MySQLExtractor", Position:=msoBarTop, Temporary:=True)
  cbar.Visible = True
  cbar.Enabled = True

  Set NewMenu = CommandBars("MySQLExtractor")
  With NewMenu
     .Visible = True
     .Enabled = True
     .Controls.Add(Type:=msoControlButton, before:=1).Caption = "Binaps"
     .Controls.Add(Type:=msoControlButton, before:=2).Caption = "CRM"
     .Controls.Add(Type:=msoControlButton, before:=3).Caption = "Configurar"
     .Protection = msoBarNoChangeVisible
  End With

  Set NewBinaps = CommandBars("MySQLExtractor").Controls("Binaps")
  With NewBinaps
     .Caption = "Binaps"
     .Visible = True
     .FaceId = 4005
     .OnAction = "'" & ThisWorkbook.Name & "'!MacroMain"
  End With

  Set NewCRM = CommandBars("MySQLExtractor").Controls("CRM")
  With NewCRM
     .Caption = "CRM"
     .Visible = True
     .FaceId = 609
     .OnAction = "'" & ThisWorkbook.Name & "'!MacroMain"
  End With

  Set NewConfig = CommandBars("MySQLExtractor").Controls("Configurar")
  With NewConfig
     .Caption = "Configurar"
     .Visible = True
     .Enabled = False
     .FaceId = 3984
     .OnAction = "'" & ThisWorkbook.Name & "'!MacroMain"
  End With

End Sub
