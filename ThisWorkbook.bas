Option Explicit
'
Private Const C_TAG = "ChipAddIn" ' C_TAG should be a string unique to this add-in.
Private Const C_TOOLS_MENU_ID As Long = 30007&

Private Sub Workbook_Open()
'''''''''''''''''''''''''''''''''''''''''''''''
' Workbook_Open
' Create a submenu on the Tools menu. The
' submenu has two controls on it.
'''''''''''''''''''''''''''''''''''''''''''''''
  Dim ToolsMenu As Office.CommandBarControl
  Dim ToolsMenuItem As Office.CommandBarControl

'''''''''''''''''''''''''''''''''''''''''''''''
' First delete any of our controls that
' may not have been properly deleted previously.
'''''''''''''''''''''''''''''''''''''''''''''''
  DeleteControls

''''''''''''''''''''''''''''''''''''''''''''''
' Get a reference to the Tools menu.
''''''''''''''''''''''''''''''''''''''''''''''
  Set ToolsMenu = Application.CommandBars.FindControl(ID:=C_TOOLS_MENU_ID)
  If ToolsMenu Is Nothing Then
      MsgBox "Unable to access Tools menu.", vbOKOnly
      Exit Sub
  End If

''''''''''''''''''''''''''''''''''''''''''''''
' Create a item on the Tools menu.   'Set ToolsMenuItem = ToolsMenu.Controls.Add(Type:=msoControlPopup, temporary:=True)
''''''''''''''''''''''''''''''''''''''''''''''
  Set ToolsMenuItem = ToolsMenu.Controls.Add(Type:=msoControlButton, Temporary:=True)
  If ToolsMenuItem Is Nothing Then
      MsgBox "Unable to add item to the Tools menu.", vbOKOnly
      Exit Sub
  End If

  With ToolsMenuItem
      .Caption = "&MySQL Extractor"
      .OnAction = "'" & ThisWorkbook.Name & "'!AddCustomIcon"
      .Tag = C_TAG
  End With

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''
' Workbook_BeforeClose
' Before closing the add-in, clean up our controls.
''''''''''''''''''''''''''''''''''''''''''''''''''''
    DeleteControls
End Sub


Private Sub DeleteControls()
''''''''''''''''''''''''''''''''''''
' Delete controls whose Tag is
' equal to C_TAG.
''''''''''''''''''''''''''''''''''''
Dim Ctrl As Office.CommandBarControl

'On Error Resume Next
Set Ctrl = Application.CommandBars.FindControl(Tag:=C_TAG)

Do Until Ctrl Is Nothing
    Ctrl.Delete
    Set Ctrl = Application.CommandBars.FindControl(Tag:=C_TAG)
Loop

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' END ThisWorkbook Code Module
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



