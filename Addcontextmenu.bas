Attribute VB_Name = "Addcontextmenu"
Sub AddToCellMenu()
    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl

    ' Delete the controls first to avoid duplicates.
    Call DeleteFromCellMenu

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

   ' Add one built-in button(Save = 3) to the Cell context menu.
    ContextMenu.Controls.Add Type:=msoControlButton, ID:=3, before:=1

    ' Add one custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, before:=2)
        .OnAction = "'" & ThisWorkbook.name & "'!" & "ToggleCaseMacro"
        .FaceId = 610
        .Caption = "D2Refine..."
        .Tag = "My_Cell_Control_Tag"
    End With
 ' Add a custom submenu with three buttons.
    Set MySubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=3)

    With MySubMenu
        .Caption = "Rest Services"
        .Tag = "My_Cell_Control_Tag"
        
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.name & "'!" & "Addservice"
            .FaceId = 50
            .Caption = "Add Service"
        End With

        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.name & "'!" & "ToggleCaseMacro"
            .FaceId = 50
            .Caption = "Service1"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.name & "'!" & "checkblank"
            .FaceId = 50
            .Caption = "CheckBlank"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.name & "'!" & "Emptysheet"
            .FaceId = 50
            .Caption = "Empty Sheet"
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.name & "'!" & "ReadXML"
            .FaceId = 50
            .Caption = "Import Xml DBGep"
        End With
      
    End With
    ' Add a separator to the Cell context menu.
    ContextMenu.Controls(4).BeginGroup = True
   
End Sub
