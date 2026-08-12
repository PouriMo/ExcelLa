Attribute VB_Name = "Module_Launcher"
Option Explicit

' Main entry point - opens the app window.
' Callable from Alt+F8, keyboard shortcut, or QAT button.
Sub StartExcelLa()
    ExcelLa_Show
End Sub

' Internal helper (avoids name collision with the UserForm itself)
Private Sub ExcelLa_Show()
    UserForms.Add("ExcelLa").Show
End Sub

' Auto-registers Ctrl+Shift+A shortcut when Excel starts with add-in loaded
Sub Auto_Open()
    Application.OnKey "^+a", "StartExcelLa"
End Sub

' Unregister shortcut when add-in is unloaded
Sub Auto_Close()
    Application.OnKey "^+a"
End Sub
