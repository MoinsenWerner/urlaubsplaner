Attribute VB_Name = "Module1"
Option Explicit

Public Sub Urlaubsuebersicht()
    ActiveWindow.FreezePanes = False
    Range("B5").Select
    ActiveWindow.FreezePanes = True
End Sub
