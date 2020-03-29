Attribute VB_Name = "MTest"
Option Explicit

Public Sub Test()

    ThisWorkbook.VBProject.VBComponents("CTest").Properties("Instancing") = 5
    
    MsgBox "test module"
End Sub
