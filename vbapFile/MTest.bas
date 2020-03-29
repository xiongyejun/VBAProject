Attribute VB_Name = "MTest"
Option Explicit

Public Sub Test()

    ThisWorkbook.VBProject.VBComponents("CTest").Properties("Instancing") = 5
    
    MsgBox "test module"
End Sub

Sub TestCFile()
    Dim f As CFile
    Set f = New CFile
    
    f.OpenFile ThisWorkbook.Path & "\test.txt"

    Dim b() As Byte
    ReDim b(12) As Byte
    Dim ret As Long
    
'    f.SeekFile 1, 1
    ret = f.Read(b)
    
    Printf "b = 0x% x, ret = %d", b, ret
    
    Dim i As Integer
    i = f.ReadInteger()
    Printf "i = %d, i = 0x%x", i, i
    
    f.CloseFile
End Sub
