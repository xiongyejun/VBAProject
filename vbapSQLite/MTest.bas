Attribute VB_Name = "MTest"
Option Explicit

Type FieldInfo
    FName As String
    Type As String
    pk As Long '主键，0不是 1是
End Type

Type TableInfo
    tableName As String
    Fields() As FieldInfo
End Type

Public Sub Test()

    ThisWorkbook.VBProject.VBComponents("CSQLite3").Properties("Instancing") = 5
    
    MsgBox "test module"
End Sub