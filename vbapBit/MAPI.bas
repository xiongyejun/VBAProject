Attribute VB_Name = "MAPI"
Option Explicit

Sub Test()
    Dim l As Long
    
    l = &H7FFFFFFF
    
    Printf "l = %b, %d", l, l
    
    BitMoveLeft l, 8
    Printf "l = %b, %d", l, l
End Sub

Function BitMoveLeft(ByRef l As Long, num As Long) As Long
    Dim i As Long
    For i = 1 To num
        '会溢出 0x7FFF FFFF
        '判断第31位是否=1
        '不管等不等于1都把第31为置换为0，负数待处理
        l = l And &H3FFFFFFF
        l = l * 2
    Next
    
    BitMoveLeft = l
End Function

Function BitMoveRight(ByRef l As Long, num As Long) As Long
    Dim i As Long
    For i = 1 To num
        l = l \ 2
    Next
    
    BitMoveRight = l
End Function
