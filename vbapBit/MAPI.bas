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
        '����� 0x7FFF FFFF
        '�жϵ�31λ�Ƿ�=1
        '���ܵȲ�����1���ѵ�31Ϊ�û�Ϊ0������������
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
