Attribute VB_Name = "MAPI"
Option Explicit

Public Declare Function gosprintf Lib "godllForVBA32.dll" (ByVal pFormat As Long, ByVal pVBAVariant As Long, ByVal nCount As Long) As MyString
Public Declare Function sum Lib "godllForVBA32.dll" (ByVal a As Long, ByVal b As Long) As Long
Private Declare Sub cfree Lib "godllForVBA32.dll" (ByVal p As Long)
Private Declare Function retVarPtr Lib "godllForVBA32.dll" () As Long


Type MyString
    pUCS2 As Long
    Len As Long
End Type


Sub teststring()
    Dim str As String
    
    str = "阿斯顿发送方abcd"
    
    Dim b() As Byte
    ReDim b(6 + Len(str) - 1) As Byte
    
    vbapAPIkernel32.CopyMemory VarPtr(b(0)), StrPtr(str) - 6, 6 + Len(str)
    
    Printf "% x", b
    
End Sub

Sub testretVarPtr()
Dim pdll As Long
    
    pdll = vbapAPIkernel32.LoadLibrary(ThisWorkbook.Path & "\godllForVBA32.dll")
    
    Dim str As String
    str = "1"
    
    Dim strVarPtr As Long
    strVarPtr = retVarPtr()
    
    Dim b(100) As Byte
    vbapAPIkernel32.CopyMemory VarPtr(b(0)), strVarPtr - 6, 100
    Printf "% x", b
    
    vbapAPIkernel32.CopyMemory VarPtr(str), VarPtr(strVarPtr), 4
    Debug.Print VarPtr(str)
    
    
    
    Debug.Print str
    
    
    vbapAPIkernel32.FreeLibrary pdll
End Sub

Public Sub testPrintf()
    Dim i As Long

    i = 2400
    Dim s As String
    s = "as中文dfaf"
    
    Dim b As Boolean
    b = True
    
    Dim arr() As Byte
    ReDim arr(3) As Byte
    arr(0) = 13
    arr(1) = 24
    arr(2) = 236
    arr(3) = 34
    
    Dim arr1(3) As Byte
    arr1(0) = 13
    arr1(1) = 24
    arr1(2) = 236
    arr1(3) = 34
    
    Dim arr2(2) As Long
    arr2(0) = 132414
    arr2(1) = 34534546

    Dim arr3(2) As String
    arr3(0) = "ab"
    arr3(1) = "cd在啊"
    
    Dim arr4(6) As Boolean
    arr4(0) = True
    arr4(5) = True
    
    Dim f As Double
    f = 66.66666
    
    Dim arr5(2) As Single
    arr5(0) = 66.66666
    arr5(1) = 142424.13
    
    Printf "num = %d,  % d, %s, %t, %d, %d, %d, %s, %t, %f, %f, %f", i, 24, s, b, arr, arr1, arr2, arr3, arr4, 5.6, f, arr5
End Sub


Public Sub Printf(format As Variant, ParamArray args() As Variant)
    
    Dim pdll As Long
    
    pdll = vbapAPIkernel32.LoadLibrary(ThisWorkbook.Path & "\godllForVBA32.dll")
    
    Dim ms As MyString
    Dim nCount As Long
    '可能传入0个参数
    On Error Resume Next
    nCount = UBound(args) + 1
    On Error GoTo 0
    
    
'    Dim b() As Byte
'    ReDim b(15) As Byte
'    vbapAPIkernel32.CopyMemory VarPtr(b(0)), VarPtr(args(6)), 16
'    printbyte b
'    Stop
'
'    ReDim b(3) As Byte
'    'variant中保存的地址，得到的是safearray结构地址
'    vbapAPIkernel32.CopyMemory VarPtr(b(0)), &H37F1E0, 16
'    printbyte b
'    Stop
'
'    'safearray结构地址
'    vbapAPIkernel32.CopyMemory VarPtr(b(0)), &H37F270, 4
'    printbyte b
'    Stop
'
'    '元素长度
'    vbapAPIkernel32.CopyMemory VarPtr(b(0)), &H37F270 + 4, 4
'    printbyte b
'    Stop
'    '数据地址
'    vbapAPIkernel32.CopyMemory VarPtr(b(0)), &H37F270 + 12, 4
'    printbyte b
'    Stop
'    '数据地址指向string地址
'    vbapAPIkernel32.CopyMemory VarPtr(b(0)), &H13C20AE8, 4
'    printbyte b
'    Stop
'
'    '得到数据
'    vbapAPIkernel32.CopyMemory VarPtr(b(0)), &H1401CE6C, 4
'    printbyte b
'    Stop
'
'
'    vbapAPIkernel32.CopyMemory VarPtr(b(0)), &H13C20AE8 + 8, 4
'    printbyte b
'    Stop
'
'    vbapAPIkernel32.CopyMemory VarPtr(b(0)), &H1B4F203C, 4
'    printbyte b
'    Stop

    If nCount Then
        ms = gosprintf(VarPtr(format), VarPtr(args(0)), nCount)
    Else
        ms = gosprintf(VarPtr(format), 0, 0)
    End If

    Dim str As String
    Dim b() As Byte
    ReDim b(ms.Len - 1) As Byte
    vbapAPIkernel32.CopyMemory VarPtr(b(0)), ms.pUCS2, ms.Len
    
    str = b

    Debug.Print str
    
    cfree ms.pUCS2
    vbapAPIkernel32.FreeLibrary pdll
    
End Sub

'Function GetStringFromGOString(g As GoString) As String
'    Dim b() As Byte
'
'    ReDim b(g.Len - 1) As Byte
'
'    CopyMemory VarPtr(b(0)), g.p, g.Len
'
'    Dim str As String
'    str = ByteToStr(b, "utf-8")
'
'    GetStringFromGOString = str
'End Function

'Function ByteToStr(arrByte() As Byte, strCharset As String) As String
'    With CreateObject("Adodb.Stream")
'        .Type = 1 'adTypeBinary
'        .Open
'        .Write arrByte
'        .Position = 0
'        .Type = 2 'adTypeText
'        .Charset = strCharset
'        ByteToStr = .Readtext
'        .Close
'    End With
'
'End Function


Sub printbyte(b() As Byte)
    Dim i As Long
    
    For i = 0 To UBound(b)
        Debug.Print i, VBA.Hex(b(i))
    Next
End Sub