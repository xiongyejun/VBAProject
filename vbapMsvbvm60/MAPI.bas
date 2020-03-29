Attribute VB_Name = "MAPI"
Option Explicit

'获取到的是保存了数组地址的地址
'比如函数返回 &H12345678，在&H12345678再读取4个字节返回&H0ABCDE0F，&H0ABCDE0F才是数组SafeArray结构的地址
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Var() As Any) As Long

'http://www.exceloffice.net/archives/1008
'typedef struct tagSAFEARRAY {
'  USHORT         cDims;
'  USHORT         fFeatures;
'  ULONG          cbElements;
'  ULONG          cLocks;
'  PVOID          pvData;
'  SAFEARRAYBOUND rgsabound[1];
'} SAFEARRAY, *LPSAFEARRAY;

'https://www.cnblogs.com/jiabei521/archive/2012/10/31/2747797.html
'1 typedef struct tagSAFEARRAYBOUND
'2 {
'3 　　unsigned long cElements;
'4 　　unsigned long lLbound;
'5 } SAFEARRAYBOUND;

Type SafeArrayBound
      cElements As Long '该维的长度
      lLbound As Long '该维的数组存取的下限，一般为0
End Type

Type SafeArray
    cDims As Integer ' 数组的维度
    fFeatures As Integer
    cbElements As Long ' 数组元素的字节大小
    cLocksas As Long
    pvDataas As Long '数组的数据
   rgsabound(1) As SafeArrayBound
End Type

Sub TestArray()
    Dim b(3) As Byte
    Dim bptr As Long
    
    b(0) = &H56
    b(3) = &H55
    
'    TestVariantPtr b
    
    bptr = VarPtrArray(b)
    
    Dim bb(4 - 1) As Byte
    CopyMemory VarPtr(bb(0)), bptr, 4
    
    Dim bptrptr As Long
    bptrptr = 0
    CopyMemory VarPtr(bptrptr), bptr, 4
    
    Printf "VarPtr(bptr) = 0x%x, bptr = 0x%x, VarPtr(bb(0)) = 0x%x, bb = % x, VarPtr(bptrptr) = 0x%x, bptrptr = 0x%x", VarPtr(bptr), bptr, VarPtr(bb(0)), bb, VarPtr(bptrptr), bptrptr
    Exit Sub
    
    Dim sa As SafeArray
    CopyMemory VarPtr(sa.cDims), bptrptr, Len(sa)
    Printf "Sub bptr = 0x%x, bptrptr = 0x%x, %d, %x, Len(sa) = %d", bptr, bptrptr, sa.cDims, sa.pvDataas, Len(sa)
End Sub


Function TestVariantPtr(v As Variant)
    Dim lenth As Long
    lenth = 16
    
    Dim b() As Byte
    ReDim b(lenth - 1) As Byte
    
    CopyMemory VarPtr(b(0)), VarPtr(v), lenth
    
    Dim ptr As Long
    CopyMemory VarPtr(ptr), VarPtr(b(8)), 4
    
    Printf "VarType(v) = 0x%x, b = 0x% x, ptr = 0x%x", VarType(v), b, ptr
    Dim bptrptr As Long
    CopyMemory VarPtr(bptrptr), ptr, 4
    
    Dim sa As SafeArray
    CopyMemory VarPtr(sa.cDims), bptrptr, Len(sa)
    Printf "Function ptr = 0x%x,bptrptr = 0x%x, %d, %x", ptr, bptrptr, sa.cDims, sa.pvDataas
    
End Function