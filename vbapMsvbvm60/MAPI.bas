Attribute VB_Name = "MAPI"
Option Explicit

'��ȡ�����Ǳ����������ַ�ĵ�ַ
'���纯������ &H12345678����&H12345678�ٶ�ȡ4���ֽڷ���&H0ABCDE0F��&H0ABCDE0F��������SafeArray�ṹ�ĵ�ַ
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
'3 ����unsigned long cElements;
'4 ����unsigned long lLbound;
'5 } SAFEARRAYBOUND;

Type SafeArrayBound
      cElements As Long '��ά�ĳ���
      lLbound As Long '��ά�������ȡ�����ޣ�һ��Ϊ0
End Type

Type SafeArray
    cDims As Integer ' �����ά��
    fFeatures As Integer
    cbElements As Long ' ����Ԫ�ص��ֽڴ�С
    cLocksas As Long
    pvDataas As Long '���������
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