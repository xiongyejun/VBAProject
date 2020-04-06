Attribute VB_Name = "CFile"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'5-MultiUse����ô����ļ��е�VBA��ĿҲ��ʹ�������
'VBE.ActiveVBProject.VBComponents("ClassName").Properties("Instancing") = 5

Private Enum SeekPos
    OriginF
    CurrentF
    EndF
End Enum

Private lFileLen As Long
Private num_file As Integer

'��ȡlen(b)��byte
Function Read(b() As Byte) As Long
    Dim ilen As Long
    ilen = UBound(b) - LBound(b) + 1
    
    Dim iseek As Long
    iseek = VBA.Seek(num_file)
    If iseek + ilen > lFileLen Then
        ilen = lFileLen - iseek + 1
    End If
    
    iseek = iseek + ilen
    
    Get #num_file, , b
    
    Read = ilen
End Function
Function ReadInteger() As Integer
    Dim i As Integer
    Get #num_file, , i
    ReadInteger = i
End Function
Function ReadLong() As Long
    Dim i As Long
    Get #num_file, , i
    ReadLong = i
End Function

'// ReadAt reads len(b) bytes from the File starting at byte offset off.
Function ReadAt(b() As Byte, offset As Long) As Long
    SeekFile offset, 0
    ReadAt = Read(b)
End Function

'// Seek sets the offset for the next Read or Write on file to offset
' whence:
'    0 means relative to the origin of the file,
'    1 means relative to the current offset
'    ���� means relative to the end.���ʱ��һ��Ҫ���ø�����
'returns the new offset
Function SeekFile(offset As Long, whence As Long) As Long
    Dim iseek As Long
    iseek = VBA.Seek(num_file)
    
    'vba Seek���±�1��ʼ
    If whence = SeekPos.OriginF Then
        iseek = 1 + offset
    ElseIf whence = SeekPos.CurrentF Then
        iseek = iseek + offset
    Else
        iseek = 1 + lFileLen + offset
    End If
    
    Seek #num_file, iseek
    
    SeekFile = iseek
End Function

Public Sub Test()
    MsgBox "test class"
End Sub

'���ֽڷ�ʽ��ȡ�ı�
Function OpenFile(Filename As String) As Long
    num_file = VBA.FreeFile
    
    Open Filename For Binary Access Read As #num_file
    
    lFileLen = VBA.FileLen(Filename)
End Function

Function CloseFile()
    Close #num_file
End Function

Private Sub Class_Terminate()
    CloseFile
End Sub