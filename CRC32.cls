VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
                (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GlobalAlloc Lib "kernel32" _
                (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
                
Private Declare Function GlobalFree Lib "kernel32" _
                (ByVal hMem As Long) As Long

Private Declare Function VirtualProtect Lib "kernel32" _
                (lpAddress As Any, ByVal dwSize As Long, _
                 ByVal flNewProtect As Long, lpflOldProtect _
                 As Long) As Long
                 
Private Declare Function CallWindowProc Lib "user32" Alias _
                "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                ByVal hWnd As Long, ByVal Msg As Long, ByVal _
                wParam As Long, ByVal lParam As Long) As Long


Private mc_EntriesC As Long
Private ptrVTable   As Long
Private ptrCode     As Long

Private Const CODE_STR As String = "FC608B7424248B7C242833C9498BD133C033DBAC32C18ACD8AEA8AD6B60866D1EB66D1D87309663520836681F3B8EDFECE75EB33C833D34F75D5F7D2F7D192C1C010668BC18944241C61C21000"



Private Function pvAddMC(ByVal m_Hex As String, Optional ByVal m_Redirect As Boolean = False, Optional ByVal m_EntryIndex As Long = -1) As Long

Dim bCode() As Byte
Dim nLen As Long
Dim iCount As Long
    
    nLen = Len(m_Hex)
    ReDim bCode(nLen / 2 - 1) As Byte
    
    For iCount = 1 To nLen Step 2
        bCode((iCount - 1) / 2) = Val("&H" & Mid(m_Hex, iCount, 2))
    Next iCount
    
    nLen = nLen / 2
    pvAddMC = GlobalAlloc(0&, nLen)
    Call VirtualProtect(ByVal pvAddMC, nLen, &H40&, iCount)
    CopyMem ByVal pvAddMC, bCode(0), nLen
    
    If m_Redirect Then
        If m_EntryIndex = -1 Then
            CopyMem ByVal (ptrVTable + (mc_EntriesC * 4)), pvAddMC, 4
        Else
            CopyMem ByVal (ptrVTable + ((m_EntryIndex - 1) * 4)), pvAddMC, 4
        End If
    End If
    
    mc_EntriesC = mc_EntriesC + 1

End Function

Private Sub Class_Initialize()

' --- INITIALIZATION ---
    CopyMem ptrVTable, ByVal ObjPtr(Me), 4
    ptrVTable = ptrVTable + &H1C

' --- USER PATCHES ---
    ptrCode = pvAddMC(CODE_STR)
    
End Sub





Public Function CRC32_String(ByVal mString As String) As Long

    Dim mData() As Byte
    
    mData = StrConv(mString, vbFromUnicode)
    CRC32_String = pvCrc32(mData)
    
End Function
Public Function CRC32_File(ByVal mFileName As String) As Long

    Dim mData() As Byte
    Dim F As Integer
    
    F = FreeFile
    ReDim mData(FileLen(mFileName) - 1) As Byte
    
    Open mFileName For Binary Access Read As #F
        Get #F, , mData
    Close #F
    
    CRC32_File = pvCrc32(mData)
    
End Function


Private Function pvCrc32(ByRef mData() As Byte) As Long
    pvCrc32 = CallWindowProc(ptrCode, VarPtr(mData(0)), UBound(mData) + 1, 0, 0)
End Function

Private Sub Class_Terminate()
    GlobalFree ptrCode
End Sub


