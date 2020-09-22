Attribute VB_Name = "ASMMemory"
Option Explicit

Private Declare Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As Long, ByVal dwSize As Long, _
    ByVal flAllocType As Long, ByVal flProtect As Long _
) As Long

Private Declare Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As Long, ByVal dwSize As Long, _
    ByVal dwFreeType As Long _
) As Long

Public Declare Function VirtualProtect Lib "kernel32" ( _
    ByVal lpAddress As Long, ByVal dwSize As Long, _
    ByVal flNewProtect As Long, lpflOldProtect As Long _
) As Long

Private Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, Optional ByVal dwLen As Long = 4 _
)

Public Declare Function CallWindowProc Lib "user32" _
Alias "CallWindowProcA" ( _
    ptr As Any, p1 As Any, p2 As Any, p3 As Any, p4 As Any _
) As Long

Private Enum VirtualFreeTypes
    MEM_DECOMMIT = &H4000
    MEM_RELEASE = &H8000
End Enum

Private Enum VirtualAllocTypes
    MEM_COMMIT = &H1000
    MEM_RESERVE = &H2000
    MEM_RESET = &H8000
    MEM_LARGE_PAGES = &H20000000
    MEM_PHYSICAL = &H100000
    MEM_WRITE_WATCH = &H200000
End Enum

Private Enum VirtualAllocPageFlags
    PAGE_EXECUTE = &H10
    PAGE_EXECUTE_READ = &H20
    PAGE_EXECUTE_READWRITE = &H40
    PAGE_EXECUTE_WRITECOPY = &H80
    PAGE_NOACCESS = &H1
    PAGE_READONLY = &H2
    PAGE_READWRITE = &H4
    PAGE_WRITECOPY = &H8
    PAGE_GUARD = &H100
    PAGE_NOCACHE = &H200
    PAGE_WRITECOMBINE = &H400
End Enum

Public Type Memory
    address                     As Long
    Bytes                       As Long
End Type

Private Const IDE_ADDROF_REL    As Long = 22


Public Function AsmToMem(ByVal asm As String) As Memory
    Dim clsAsm   As ASMBler
    Dim udtMem   As Memory
    Dim btAsm()  As Byte
    
    Set clsAsm = New ASMBler
    
    ' first we have to determine the size of the output
    ' because the base address isn't known yet, but we
    ' can't allocate memory without knowing the size
    If Not clsAsm.Assemble(asm, True) Then
        Err.Raise 54321, , "(line " & clsAsm.LastErrorLine & ")  " & clsAsm.LastErrorMessage
    End If
    
    udtMem = AllocMemory(clsAsm.OutputSize, , PAGE_EXECUTE_READWRITE)
    
    clsAsm.BaseAddress = udtMem.address
    If Not clsAsm.Assemble(asm) Then
        Err.Raise 54321, , "(line " & clsAsm.LastErrorLine & ")  " & clsAsm.LastErrorMessage
    End If
    
    btAsm = clsAsm.GetOutput()
    CpyMem ByVal udtMem.address, btAsm(0), udtMem.Bytes
    
    AsmToMem = udtMem
End Function


Private Function AllocMemory( _
    ByVal Bytes As Long, _
    Optional ByVal lpAddr As Long = 0, _
    Optional ByVal PageFlags As VirtualAllocPageFlags = PAGE_READWRITE _
) As Memory

    With AllocMemory
        .address = VirtualAlloc(lpAddr, Bytes, MEM_COMMIT, PageFlags)
        .Bytes = Bytes
    End With
End Function


Public Function FreeMemory(udtMem As Memory) As Boolean
    VirtualFree udtMem.address, udtMem.Bytes, MEM_DECOMMIT

    udtMem.address = 0
    udtMem.Bytes = 0
End Function

