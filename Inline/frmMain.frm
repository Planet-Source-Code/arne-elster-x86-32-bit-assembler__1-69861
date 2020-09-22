VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   """Inline"" Assembler"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdRawData 
      Caption         =   "Raw Data"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   3900
      Width           =   1275
   End
   Begin VB.CommandButton cmdFDiv 
      Caption         =   "FDIV Test"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   3900
      Width           =   1275
   End
   Begin VB.CommandButton cmdCPUName 
      Caption         =   "CPU Name"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3900
      Width           =   1275
   End
   Begin VB.ListBox lstSorted 
      Height          =   1815
      Left            =   3008
      TabIndex        =   5
      Top             =   1800
      Width           =   2595
   End
   Begin VB.ListBox lstUnsorted 
      Height          =   1815
      Left            =   308
      TabIndex        =   4
      Top             =   1800
      Width           =   2595
   End
   Begin VB.CommandButton cmdBubble 
      Caption         =   "ASM"
      Height          =   495
      Index           =   1
      Left            =   315
      TabIndex        =   1
      Top             =   840
      Width           =   1395
   End
   Begin VB.CommandButton cmdBubble 
      Caption         =   "VB"
      Height          =   495
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   240
      Width           =   1395
   End
   Begin VB.Label lblSorted 
      Caption         =   "Sortiert:"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblUnsorted 
      Caption         =   "Unsortiert:"
      Height          =   255
      Left            =   300
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   300
      X2              =   5580
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Label lblTimeASM 
      AutoSize        =   -1  'True
      Caption         =   "Zeit:"
      Height          =   195
      Left            =   1935
      TabIndex        =   3
      Top             =   960
      Width           =   330
   End
   Begin VB.Label lblTimeVB 
      AutoSize        =   -1  'True
      Caption         =   "Zeit:"
      Height          =   195
      Left            =   1935
      TabIndex        =   2
      Top             =   360
      Width           =   330
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' VB "Inline" ASM
'
' Arne Elster 2007 / 2008


' all samples are called with CallWindowProc,
' so they must accept 4 arguments
'
' a1: [ebp+12]
' a2: [ebp+16]
' a3: [ebp+20]
' a4: [ebp+24]
'
' "|" = vbCrLf, it's just shorter


' a1: Pointer to Long Array
' a2: Arraylength
Private Const ASM_TEST_BUBBLESORT   As String = _
    "               pushad                                          |" & _
    "               mov esi, [ebp+16]    ; Arraylength              |" & _
    "   outer_loop: mov ebx, [ebp+12]    ; ArrPtr                   |" & _
    "               mov edx, [ebp+16]    ; Arraylength              |" & _
    "               xor edi, edi                                    |" & _
    "   inner_loop: mov eax, [ebx+0]     ; arr(j)                   |" & _
    "               mov ecx, [ebx+4]     ; arr(j+1)                 |" & _
    "               cmp eax, ecx                                    |" & _
    "               jle byte next_loop   ; swap if eax > ecx        |" & _
    "               mov [ebx+0], ecx     ; swap arr(j), arr(j+1)    |" & _
    "               mov [ebx+4], eax                                |" & _
    "               mov edi, 1           ; swapped                  |" & _
    "   next_loop:  add ebx, 4           ;                          |" & _
    "               dec edx              ;                          |" & _
    "               jnz byte inner_loop  ; i > 0 => still in inner  |" & _
    "               test edi, edi        ; swapped?                 |" & _
    "               jz  byte return      ; no => sorted             |" & _
    "               dec esi                                         |" & _
    "               jnz byte outer_loop                             |" & _
    "   return:     popad                                           |" & _
    "               ret &H10                                         "


' a1: Numerator
' a2: Divisor
' a3: Ptr to result (float = single)
Private Const ASM_TEST_FDIV         As String = _
    "   mov    eax,    [ebp+20]   ; Ptr to output float             |" & _
    "   fild   dword   [ebp+12]   ; st0 = numerator                 |" & _
    "   fild   dword   [ebp+16]   ; st0 = divisor, st1 = numerator  |" & _
    "   fdivp                     ; st1 = st1 / st0, pop st0        |" & _
    "   fstp   float   [eax]      ; pop st0 to output float         |" & _
    "   ret    16                                                    "


' a1: ptr to 12 bytes of writable memory
Private Const ASM_TEST_CPUID        As String = _
    "   pushad                              |" & _
    "   mov edi, [ebp+12]                   |" & _
    "   xor eax, eax                        |" & _
    "   cpuid                               |" & _
    "   mov [edi+0], ebx                    |" & _
    "   mov [edi+4], edx                    |" & _
    "   mov [edi+8], ecx                    |" & _
    "   popad                               |" & _
    "   ret 16                               "


' no params
Private Const ASM_TEST_RAWDATA      As String = _
    "         mov eax, [Data]               |" & _
    "         ret 16                        |" & _
    "   Data: dd 123454321                   "


Private m_asmBubblesort             As Memory
Private m_asmFDiv                   As Memory
Private m_asmCPUName                As Memory
Private m_asmRawData                As Memory


Private Sub Form_Load()
    Dim asm As New ASMBler
    
    ' assemble all the code samples above
    m_asmBubblesort = AsmToMem(ASM_TEST_BUBBLESORT)
    m_asmRawData = AsmToMem(ASM_TEST_RAWDATA)
    m_asmCPUName = AsmToMem(ASM_TEST_CPUID)
    m_asmFDiv = AsmToMem(ASM_TEST_FDIV)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    FreeMemory m_asmBubblesort
    FreeMemory m_asmRawData
    FreeMemory m_asmCPUName
    FreeMemory m_asmFDiv
End Sub


Private Sub cmdCPUName_Click()
    Dim btName(11) As Byte
    
    CallWindowProc ByVal m_asmCPUName.address, btName(0), ByVal 0&, ByVal 0&, ByVal 0&
    MsgBox StrConv(btName, vbUnicode)
End Sub


Private Sub cmdFDiv_Click()
    Dim lngDividend As Long, lngDivisor As Long, sngQuotient As Single
    
    lngDividend = 2
    lngDivisor = 5
    
    CallWindowProc ByVal m_asmFDiv.address, _
                   ByVal lngDividend, _
                   ByVal lngDivisor, _
                   sngQuotient, ByVal 0&
    
    MsgBox lngDividend & " / " & lngDivisor & " = " & sngQuotient
End Sub


Private Sub cmdRawData_Click()
    MsgBox CallWindowProc(ByVal m_asmRawData.address, 0, 0, 0, 0)
End Sub


Private Sub cmdBubble_Click(Index As Integer)
    DoBubble Index = 1
End Sub


Private Sub DoBubble(ByVal asm As Boolean)
    Dim d           As Double
    Dim lng(3999)   As Long
    Dim i           As Long
    
    lstUnsorted.Clear
    lstSorted.Clear
    
    Randomize
    For i = 0 To UBound(lng)
        lng(i) = Rnd() * (UBound(lng) + 1) * Sgn(Rnd() - 0.5)
        lstUnsorted.AddItem lng(i)
    Next
    
    d = Timer
    If asm Then
        ASM_Bubblesort lng, UBound(lng) + 1
        lblTimeASM.Caption = "Zeit: " & CLng((Timer - d) * 1000) & " ms (" & UBound(lng) + 1 & " Items)"
    Else
        VB_Bubblesort lng, UBound(lng) + 1
        lblTimeVB.Caption = "Zeit: " & CLng((Timer - d) * 1000) & " ms (" & UBound(lng) + 1 & " Items)"
    End If
    
    For i = 0 To UBound(lng)
        lstSorted.AddItem lng(i)
    Next
End Sub


Private Sub ASM_Bubblesort(lng() As Long, ByVal cnt As Long)
    CallWindowProc ByVal m_asmBubblesort.address, lng(0), ByVal cnt, ByVal 0&, ByVal 0&
End Sub


Private Sub VB_Bubblesort(lng() As Long, ByVal cnt As Long)
    Dim i   As Long
    Dim j   As Long
    Dim t   As Long
    Dim b   As Boolean
    
    For i = 0 To cnt - 1
        b = True
        For j = 0 To cnt - 2
            If lng(j) > lng(j + 1) Then
                t = lng(j + 1)
                lng(j + 1) = lng(j)
                lng(j) = t
                b = False
            End If
        Next
        If b Then Exit For
    Next
End Sub
