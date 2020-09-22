VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VB X86 32 Bit Assembler"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   9390
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
   ScaleHeight     =   6495
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   2340
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.PictureBox picFiles 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5595
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton cmdFileRem 
         Caption         =   "-"
         Enabled         =   0   'False
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   5100
         Width           =   495
      End
      Begin VB.CommandButton cmdFileAdd 
         Caption         =   "+"
         Enabled         =   0   'False
         Height          =   375
         Left            =   60
         TabIndex        =   2
         Top             =   5100
         Width           =   495
      End
      Begin VB.ListBox lstFiles 
         Height          =   4980
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2115
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPrjNew 
         Caption         =   "&New project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuPrjOpen 
         Caption         =   "&Open project"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrjSave 
         Caption         =   "&Save project"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAssemble 
      Caption         =   "&Assemble"
      Enabled         =   0   'False
      Begin VB.Menu mnuAssembleExe 
         Caption         =   "Write &EXE"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExecute 
         Caption         =   "&Start EXE"
      End
      Begin VB.Menu mnuS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubsystem 
         Caption         =   "S&ubsystem"
         Begin VB.Menu mnuSubsysGUI 
            Caption         =   "&GUI"
         End
         Begin VB.Menu mnuSubsysCUI 
            Caption         =   "&CUI"
         End
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "Abo&ut"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PRJ_MAGIC As Long = &H4D5341

Private Type ProjectFile
    Filename()          As String
    FileContent()       As String
    FileCount           As Long
    Subsystem           As Long
End Type

Private m_strPrjFile    As String
Private m_udtProj       As ProjectFile
Private m_lngCurrentIdx As Long

Private m_clsAssembler  As ASMBler
Private m_clsPreproc    As ASMPreprocessor
Private m_clsDlg        As CommonDialog


Private Sub cmdFileAdd_Click()
    Dim strFName    As String
    
    strFName = InputBox("Name of new file:")
    If StrPtr(strFName) <> 0 Then
        With m_udtProj
            ReDim Preserve .FileContent(.FileCount) As String
            ReDim Preserve .Filename(.FileCount) As String
            .Filename(.FileCount) = strFName
            .FileCount = .FileCount + 1
        End With
        
        UpdateFileList
    End If
End Sub


Private Sub UpdateFileList()
    Dim i   As Long
    
    lstFiles.Clear
    
    For i = 0 To m_udtProj.FileCount - 1
        lstFiles.AddItem m_udtProj.Filename(i)
    Next
End Sub


Private Sub cmdFileRem_Click()
    Dim i   As Long
    
    If lstFiles.ListCount > 0 Then
        If lstFiles.ListIndex > -1 Then
            With m_udtProj
                For i = lstFiles.ListIndex To m_udtProj.FileCount - 2
                    .FileContent(i) = .FileContent(i + 1)
                    .Filename(i) = .Filename(i + 1)
                Next
        
                m_udtProj.FileCount = m_udtProj.FileCount - 1
                If m_udtProj.FileCount = 0 Then
                    Erase .FileContent
                    Erase .Filename
                Else
                    ReDim Preserve .FileContent(.FileCount - 1) As String
                    ReDim Preserve .Filename(.FileCount - 1) As String
                End If
            End With
            
            UpdateFileList
            
            If lstFiles.ListCount = 0 Then
                txtCode.Visible = False
                m_lngCurrentIdx = -1
            Else
                lstFiles.ListIndex = 0
            End If
        End If
    End If
End Sub


Private Sub Form_Load()
    Set m_clsAssembler = New ASMBler
    Set m_clsPreproc = New ASMPreprocessor
    Set m_clsDlg = New CommonDialog
    m_lngCurrentIdx = -1
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    picFiles.Width = 160 * Screen.TwipsPerPixelX
    txtCode.Width = Me.ScaleWidth - picFiles.Width
    
    picFiles.Left = 0
    txtCode.Left = picFiles.Width
    
    txtCode.Top = 0
    picFiles.Top = 0
    
    txtCode.Height = Me.ScaleHeight
    picFiles.Height = Me.ScaleHeight
End Sub


Private Sub lstFiles_Click()
    If m_lngCurrentIdx > -1 Then
        m_udtProj.FileContent(m_lngCurrentIdx) = txtCode.Text
    End If
    
    If lstFiles.ListIndex > -1 And lstFiles.ListCount > 0 Then
        txtCode.Text = m_udtProj.FileContent(lstFiles.ListIndex)
        m_lngCurrentIdx = lstFiles.ListIndex
        txtCode.Visible = True
    End If
End Sub


Private Sub mnuAbout_Click()
    MsgBox "X86 32 Bit Assembler" & vbCrLf & _
           vbCrLf & _
           "Arne Elster 2007 / 2008", vbInformation
End Sub


Private Sub mnuAssembleExe_Click()
    Dim strSource   As String
    Dim strNewName  As String
    Dim i           As Long
    Dim fh          As Integer
    Dim btASM()     As Byte
    
    If m_udtProj.FileCount = 0 Then
        MsgBox "No files in current project", vbExclamation
        Exit Sub
    End If
    
    If lstFiles.ListIndex > -1 Then
        m_udtProj.FileContent(lstFiles.ListIndex) = txtCode.Text
    End If
    
    For i = 0 To m_udtProj.FileCount - 1
        strSource = strSource & "#" & m_udtProj.Filename(i) & vbCrLf
        strSource = strSource & m_udtProj.FileContent(i) & vbCrLf
    Next
    
    strSource = m_clsPreproc.Process(strSource)
    
    m_clsAssembler.PEHeader = True
    m_clsAssembler.BaseAddress = &H400000
    m_clsAssembler.Subsystem = m_udtProj.Subsystem
    
    If Not m_clsAssembler.Assemble(strSource) Then
        MsgBox m_clsAssembler.LastErrorMessage & _
               " (" & m_clsAssembler.LastErrorSection & " - " & _
               "line " & m_clsAssembler.LastErrorLine & ")", vbExclamation, "Error"
    Else
        btASM = m_clsAssembler.GetOutput()
        fh = FreeFile()
        
        strNewName = StripExt(m_strPrjFile) & ".exe"
        If Len(Dir(strNewName)) > 0 Then Kill strNewName
        
        Open strNewName For Binary Access Write As #fh
        Put #fh, , btASM
        Close #fh
        
        MsgBox "Assembled successfully!", vbInformation
    End If
End Sub


Private Function StripExt(ByVal strFile As String) As String
    If InStrRev(strFile, ".") > 0 Then
        StripExt = Left$(strFile, InStrRev(strFile, ".") - 1)
    Else
        StripExt = strFile
    End If
End Function


Private Sub mnuExecute_Click()
    Dim strExeFile  As String
    
    strExeFile = StripExt(m_strPrjFile) & ".exe"
    
    If Len(Dir(strExeFile)) > 0 Then
        On Error Resume Next
        Shell strExeFile, vbNormalFocus
        If Err Then MsgBox "Error occured while starting the file", vbExclamation
        On Error GoTo 0
    Else
        MsgBox strExeFile & " not found", vbExclamation
    End If
End Sub


Private Sub mnuExit_Click()
    SaveProject
    Unload Me
End Sub


Private Sub mnuPrjNew_Click()
    Dim strFile     As String
    Dim strFilter   As String
    
    strFilter = "Projects (*.apr)|*.apr|All files (*.*)|*.*"

    If Not m_clsDlg.VBGetSaveFileName(strFile, Filter:=strFilter) Then
        Exit Sub
    End If
    
    If strFile <> "" Then
        m_strPrjFile = strFile
        
        With m_udtProj
            .FileCount = 0
            .Subsystem = Subsystem_GUI
            Erase .Filename
            Erase .FileContent
            
            mnuSubsysGUI.Checked = True
            mnuSubsysCUI.Checked = False
        End With
        
        SetGUIState True
        txtCode.Visible = False
        m_lngCurrentIdx = -1
        UpdateFileList
    End If
End Sub


Private Sub SetGUIState(ByVal state As Boolean)
    cmdFileAdd.Enabled = state
    cmdFileRem.Enabled = state
    mnuPrjSave.Enabled = state
    mnuAssemble.Enabled = state
    If Not state Then m_lngCurrentIdx = -1
End Sub


Private Sub mnuPrjOpen_Click()
    Dim strFile As String
    Dim strFilter   As String
    
    strFilter = "Projects (*.apr)|*.apr|All files (*.*)|*.*"
    
    If Not m_clsDlg.VBGetOpenFileName(strFile, , False, , , , strFilter, _
                                      flags:=OFN_OVERWRITEPROMPT) Then
        Exit Sub
    End If
    
    If strFile <> vbNullString Then LoadProject strFile
End Sub


Private Sub LoadProject(ByVal strFile As String)
    Dim fh          As Integer
    Dim lngMagic    As Long
    
    m_strPrjFile = strFile
    
    fh = FreeFile()
    Open m_strPrjFile For Binary Access Read As #fh
    Get #fh, , lngMagic
    If lngMagic = PRJ_MAGIC Then Get #fh, , m_udtProj
    Close #fh
    
    If lngMagic <> PRJ_MAGIC Then
        SetGUIState False
        m_strPrjFile = vbNullString
        MsgBox "Not an assembler project!", vbExclamation
    Else
        SetGUIState True
    End If
    
    mnuSubsysGUI.Checked = m_udtProj.Subsystem = Subsystem_GUI
    mnuSubsysCUI.Checked = m_udtProj.Subsystem = Subsystem_CUI
    
    txtCode.Visible = False
    m_lngCurrentIdx = -1
    UpdateFileList
End Sub


Private Sub mnuPrjSave_Click()
    If lstFiles.ListIndex > -1 And lstFiles.ListCount > 0 Then
        m_udtProj.FileContent(lstFiles.ListIndex) = txtCode.Text
    End If
    
    SaveProject
End Sub


Private Sub SaveProject()
    Dim fh As Integer
    
    If Len(m_strPrjFile) > 0 Then
        If Len(Dir(m_strPrjFile)) > 0 Then Kill m_strPrjFile
        
        fh = FreeFile()
        Open m_strPrjFile For Binary Access Write As #fh
        Put #fh, , PRJ_MAGIC
        Put #fh, , m_udtProj
        Close #fh
    End If
End Sub


Private Sub mnuSubsysCUI_Click()
    mnuSubsysCUI.Checked = True
    mnuSubsysGUI.Checked = False
    m_udtProj.Subsystem = Subsystem_CUI
End Sub


Private Sub mnuSubsysGUI_Click()
    mnuSubsysCUI.Checked = False
    mnuSubsysGUI.Checked = True
    m_udtProj.Subsystem = Subsystem_GUI
End Sub


Private Sub picFiles_Resize()
    On Error Resume Next
    
    lstFiles.Top = 0
    lstFiles.Left = 0
    
    lstFiles.Width = picFiles.ScaleWidth
    lstFiles.Height = picFiles.ScaleHeight - cmdFileAdd.Height - 6 * Screen.TwipsPerPixelY
    
    cmdFileAdd.Top = lstFiles.Height + 3 * Screen.TwipsPerPixelY
    cmdFileRem.Top = cmdFileAdd.Top
End Sub
