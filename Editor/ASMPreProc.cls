VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASMPreprocessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ASM Source Preprocessor
'
' Arne Elster 2007 / 2008


' Quick'n'Dirty. Heavily expensive and unefficient.
' Searchs for EQUs and replaces.
Public Function Process(ByVal strASM As String) As String
    Dim strLines()  As String
    Dim strParts()  As String
    Dim strOutput   As String
    Dim strID       As String
    Dim strSymb()   As String
    Dim strCnt()    As String
    Dim lngSymbCnt  As Long
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim blnInQt     As Boolean
    Dim lngWordCnt  As Long
    Dim blnInWS     As Boolean
    Dim blnIsEQU    As Boolean
    
    strLines = Split(strASM, vbCrLf)
    
    For i = 0 To UBound(strLines)
        blnInQt = False
        blnInWS = False
        blnIsEQU = False
        lngWordCnt = 0
        
        For j = 1 To Len(strLines(i)) - 3
            If Mid$(strLines(i), j, 1) = """" Then
                blnInQt = Not blnInQt
            Else
                If Not blnInQt Then
                    If UCase$(Mid$(strLines(i), j, 3)) = "EQU" Then
                        ReDim Preserve strSymb(lngSymbCnt)
                        ReDim Preserve strCnt(lngSymbCnt)
                        strSymb(lngSymbCnt) = Trim$(Left$(strLines(i), j - 1))
                        strCnt(lngSymbCnt) = Trim$(Mid$(strLines(i), j + 3))
                        lngSymbCnt = lngSymbCnt + 1
                        blnIsEQU = True
                        Exit For
                    ElseIf Mid$(strLines(i), j, 1) = " " Then
                        If Not blnInWS Then
                            lngWordCnt = lngWordCnt + 1
                            blnInWS = True
                        End If
                    Else
                        blnInWS = False
                    End If
                End If
            End If
            If lngWordCnt > 1 Then Exit For
            DoEvents
        Next
        
        If Not blnIsEQU Then
            For j = 1 To Len(strLines(i))
                Select Case Mid$(strLines(i), j, 1)
                    Case "A" To "Z", "a" To "z":
                        strID = ""
                        Do
                            strID = strID & Mid$(strLines(i), j, 1)
                            j = j + 1
                            DoEvents
                        Loop While Mid$(strLines(i), j, 1) Like "[a-zA-Z0-9_]"
                        j = j - 1
                        
                        For k = 0 To lngSymbCnt - 1
                            If StrComp(strSymb(k), strID, vbTextCompare) = 0 Then
                                strOutput = strOutput & strCnt(k)
                                Exit For
                            End If
                        Next
                        
                        If k = lngSymbCnt Then
                            strOutput = strOutput & strID
                        End If
                        
                    Case "0" To "9":
                        strID = ""
                        Do
                            strID = strID & Mid$(strLines(i), j, 1)
                            j = j + 1
                            DoEvents
                        Loop While Mid$(strLines(i), j, 1) Like "[0-9]"
                        j = j - 1
                        
                        strOutput = strOutput & strID
                        
                    Case """":
                        strID = """"
                        Do
                            j = j + 1
                            strID = strID & Mid$(strLines(i), j, 1)
                            DoEvents
                        Loop While Mid$(strLines(i), j, 1) <> """"
                        
                        strOutput = strOutput & strID
                    
                    Case "#":
                        If j = 1 Then
                            strOutput = strOutput & strLines(i)
                            Exit For
                        Else
                            strOutput = strOutput & "#"
                        End If
                    
                    Case Else:
                        strOutput = strOutput & Mid$(strLines(i), j, 1)
                        
                End Select
                DoEvents
            Next
            strOutput = strOutput & vbCrLf
        End If
        DoEvents
    Next
    
    Process = strOutput
End Function
