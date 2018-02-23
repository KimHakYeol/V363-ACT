Attribute VB_Name = "mdlLib"
Option Explicit


Dim cFrequency As Currency
Dim dOldTime(99) As Double

Public Function InitPerformanceLevelUp() As Boolean
    Dim bRes As Boolean
    
    bRes = True
    
    If QueryPerformanceFrequencyAny(cFrequency) = 0 Then
        Call MsgBox("This computer doesn't support high-res timers", vbCritical)
        bRes = False
    End If
    
    InitPerformanceLevelUp = bRes
End Function

Public Static Sub Delay(ByVal dTime As Double)
    Static dGetTime     As Double
    Static bTimeEscape  As Boolean
    
    If dTime <> 0# Then dTime = dTime * 0.001
    
    bTimeEscape = False
    dGetTime = GetSystemTime()
    
    While bTimeEscape = False
        DoEvents
        
        If (GetSystemTime() - dGetTime) >= dTime Then
            bTimeEscape = True
        End If
    Wend
End Sub

Private Static Function GetSystemTime() As Double
    Static cSec As Currency
    
    Call QueryPerformanceCounterAny(cSec)
    GetSystemTime = CDbl(cSec) / CDbl(cFrequency)
End Function

Public Sub ClearTime(ByVal nNo As Integer)
    dOldTime(nNo) = 0
End Sub

Public Sub SetTime(ByVal nNo As Integer)
    dOldTime(nNo) = GetSystemTime
End Sub

Private Function GetTime(ByVal nNo As Integer) As Double
    GetTime = dOldTime(nNo)
End Function

Public Function ElapseTime(ByVal nNo As Integer) As Double
    ElapseTime = GetSystemTime - dOldTime(nNo)
End Function

Public Sub CheckKeyBoardDisplay(ByRef refTextBox As TextBox, ByVal isNumber As Boolean)
    If SysVar.bScreenKeyboardUse = True Then
        Set KeyBoardObj = refTextBox
        bKeyBoardNum = isNumber
        frmKeyboard.Show vbModal
    End If
End Sub

Public Sub ModelNameToCombo(ByRef refCombo As ComboBox, ByRef refVar As TagSetupVariables)
    Dim lpRes As String
    
    If refCombo.ListCount = 0 Then
        If SETUPSELECT Then
            lpRes = SelectCar(0).ModelNameSub(0)
            
            If SelectCar(0).ModelName <> "" And lpRes <> "" Then
                lpNowModel = Format(nNowModelNo, "0000") & "_" & SelectCar(0).ModelName & "_" & lpRes
            Else
                lpNowModel = "0000_DEFAULT"
            End If
        Else
            lpNowModel = "0000_DEFAULT"
        End If
        
        refCombo.AddItem UCase(Trim$(lpNowModel))
        Call MsgBox("Default file select.", vbOKOnly, "Information")
    End If
End Sub

Public Function SearchModelNo(ByRef refCombo As ComboBox, ByVal lpModelName As String) As Integer
    Dim nRes    As Integer
    Dim i       As Integer
    
    nRes = RT_ERROR
    
    For i = 0 To refCombo.ListCount - 1
        If Trim$(lpModelName) = Trim$(refCombo.List(i)) Then
            nRes = i
            Exit For
        End If
    Next
    
    SearchModelNo = nRes
End Function

Public Function ModelChange(ByRef refCbo As ComboBox, ByVal lpFileName As String) As Boolean
    Dim nModelNo As Integer
    Dim bRes As Boolean

RELOAD:

    bRes = LoadSetupFile(lpFileName)
    
    refCbo.Clear
    
    If bRes Then
        Call INIModelListRead(refCbo)
        nModelNo = SearchModelNo(refCbo, Trim$(lpFileName))
    Else
        nModelNo = RT_ERROR
    End If
    
    If nModelNo = RT_ERROR Then
        Call SetupDisp2Mem
        Call SaveSetupFile(Trim$(lpNowModel))
        
        GoTo RELOAD
    End If
    
    refCbo.ListIndex = nModelNo
    ModelChange = True
End Function

Public Function MakeFolder() As String
    Dim lpFolder As String
    Call SearchDir(lpPath & "\\SetupFile")
    Call SearchDir(lpPath & "\\DataFile")
    Call SearchDir(lpPath + "\\DataFile\\" + Format(Date, "yyyy"))
    
    lpFolder = lpPath + "\\DataFile\\" + Format(Date, "yyyy") + "\\" + Format(Date, "yyyymm")
    Call SearchDir(lpFolder)
    Call SearchDir(lpFolder & "\\NGFiles")
    
    MakeFolder = lpFolder
End Function

Public Sub SearchDir(ByVal lpDirName As String)
    If Dir(lpDirName, vbDirectory) = "" Then
        Call MkDir(lpDirName)
    End If
End Sub

Public Function SearchFile(ByVal lpFileName As String) As Boolean
    If Dir(lpFileName, vbNormal) = "" Then
        SearchFile = False
    Else
        SearchFile = True
    End If
End Function

Public Sub ActBoardChange()
    Dim i As Integer
    
    For i = 0 To UBound(ActNo)
        Select Case SetupVar.nActBoardNo(i)
            Case 1:
                ActNo(i).AD_CURR = AD_ACT01_CURR
                ActNo(i).AD_VOLT = AD_ACT01_VOLT
                ActNo(i).O_POWER = O_ACT01_POWER
                ActNo(i).DA_NO = DA_ACT01
            
            Case 2:
                ActNo(i).AD_CURR = AD_ACT02_CURR
                ActNo(i).AD_VOLT = AD_ACT02_VOLT
                ActNo(i).O_POWER = O_ACT02_POWER
                ActNo(i).DA_NO = DA_ACT02
            
            Case 3:
                ActNo(i).AD_CURR = AD_ACT03_CURR
                ActNo(i).AD_VOLT = AD_ACT03_VOLT
                ActNo(i).O_POWER = O_ACT03_POWER
                ActNo(i).DA_NO = DA_ACT03
            
            Case 4:
                ActNo(i).AD_CURR = AD_ACT04_CURR
                ActNo(i).AD_VOLT = AD_ACT04_VOLT
                ActNo(i).O_POWER = O_ACT04_POWER
                ActNo(i).DA_NO = DA_ACT04
        
        End Select
    Next
End Sub

Private Sub SWAP(ByRef A As Double, ByRef B As Double)
    Dim c As Double
    
    c = A
    A = B
    B = c
End Sub

Public Sub DataSort(ByRef dBuf() As Double, ByVal nCount As Integer)
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To nCount
        For j = i To nCount
            If dBuf(i) < dBuf(j) Then
                Call SWAP(dBuf(i), dBuf(j))
            End If
        Next
    Next
End Sub
