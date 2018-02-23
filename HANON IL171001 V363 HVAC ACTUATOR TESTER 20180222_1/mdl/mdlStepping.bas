Attribute VB_Name = "mdlStepping"
Option Explicit

Public Function StepOpen() As Boolean
    Dim i As Integer
    Dim bRes(1) As Boolean
    
    If LOCALTEST Or STEPUSE = False Then
        StepOpen = True
        
        Exit Function
    End If
    
    If SysVar.nStepPort(0) <> 0 Then
        frmMain.comStepping1.CommPort = SysVar.nStepPort(0)
        
        If frmMain.comStepping1.PortOpen = False Then
            frmMain.comStepping1.PortOpen = True
            
            Call Delay(20)
            
            bRes(0) = frmMain.comStepping1.PortOpen
        End If
    Else
        bRes(0) = True
    End If
    
    If SysVar.nStepPort(1) <> 0 Then
        frmMain.comStepping2.CommPort = SysVar.nStepPort(1)
        
        If frmMain.comStepping2.PortOpen = False Then
            frmMain.comStepping2.PortOpen = True
            
            Call Delay(20)
            
            bRes(1) = frmMain.comStepping2.PortOpen
        End If
    Else
        bRes(1) = True
    End If
    
    If bRes(0) And bRes(1) Then
        StepOpen = True
    Else
        StepOpen = False
    End If
End Function

Public Sub StepClose()
    If LOCALTEST Or STEPUSE = False Then Exit Sub
    
    If SysVar.nStepPort(0) <> 0 Then
        If frmMain.comStepping1.PortOpen = True Then
            frmMain.comStepping1.PortOpen = False
        End If
        
        Call Delay(20)
    End If
    
    If SysVar.nStepPort(1) <> 0 Then
        If frmMain.comStepping2.PortOpen = True Then
            frmMain.comStepping2.PortOpen = False
        End If
        
        Call Delay(20)
    End If
End Sub

Public Sub Step1Received()
    Dim i As Integer
    Dim A As Variant
    Dim B As String
    Dim lRes(1) As Long
    
    If LOCALTEST Or STEPUSE = False Then Exit Sub
    
    A = frmMain.comStepping1.Input
    
    For i = 0 To UBound(A)
        ' 입력된 문자를 버퍼에 쌓아 둔다.
        B = Dec2Hex(CStr(A(i)))
        lpSteppingData(0) = lpSteppingData(0) & B
    Next
    
    ' 1  2  3  4  5  6  7  8  9  0
    ' 12 34 56 78 90 12 34 56 78 90
    ' 00 00 00 00 00 00 00 00 00 00
    If Left(lpSteppingData(0), 2) = STEP_START And Right(lpSteppingData(0), 2) = STEP_END And Mid(lpSteppingData(0), 3, 2) = STEP_READ Then
        If STEPLOGUSE Then
            Call OnLog("[STEP] READ 1 : " & lpSteppingData(0))
        End If
        
        ' ACT01
        lpReadStepData(0) = Mid(lpSteppingData(0), 9, 2)
        lpReadStepData(1) = Mid(lpSteppingData(0), 11, 2)
        bActStallBit(0) = IIf(Mid(lpSteppingData(0), 7, 2) = "01", True, False)
        
        lRes(0) = Val("&H" & lpReadStepData(1)) * 256
        lRes(1) = Val("&H" & lpReadStepData(0))
        
        lStepActData(0) = lRes(0) + lRes(1)
        
        ' ACT02
        lpReadStepData(2) = Mid(lpSteppingData(0), 15, 2)
        lpReadStepData(3) = Mid(lpSteppingData(0), 17, 2)
        bActStallBit(1) = IIf(Mid(lpSteppingData(0), 13, 2) = "01", True, False)
        
        lRes(0) = Val("&H" & lpReadStepData(3)) * 256
        lRes(1) = Val("&H" & lpReadStepData(2))
        
        lStepActData(1) = lRes(0) + lRes(1)
        
        lpSteppingData(0) = ""
    End If
End Sub

Public Sub Step2Received()
    Dim i As Integer
    Dim A As Variant
    Dim B As String
    Dim lRes(1) As Long
    
    If LOCALTEST Or STEPUSE = False Then Exit Sub
    
    A = frmMain.comStepping2.Input
    
    For i = 0 To UBound(A)
        ' 입력된 문자를 버퍼에 쌓아 둔다.
        B = Dec2Hex(CStr(A(i)))
        lpSteppingData(1) = lpSteppingData(1) & B
    Next
    
    ' 12 34 56 78 90 12 34 56 78 90
    ' 00 00 00 00 00 00 00 00 00 00
    If Left(lpSteppingData(1), 2) = STEP_START And Right(lpSteppingData(1), 2) = STEP_END And Mid(lpSteppingData(1), 3, 2) = STEP_READ Then
        If STEPLOGUSE Then
            Call OnLog("[STEP] READ 2 : " & lpSteppingData(1))
        End If
        
        ' ACT03
        lpReadStepData(4) = Mid(lpSteppingData(1), 9, 2)
        lpReadStepData(5) = Mid(lpSteppingData(1), 11, 2)
        bActStallBit(2) = IIf(Mid(lpSteppingData(1), 7, 2) = "01", True, False)
        
        lRes(0) = Val("&H" & lpReadStepData(5)) * 256
        lRes(1) = Val("&H" & lpReadStepData(4))
        
        lStepActData(2) = lRes(0) + lRes(1)
        
        ' ACT04
        lpReadStepData(6) = Mid(lpSteppingData(1), 15, 2)
        lpReadStepData(7) = Mid(lpSteppingData(1), 17, 2)
        bActStallBit(3) = IIf(Mid(lpSteppingData(1), 13, 2) = "01", True, False)
        
        lRes(0) = Val("&H" & lpReadStepData(7)) * 256
        lRes(1) = Val("&H" & lpReadStepData(6))
        
        lStepActData(3) = lRes(0) + lRes(1)
        
        lpSteppingData(1) = ""
    End If
End Sub

Public Sub StepStartInfo()
    Dim i As Integer
    Dim nCount As Integer
    Dim lpCmd As String
    Dim iCmd(300) As Integer
    
    If LOCALTEST Or STEPUSE = False Then Exit Sub
    
    frmRun.tmrSteppingSend.Enabled = False
    
    Call Delay(500)
    
    On Error Resume Next
    
    ' port 1
    nCount = 0
    
    iCmd(nCount) = Val("&H" & STEP_START)
    nCount = nCount + 1
    
    iCmd(nCount) = Val("&H" & STEP_INFO)
    nCount = nCount + 1
    
    iCmd(nCount) = &H0
    nCount = nCount + 1
    
    For i = 0 To 15
        iCmd(nCount) = Val2Byte(CM_HI, SetupVar.nSteppingSet(0, i))
        nCount = nCount + 1
        
        iCmd(nCount) = Val2Byte(CM_LO, SetupVar.nSteppingSet(0, i))
        nCount = nCount + 1
    Next
    
    For i = 0 To 15
        iCmd(nCount) = Val2Byte(CM_HI, SetupVar.nSteppingSet(1, i))
        nCount = nCount + 1
        
        iCmd(nCount) = Val2Byte(CM_LO, SetupVar.nSteppingSet(1, i))
        nCount = nCount + 1
    Next
    
    For i = 0 To 15
        iCmd(nCount) = Val2Byte(CM_HI, "0000")
        nCount = nCount + 1
        iCmd(nCount) = Val2Byte(CM_LO, "0000")
        nCount = nCount + 1
    Next
    
    For i = 0 To 15
        iCmd(nCount) = Val2Byte(CM_HI, "0000")
        nCount = nCount + 1
        iCmd(nCount) = Val2Byte(CM_LO, "0000")
        nCount = nCount + 1
    Next
    
    iCmd(nCount) = &H0
    nCount = nCount + 1
    
    iCmd(nCount) = Val("&H" & STEP_END)
    
    lpCmd = ""
    
    For i = 0 To nCount
        lpCmd = lpCmd & Dec2Hex(CStr(iCmd(i)))
    Next
    
    Call OnLog("[STEP] SEND PORT " & SysVar.nStepPort(0) & " : " & lpCmd)
    
    If frmMain.comStepping1.PortOpen Then
        frmMain.comStepping1.Output = iCmd
    Else
        Call OnLog("[STEP] SEND PORT " & SysVar.nStepPort(0) & " : DUMMY")
    End If
    
    Call Delay(500)
    
    Call OnLog("[STEP] SEND PORT " & SysVar.nStepPort(1) & " : " & lpCmd)
    
    If frmMain.comStepping1.PortOpen Then
        frmMain.comStepping1.Output = iCmd
    Else
        Call OnLog("[STEP] SEND PORT " & SysVar.nStepPort(1) & " : DUMMY")
    End If
    
    Call Delay(500)
    
    frmRun.tmrSteppingSend.Enabled = True
    
    Exit Sub
    
    ' port 2
    nCount = 0
    
    iCmd(nCount) = Val("&H" & STEP_START)
    nCount = nCount + 1
    
    iCmd(nCount) = Val("&H" & STEP_INFO)
    nCount = nCount + 1
    
    iCmd(nCount) = &H0
    nCount = nCount + 1
    
    For i = 0 To 15
        iCmd(nCount) = Val2Byte(CM_HI, SetupVar.nSteppingSet(2, i))
        nCount = nCount + 1
        
        iCmd(nCount) = Val2Byte(CM_LO, SetupVar.nSteppingSet(2, i))
        nCount = nCount + 1
    Next
    
    For i = 0 To 15
        iCmd(nCount) = Val2Byte(CM_HI, SetupVar.nSteppingSet(3, i))
        nCount = nCount + 1
        
        iCmd(nCount) = Val2Byte(CM_LO, SetupVar.nSteppingSet(3, i))
        nCount = nCount + 1
    Next
    
    For i = 0 To 15
        iCmd(nCount) = Val2Byte(CM_HI, "0000")
        nCount = nCount + 1
        iCmd(nCount) = Val2Byte(CM_LO, "0000")
        nCount = nCount + 1
    Next
    
    For i = 0 To 15
        iCmd(nCount) = Val2Byte(CM_HI, "0000")
        nCount = nCount + 1
        iCmd(nCount) = Val2Byte(CM_LO, "0000")
        nCount = nCount + 1
    Next
    
    iCmd(nCount) = &H0
    nCount = nCount + 1
    
    iCmd(nCount) = Val("&H" & STEP_END)
    
    lpCmd = ""
    
    For i = 0 To nCount
        lpCmd = lpCmd & Dec2Hex(CStr(iCmd(i)))
    Next
    
    Call OnLog("[STEP] SEND PORT " & SysVar.nStepPort(1) & " : " & lpCmd)
    
    If frmMain.comStepping2.PortOpen Then
        frmMain.comStepping2.Output = iCmd
    Else
        Call OnLog("[STEP] SEND PORT " & SysVar.nStepPort(1) & " : DUMMY")
    End If
    
    Call Delay(500)
    
    frmRun.tmrSteppingSend.Enabled = True
End Sub

Public Function Val2Byte(ByVal HiLo As Byte, ByVal nData As Integer) As Byte
    Dim bRes As Byte
    
    If HiLo = CM_LO Then
        bRes = Int(nData / 256)
    Else
        bRes = nData Mod 256
    End If
    
    Val2Byte = bRes
End Function

Public Function Dec2Hex(tmpStr As String) As String ' return 4 digit string
    Dim A As String
    
    A = Hex(tmpStr)
    
    Select Case Len(A)
        Case 0: Dec2Hex = "00"
        Case 1: Dec2Hex = "0" & A
        Case 2: Dec2Hex = A
    End Select
End Function

Public Function StepSend(ByVal lpStr As String, ByVal nPortNo As Integer) As String
    Dim i As Integer
    Dim nRes() As Byte
    
    If LOCALTEST Or STEPUSE = False Or nPortNo > 1 Then
        Exit Function
    End If
    
    Call OnLog("[STEP] PORT " & SysVar.nStepPort(nPortNo) & " : " & lpStr)
    
    ReDim nRes((Len(lpStr) / 2) - 1)
    
    For i = 0 To UBound(nRes)
        nRes(i) = Val("&H" & Mid(lpStr, (i * 2) + 1, 2))
    Next
    
    Select Case nPortNo
        Case 0: frmMain.comStepping1.Output = nRes
        Case 1: frmMain.comStepping2.Output = nRes
    End Select
    
    StepSend = lpStr
End Function

Public Function CalcStep(ByVal lData As Long, ByVal nCh As Integer, ByVal nPos As Integer) As Long
    Dim i As Integer
    Dim j As Integer
    Dim dOldVolt As Double
    Dim bLastPos As Boolean
    
    If nPos < 0 Then
        nPos = 0
    End If
    
    If nPos = 999 Then
        Exit Function
    End If
    
    dOldVolt = 0
    
    If nPos > 0 Then
        Select Case nCh
            Case 0:
                dOldVolt = Val(frmRun.pnlAct01Volt(nPos).Caption)
                
                j = 0
                
                For i = nPos To 1 Step -1
                    j = j + 1
                    
                    If dOldVolt < 1 Then
                        dOldVolt = Val(frmRun.pnlAct01Volt(nPos - j).Caption)
                    End If
                Next
            Case 1:
                dOldVolt = Val(frmRun.pnlAct02Volt(nPos).Caption)
                
                j = 0
                
                For i = nPos To 1 Step -1
                    j = j + 1
                    
                    If dOldVolt < 1 Then
                        dOldVolt = Val(frmRun.pnlAct02Volt(nPos - j).Caption)
                    End If
                Next
            Case 2:
                dOldVolt = Val(frmRun.pnlAct03Volt(nPos).Caption)
                
                j = 0
                
                For i = nPos To 1 Step -1
                    j = j + 1
                    
                    If dOldVolt < 1 Then
                        dOldVolt = Val(frmRun.pnlAct03Volt(nPos - j).Caption)
                    End If
                Next
            Case 3:
                dOldVolt = Val(frmRun.pnlAct04Volt(nPos).Caption)
                
                j = 0
                
                For i = nPos To 1 Step -1
                    j = j + 1
                    
                    If dOldVolt < 1 Then
                        dOldVolt = Val(frmRun.pnlAct04Volt(nPos - j).Caption)
                    End If
                Next
        End Select
    End If
    
    bLastPos = False
    
    Select Case nCh
        Case 0: If nPos = (RunVar.nAct01MaxLoop - 1) Then bLastPos = True
        Case 1: If nPos = (RunVar.nAct02MaxLoop - 1) Then bLastPos = True
        Case 2: If nPos = (RunVar.nAct03MaxLoop - 1) Then bLastPos = True
        Case 3: If nPos = (RunVar.nAct04MaxLoop - 1) Then bLastPos = True
    End Select
    
    If bLastPos Then
        CalcStep = Abs(lData - CLng(dOldVolt))
    Else
        CalcStep = Abs(lData + CLng(dOldVolt))
    End If
End Function

