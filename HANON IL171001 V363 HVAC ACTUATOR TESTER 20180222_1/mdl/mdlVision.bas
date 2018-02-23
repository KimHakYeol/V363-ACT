Attribute VB_Name = "mdlVision"
Option Explicit

Public Function VisionOpen() As Boolean
    If VISIONUSE = False Then
        VisionOpen = True
        
        Exit Function
    End If
    
    If SysVar.nVisionPort = 0 Then
        Call MsgBox("VISION PORT CHECK...")
        
        VisionOpen = False
        
        Exit Function
    End If
    
    frmMain.comVision.CommPort = SysVar.nVisionPort
    
    If frmMain.comVision.PortOpen = False Then
        frmMain.comVision.PortOpen = True
    End If
    
    Call Delay(20)
    
    VisionOpen = True
End Function

Public Sub VisionClose()
    If VISIONUSE = False Then Exit Sub
    
    If frmMain.comVision.PortOpen = True Then frmMain.comVision.PortOpen = False
    
    Call Delay(20)
End Sub

Public Sub VisionReceived()
    Dim bRes As Boolean
    Dim A As String
    Dim B As Integer
    
    If VISIONUSE = False Then Exit Sub
    
    A = frmMain.comVision.Input
    
    If A <> "" Then
        B = Asc(A)
        lpVisionCom = lpVisionCom & A
    End If
End Sub

Public Function VisionSend(ByVal lpModelName As String) As String
    Dim i As Integer
    Dim lpStr As String
    Dim nInt(1) As Integer
    Dim nRes(1) As Integer
    
    lpVisionCom = ""
    
    lpStr = Left(lpModelName, 4)
    nInt(0) = Val(Left(lpStr, 2))
    nInt(1) = Val(Right(lpStr, 2))
    
    If nInt(0) = 0 Or nInt(1) = 0 Then
        Call OnLog("[VISION] NUMBER ERROR...")
        
        Exit Function
    End If
    
    lpModelName = "START" & Format(CStr(nInt(0)), "00") & Format(CStr(nInt(1)), "00") & "A" & Format(SysVar.lOkCounter + 1, "0000000")
    
    If VISIONUSE Then
        frmMain.comVision.Output = Chr$(2) & lpModelName & Chr$(3)
    End If
    
    VisionSend = lpModelName
End Function

Public Sub VisionDoorSend(ByVal lpCmd As String)
    If VISIONUSE = False Then Exit Sub
    
    lpVisionCom = frmMain.comVision.Input
    lpVisionCom = ""
    
    frmMain.comVision.Output = Chr$(2) & lpCmd & Chr$(3)
End Sub

Public Sub VisionFinalSend()
    If VISIONUSE = False Then Exit Sub
    If SetupVar.bVisionUse = False Then Exit Sub
    
    lpVisionCom = ""
    
    frmMain.comVision.Output = Chr$(2) & "END" & Chr$(3)
End Sub

Public Function VisionTest(Optional ByVal bResult As Boolean = False) As Boolean
    Dim lBkColor    As Long
    Dim bRes        As Boolean
    Dim lpLog       As String
    Dim bSignal     As Boolean
    Dim lpOK As String
    Dim lpNG As String
    
    bRes = False
    
    If RunVar.nVisionPos = POS_INIT Then
        RunVar.nVisionReCount = 0
        RunVar.bVisionAck = False
        RunVar.nVisionPos = POS_RUN_INIT
    End If
    
    If RunVar.nVisionPos = POS_RUN_INIT Then
        lpLog = VisionSend(Trim$(frmRun.cboCarType.Text))
        
        Call OnLog("VISION SEND : " & lpLog)
        Call SetTime(TM_VISION)
        
        RunVar.nVisionPos = POS_RUN
    End If
    
    If RunVar.nVisionPos = POS_RUN Then
        lpOK = Chr(2) & "RUNNING" & Chr(3) & Chr(2) & "OK" & Chr(3)
        lpNG = Chr(2) & "RUNNING" & Chr(3) & Chr(2) & "NG" & Chr(3)
        
        If ElapseTime(TM_VISION) > 10 Then
            Call OnLog("[ERROR] VISION TEST TIME OVER !!!")
            lBkColor = vbRed
            RunVar.nVisionPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "RUNNING" Then
            bSignal = RunVar.bVisionAck
            RunVar.bVisionAck = True
            
            If bSignal <> RunVar.bVisionAck Then
                lBkColor = vbWhite
                Call OnLog("VISION ACK...")
            End If
        End If
        
        Select Case Trim$(lpVisionCom)
            Case ("RT" & Format((Left((Trim$(frmRun.cboCarType.Text)), 3)), "000")):
                lBkColor = vbRed
                RunVar.nVisionPos = POS_END
            
            Case lpOK:
                lBkColor = vbGreen
                RunVar.nVisionPos = POS_END
            
            Case lpNG:
                lBkColor = vbRed
                RunVar.nVisionPos = POS_END
        
        End Select
    End If
    
    If RunVar.nVisionPos = POS_END Then
        Call VisionResult(lBkColor)
        
        RunVar.bVisionUse = False
        RunVar.nVisionPos = 0
        bRes = True
    End If
    
    VisionTest = bRes
End Function

Public Sub ManualVision(ByVal nCamera As Integer, ByVal bOpen As Integer, ByVal nVarPos As Integer, ByVal AdCh As Integer)
    Dim nTimeNum As Integer
    Dim nRes As Integer
    Dim nActPeakCurrCnt As Integer
    Dim nActEndPosCnt As Integer
    Dim nPeakCurrCnt As Integer
    Dim nEndPosCnt As Integer
    
    Select Case AdCh
        Case ActNo(0).AD_CURR:
            nTimeNum = TM_ACT01
            nActPeakCurrCnt = SetupVar.nActPeakCurrCount(0)
            nActEndPosCnt = SetupVar.nActEndPosCount(0)
        
        Case ActNo(1).AD_CURR:
            nTimeNum = TM_ACT02
            nActPeakCurrCnt = SetupVar.nActPeakCurrCount(1)
            nActEndPosCnt = SetupVar.nActEndPosCount(1)
        
        Case ActNo(2).AD_CURR:
            nTimeNum = TM_ACT03
            nActPeakCurrCnt = SetupVar.nActPeakCurrCount(2)
            nActEndPosCnt = SetupVar.nActEndPosCount(2)
        
        Case ActNo(3).AD_CURR:
            nTimeNum = TM_ACT04
            nActPeakCurrCnt = SetupVar.nActPeakCurrCount(3)
            nActEndPosCnt = SetupVar.nActEndPosCount(3)
        
        Case Else:
            Exit Sub
    
    End Select
    
    nRes = IIf(bOpen, SetupVar.nOpenCameraPos(nVarPos), SetupVar.nCloseCameraPos(nVarPos))
    
    If bOpen Then
        frmRun.pnlVisionDoorOpen(nCamera).Caption = "M"
        frmRun.pnlVisionDoorOpen(nCamera).BackColor = vbYellow
    Else
        frmRun.pnlVisionDoorClose(nCamera).Caption = "M"
        frmRun.pnlVisionDoorClose(nCamera).BackColor = vbYellow
    End If
    
    frmRun.optActManual(nRes).Value = True
    
    nPeakCurrCnt = 0
    nEndPosCnt = 0
    
    Call SetTime(nTimeNum)
    Do
        DoEvents
        
        If ElapseTime(nTimeNum) > 0.3 Then
            
            ' manual vision
            
            If ElapseTime(nTimeNum) > 10# Then
                frmRun.pnlVisionDoorOpen(nCamera).Caption = "E"
                frmRun.pnlVisionDoorOpen(nCamera).BackColor = vbRed
                
                Exit Do
            End If
            
            If RunVar.bAutoManual <> DIS(I_AUTO_SW) Then
                Exit Do
            End If
            
            If DIS(I_STOP_SW) Then
                Exit Do
            End If
        End If
    Loop
End Sub

Public Sub VisionResult(ByVal lResult As Long)
    If lResult = vbGreen Then
        frmRun.pnlVisionResult.Caption = "OK"
    Else
        frmRun.pnlVisionResult.Caption = "NG"
        
        RunVar.bReVisionUse = True
        RunVar.bFinal = False
    End If
    
    frmRun.pnlVisionResult.BackColor = lResult
End Sub

Public Function OnVisionAct01Door(ByVal nPos As Integer) As Boolean
    Dim bRes As Boolean
    
    bRes = False
    
    If RunVar.bVisionAct01DoorMarking(nPos) Then
        RunVar.nVisionAct01OnDoorPos = POS_END
    End If
    
    If RunVar.nVisionAct01OnDoorPos = POS_INIT Then
        RunVar.bVisionAct01DoorExit = False
        RunVar.nVisionAct01DoorLoop = 0 ' 한 포지션에서 두개 이상이 겹질 경우를 위해
        
        If nPos = SetupVar.nOpenCameraPos(0) Then RunVar.nVisionAct01DoorLoop = RunVar.nVisionAct01DoorLoop + &H1
        If nPos = SetupVar.nCloseCameraPos(0) Then RunVar.nVisionAct01DoorLoop = RunVar.nVisionAct01DoorLoop + &H2
        
        RunVar.nVisionAct01OnDoorPos = POS_RUN_INIT
    End If
    
    If RunVar.nVisionAct01OnDoorPos = POS_RUN_INIT Then
        If RunVar.nVisionAct01DoorLoop = 0 Then
            RunVar.nVisionAct01OnDoorPos = POS_END
        Else
'            Call OnLog("START VISION DOOR COUNTER : 0x" & Hex(RunVar.nVisionAct01DoorLoop))
            
            RunVar.nVisionAct01DoorPos = POS_INIT
            RunVar.nVisionAct01OnDoorPos = POS_RUN
        End If
    End If
    
    If RunVar.nVisionAct01OnDoorPos = POS_RUN Then
        Select Case RunVar.nVisionAct01DoorLoop
            Case &H1:
                bRes = VisionAct01DoorTest(0, True)
            
            Case &H2:
                bRes = VisionAct01DoorTest(0, False)
            
            Case Else:
                RunVar.nVisionAct01OnDoorPos = POS_END
        
        End Select
        
        If bRes Then
            If RunVar.nVisionAct01DoorLoop = 0 Then
                RunVar.nVisionAct01OnDoorPos = POS_END
            Else
                RunVar.nVisionAct01DoorDone = (RunVar.nVisionAct01DoorDone Or RunVar.nVisionAct01DoorReturn) ' 테스트 한것은 테스트 하지 않기 위해 테스트 프래그를 설정한다.
                RunVar.nVisionAct01DoorLoop = RunVar.nVisionAct01DoorLoop - (RunVar.nVisionAct01DoorLoop And RunVar.nVisionAct01DoorDone)   ' 테스트 한것은 테스트 하지 않는다.
                
                RunVar.nVisionAct01OnDoorPos = POS_INIT
'                Call OnLog("REMAIN VISION DOOR COUNTER : 0x" & Hex(RunVar.nVisionAct01DoorLoop))
                
                If RunVar.bVisionAct01DoorExit And RunVar.nVisionAct01DoorLoop = 0 Then
                    RunVar.nVisionAct01OnDoorPos = POS_END
                Else
                    RunVar.bVisionAct01DoorExit = False
                    RunVar.nVisionAct01OnDoorPos = POS_RUN_INIT
                End If
            End If
        End If
        
        bRes = False    ' 안해주면 안되용
    End If
    
    If RunVar.nVisionAct01OnDoorPos = POS_END Then
        RunVar.bVisionAct01DoorMarking(nPos) = True ' Marking
        bRes = True
    End If
    
    OnVisionAct01Door = bRes
End Function

Public Function VisionAct01DoorTest(ByVal nCameraNo As Integer, ByVal bOpen As Boolean) As Boolean
    Dim lBkColor As Long
    Dim bRes As Boolean
    Dim bSignal As Boolean
    Dim lpCmd As String
    
    bRes = False
    
    If RunVar.nVisionAct01DoorPos = POS_INIT Then
        RunVar.bVisionAct01DoorRunning = True
        RunVar.nVisionAct01DoorReCount = 0
        RunVar.bVisionAct01DoorAck = False
        RunVar.nVisionAct01DoorPos = POS_RUN_INIT
    End If
    
    If RunVar.nVisionAct01DoorPos = POS_RUN_INIT Then
        If bOpen Then
            lpCmd = "OP" & Format(SetupVar.nOpenCameraNo(nCameraNo) + 1, "000")
        Else
            lpCmd = "CL" & Format(SetupVar.nCloseCameraNo(nCameraNo) + 1, "000")
        End If
        
        Call VisionDoorSend(lpCmd)
        Call OnLog("VISION ACT 01 DOOR SEND:" & lpCmd)
        
        RunVar.nVisionAct01DoorPos = POS_RUN
        Call SetTime(TM_VISION)
    End If
    
    If RunVar.nVisionAct01DoorPos = POS_RUN Then
        If ElapseTime(TM_VISION) > 10 Then
            Call OnLog("[ERROR] VISION Act01 DOOR TEST TIME OVER !!!")
            lBkColor = vbRed
            RunVar.nVisionAct01DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "RUNNING" Then
            bSignal = RunVar.bVisionAct01DoorAck
            RunVar.bVisionAct01DoorAck = True
            
            If bSignal <> RunVar.bVisionAct01DoorAck Then
                lBkColor = vbWhite
                Call OnLog("VISION Act01 DOOR ACK...")
            End If
        End If
        
        If Trim$(lpVisionCom) = ("RM" & Format(nCameraNo, "000")) Then
            lBkColor = vbRed
            RunVar.nVisionAct01DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "EMOK" Then
            lBkColor = vbGreen
            RunVar.nVisionAct01DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "EMNG" Then
            lBkColor = vbRed
            RunVar.nVisionAct01DoorPos = POS_END
        End If
    End If
    
    If RunVar.nVisionAct01DoorPos = POS_END Then
        Call VisionDoorResult(nCameraNo, bOpen, lBkColor, SetupVar.nActBoardNo(0))
        
        RunVar.nVisionAct01DoorPos = 0
        RunVar.bVisionAct01DoorRunning = False
        
        If bOpen = True Then RunVar.nVisionAct01DoorReturn = &H1
        If bOpen = False Then RunVar.nVisionAct01DoorReturn = &H2
        
        RunVar.bVisionDoorManualOpenUse(nCameraNo) = False
        RunVar.bVisionDoorManualCloseUse(nCameraNo) = False
        
        RunVar.bVisionAct01DoorExit = True
        bRes = True
    End If
    
    VisionAct01DoorTest = bRes
End Function

Public Function OnVisionAct02Door(ByVal nPos As Integer) As Boolean
    Dim bRes As Boolean
    
    bRes = False
    
    If RunVar.bVisionAct02DoorMarking(nPos) Then
        RunVar.nVisionAct02OnDoorPos = POS_END
    End If
    
    If RunVar.nVisionAct02OnDoorPos = POS_INIT Then
        RunVar.bVisionAct02DoorExit = False
        RunVar.nVisionAct02DoorLoop = 0 ' 한 포지션에서 두개 이상이 겹질 경우를 위해
        
        If nPos = SetupVar.nOpenCameraPos(1) Then RunVar.nVisionAct02DoorLoop = RunVar.nVisionAct02DoorLoop + &H1
        If nPos = SetupVar.nCloseCameraPos(1) Then RunVar.nVisionAct02DoorLoop = RunVar.nVisionAct02DoorLoop + &H2
        
        RunVar.nVisionAct02OnDoorPos = POS_RUN_INIT
    End If
    
    If RunVar.nVisionAct02OnDoorPos = POS_RUN_INIT Then
        If RunVar.nVisionAct02DoorLoop = 0 Then
            RunVar.nVisionAct02OnDoorPos = POS_END
        Else
            'Call OnLog("START VISION DOOR COUNTER : 0x" & Hex(.nVisionAct02DoorLoop))
            
            RunVar.nVisionAct02DoorPos = POS_INIT
            RunVar.nVisionAct02OnDoorPos = POS_RUN
        End If
    End If
    
    If RunVar.nVisionAct02OnDoorPos = POS_RUN Then
        Select Case RunVar.nVisionAct02DoorLoop
            Case &H1:
                bRes = VisionAct02DoorTest(1, True)
            
            Case &H2:
                bRes = VisionAct02DoorTest(1, False)
            
            Case Else:
                RunVar.nVisionAct02OnDoorPos = POS_END
        
        End Select
        
        If bRes Then
            If RunVar.nVisionAct02DoorLoop = 0 Then
                RunVar.nVisionAct02OnDoorPos = POS_END
            Else
                RunVar.nVisionAct02DoorDone = (RunVar.nVisionAct02DoorDone Or RunVar.nVisionAct02DoorReturn) ' 테스트 한것은 테스트 하지 않기 위해 테스트 프래그를 설정한다.
                RunVar.nVisionAct02DoorLoop = RunVar.nVisionAct02DoorLoop - (RunVar.nVisionAct02DoorLoop And RunVar.nVisionAct02DoorDone)   ' 테스트 한것은 테스트 하지 않는다.
                
                RunVar.nVisionAct02OnDoorPos = POS_INIT
                'Call OnLog("REMAIN VISION DOOR COUNTER : 0x" & Hex(.nVisionAct02DoorLoop))
                
                If RunVar.bVisionAct02DoorExit And RunVar.nVisionAct02DoorLoop = 0 Then
                    RunVar.nVisionAct02OnDoorPos = POS_END
                Else
                    RunVar.bVisionAct02DoorExit = False
                    RunVar.nVisionAct02OnDoorPos = POS_RUN_INIT
                End If
            End If
        End If
        
        bRes = False    ' 안해주면 안되용
    End If
    
    If RunVar.nVisionAct02OnDoorPos = POS_END Then
        RunVar.bVisionAct02DoorMarking(nPos) = True ' Marking
        bRes = True
    End If
    
    OnVisionAct02Door = bRes
End Function

Public Function VisionAct02DoorTest(ByVal nCameraNo As Integer, ByVal bOpen As Boolean) As Boolean
    Dim lBkColor As Long
    Dim bRes As Boolean
    Dim bSignal As Boolean
    Dim lpCmd As String
    
    bRes = False
    
    If RunVar.nVisionAct02DoorPos = POS_INIT Then
        RunVar.bVisionAct02DoorRunning = True
        RunVar.nVisionAct02DoorReCount = 0
        RunVar.bVisionAct02DoorAck = False
        RunVar.nVisionAct02DoorPos = POS_RUN_INIT
    End If
    
    If RunVar.nVisionAct02DoorPos = POS_RUN_INIT Then
        If bOpen Then
            lpCmd = "OP" & Format(SetupVar.nOpenCameraNo(nCameraNo) + 1, "000")
        Else
            lpCmd = "CL" & Format(SetupVar.nCloseCameraNo(nCameraNo) + 1, "000")
        End If
        
        Call VisionDoorSend(lpCmd)
        Call OnLog("VISION ACT 02 DOOR SEND:" & lpCmd)
        
        RunVar.nVisionAct02DoorPos = POS_RUN
        Call SetTime(TM_VISION)
    End If
    
    If RunVar.nVisionAct02DoorPos = POS_RUN Then
        If ElapseTime(TM_VISION) > 10 Then
            Call OnLog("[ERROR] VISION Act02 DOOR TEST TIME OVER !!!")
            lBkColor = vbRed
            RunVar.nVisionAct02DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "RUNNING" Then
            bSignal = RunVar.bVisionAct02DoorAck
            RunVar.bVisionAct02DoorAck = True
            
            If bSignal <> RunVar.bVisionAct02DoorAck Then
                lBkColor = vbWhite
                Call OnLog("VISION Act02 DOOR ACK...")
            End If
        End If
        
        If Trim$(lpVisionCom) = ("RM" & Format(nCameraNo, "000")) Then
            lBkColor = vbRed
            RunVar.nVisionAct02DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "EMOK" Then
            lBkColor = vbGreen
            RunVar.nVisionAct02DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "EMNG" Then
            lBkColor = vbRed
            RunVar.nVisionAct02DoorPos = POS_END
        End If
    End If
    
    If RunVar.nVisionAct02DoorPos = POS_END Then
        Call VisionDoorResult(nCameraNo, bOpen, lBkColor, SetupVar.nActBoardNo(1))
        
        RunVar.nVisionAct02DoorPos = 0
        RunVar.bVisionAct02DoorRunning = False
        
        If bOpen = True Then RunVar.nVisionAct02DoorReturn = &H1
        If bOpen = False Then RunVar.nVisionAct02DoorReturn = &H2
        
        RunVar.bVisionDoorManualOpenUse(nCameraNo) = False
        RunVar.bVisionDoorManualCloseUse(nCameraNo) = False
        
        RunVar.bVisionAct02DoorExit = True
        bRes = True
    End If
    
    VisionAct02DoorTest = bRes
End Function

Public Function OnVisionAct03Door(ByVal nPos As Integer) As Boolean
    Dim bRes As Boolean
    
    bRes = False
    
    If RunVar.bVisionAct03DoorMarking(nPos) Then
        RunVar.nVisionAct03OnDoorPos = POS_END
    End If
    
    If RunVar.nVisionAct03OnDoorPos = POS_INIT Then
        RunVar.bVisionAct03DoorExit = False
        RunVar.nVisionAct03DoorLoop = 0 ' 한 포지션에서 두개 이상이 겹질 경우를 위해
        
        If nPos = SetupVar.nOpenCameraPos(2) Then RunVar.nVisionAct03DoorLoop = RunVar.nVisionAct03DoorLoop + &H1
        If nPos = SetupVar.nCloseCameraPos(2) Then RunVar.nVisionAct03DoorLoop = RunVar.nVisionAct03DoorLoop + &H2
        
        RunVar.nVisionAct03OnDoorPos = POS_RUN_INIT
    End If
    
    If RunVar.nVisionAct03OnDoorPos = POS_RUN_INIT Then
        If RunVar.nVisionAct03DoorLoop = 0 Then
            RunVar.nVisionAct03OnDoorPos = POS_END
        Else
            'Call OnLog("START VISION DOOR COUNTER : 0x" & Hex(.nVisionAct03DoorLoop))
            
            RunVar.nVisionAct03DoorPos = POS_INIT
            RunVar.nVisionAct03OnDoorPos = POS_RUN
        End If
    End If
    
    If RunVar.nVisionAct03OnDoorPos = POS_RUN Then
        Select Case RunVar.nVisionAct03DoorLoop
            Case &H1:
                bRes = VisionAct03DoorTest(2, True)
            
            Case &H2:
                bRes = VisionAct03DoorTest(2, False)
            
            Case Else:
                RunVar.nVisionAct03OnDoorPos = POS_END
        
        End Select
        
        If bRes Then
            If RunVar.nVisionAct03DoorLoop = 0 Then
                RunVar.nVisionAct03OnDoorPos = POS_END
            Else
                RunVar.nVisionAct03DoorDone = (RunVar.nVisionAct03DoorDone Or RunVar.nVisionAct03DoorReturn) ' 테스트 한것은 테스트 하지 않기 위해 테스트 프래그를 설정한다.
                RunVar.nVisionAct03DoorLoop = RunVar.nVisionAct03DoorLoop - (RunVar.nVisionAct03DoorLoop And RunVar.nVisionAct03DoorDone)   ' 테스트 한것은 테스트 하지 않는다.
                
                RunVar.nVisionAct03OnDoorPos = POS_INIT
                'Call OnLog("REMAIN VISION DOOR COUNTER : 0x" & Hex(.nVisionAct03DoorLoop))
                
                If RunVar.bVisionAct03DoorExit And RunVar.nVisionAct03DoorLoop = 0 Then
                    RunVar.nVisionAct03OnDoorPos = POS_END
                Else
                    RunVar.bVisionAct03DoorExit = False
                    RunVar.nVisionAct03OnDoorPos = POS_RUN_INIT
                End If
            End If
        End If
        
        bRes = False    ' 안해주면 안되용
    End If
    
    If RunVar.nVisionAct03OnDoorPos = POS_END Then
        RunVar.bVisionAct03DoorMarking(nPos) = True ' Marking
        bRes = True
    End If
    
    OnVisionAct03Door = bRes
End Function

Public Function VisionAct03DoorTest(ByVal nCameraNo As Integer, ByVal bOpen As Boolean) As Boolean
    Dim lBkColor As Long
    Dim bRes As Boolean
    Dim bSignal As Boolean
    Dim lpCmd As String
    
    bRes = False
    
    If RunVar.nVisionAct03DoorPos = POS_INIT Then
        RunVar.bVisionAct03DoorRunning = True
        RunVar.nVisionAct03DoorReCount = 0
        RunVar.bVisionAct03DoorAck = False
        RunVar.nVisionAct03DoorPos = POS_RUN_INIT
    End If
    
    If RunVar.nVisionAct03DoorPos = POS_RUN_INIT Then
        If bOpen Then
            lpCmd = "OP" & Format(SetupVar.nOpenCameraNo(nCameraNo) + 1, "000")
        Else
            lpCmd = "CL" & Format(SetupVar.nCloseCameraNo(nCameraNo) + 1, "000")
        End If
        
        Call VisionDoorSend(lpCmd)
        Call OnLog("VISION ACT 03 DOOR SEND:" & lpCmd)
        
        RunVar.nVisionAct03DoorPos = POS_RUN
        Call SetTime(TM_VISION)
    End If
    
    If RunVar.nVisionAct03DoorPos = POS_RUN Then
        If ElapseTime(TM_VISION) > 10 Then
            Call OnLog("[ERROR] VISION Act03 DOOR TEST TIME OVER !!!")
            lBkColor = vbRed
            RunVar.nVisionAct03DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "RUNNING" Then
            bSignal = RunVar.bVisionAct03DoorAck
            RunVar.bVisionAct03DoorAck = True
            
            If bSignal <> RunVar.bVisionAct03DoorAck Then
                lBkColor = vbWhite
                Call OnLog("VISION Act03 DOOR ACK...")
            End If
        End If
        
        If Trim$(lpVisionCom) = ("RM" & Format(nCameraNo, "000")) Then
            lBkColor = vbRed
            RunVar.nVisionAct03DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "EMOK" Then
            lBkColor = vbGreen
            RunVar.nVisionAct03DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "EMNG" Then
            lBkColor = vbRed
            RunVar.nVisionAct03DoorPos = POS_END
        End If
    End If
    
    If RunVar.nVisionAct03DoorPos = POS_END Then
        Call VisionDoorResult(nCameraNo, bOpen, lBkColor, SetupVar.nActBoardNo(2))
        
        RunVar.nVisionAct03DoorPos = 0
        RunVar.bVisionAct03DoorRunning = False
        
        If bOpen = True Then RunVar.nVisionAct03DoorReturn = &H1
        If bOpen = False Then RunVar.nVisionAct03DoorReturn = &H2
        
        RunVar.bVisionDoorManualOpenUse(nCameraNo) = False
        RunVar.bVisionDoorManualCloseUse(nCameraNo) = False
        
        RunVar.bVisionAct03DoorExit = True
        bRes = True
    End If
    
    VisionAct03DoorTest = bRes
End Function

Public Function OnVisionAct04Door(ByVal nPos As Integer) As Boolean
    Dim bRes As Boolean
    
    bRes = False
    
    If RunVar.bVisionAct04DoorMarking(nPos) Then
        RunVar.nVisionAct04OnDoorPos = POS_END
    End If
    
    If RunVar.nVisionAct04OnDoorPos = POS_INIT Then
        RunVar.bVisionAct04DoorExit = False
        RunVar.nVisionAct04DoorLoop = 0 ' 한 포지션에서 두개 이상이 겹질 경우를 위해
        
        If nPos = SetupVar.nOpenCameraPos(3) Then RunVar.nVisionAct04DoorLoop = RunVar.nVisionAct04DoorLoop + &H1
        If nPos = SetupVar.nCloseCameraPos(3) Then RunVar.nVisionAct04DoorLoop = RunVar.nVisionAct04DoorLoop + &H2
        
        RunVar.nVisionAct04OnDoorPos = POS_RUN_INIT
    End If
    
    If RunVar.nVisionAct04OnDoorPos = POS_RUN_INIT Then
        If RunVar.nVisionAct04DoorLoop = 0 Then
            RunVar.nVisionAct04OnDoorPos = POS_END
        Else
            'Call OnLog("START VISION DOOR COUNTER : 0x" & Hex(.nVisionAct04DoorLoop))
            
            RunVar.nVisionAct04DoorPos = POS_INIT
            RunVar.nVisionAct04OnDoorPos = POS_RUN
        End If
    End If
    
    If RunVar.nVisionAct04OnDoorPos = POS_RUN Then
        Select Case RunVar.nVisionAct04DoorLoop
            Case &H1:
                bRes = VisionAct04DoorTest(3, True)
            
            Case &H2:
                bRes = VisionAct04DoorTest(3, False)
            
            Case Else:
                RunVar.nVisionAct04OnDoorPos = POS_END
        
        End Select
        
        If bRes Then
            If RunVar.nVisionAct04DoorLoop = 0 Then
                RunVar.nVisionAct04OnDoorPos = POS_END
            Else
                RunVar.nVisionAct04DoorDone = (RunVar.nVisionAct04DoorDone Or RunVar.nVisionAct04DoorReturn) ' 테스트 한것은 테스트 하지 않기 위해 테스트 프래그를 설정한다.
                RunVar.nVisionAct04DoorLoop = RunVar.nVisionAct04DoorLoop - (RunVar.nVisionAct04DoorLoop And RunVar.nVisionAct04DoorDone)   ' 테스트 한것은 테스트 하지 않는다.
                
                RunVar.nVisionAct04OnDoorPos = POS_INIT
                'Call OnLog("REMAIN VISION DOOR COUNTER : 0x" & Hex(.nVisionAct04DoorLoop))
                
                If RunVar.bVisionAct04DoorExit And RunVar.nVisionAct04DoorLoop = 0 Then
                    RunVar.nVisionAct04OnDoorPos = POS_END
                Else
                    RunVar.bVisionAct04DoorExit = False
                    RunVar.nVisionAct04OnDoorPos = POS_RUN_INIT
                End If
            End If
        End If
        
        bRes = False    ' 안해주면 안되용
    End If
    
    If RunVar.nVisionAct04OnDoorPos = POS_END Then
        RunVar.bVisionAct04DoorMarking(nPos) = True ' Marking
        bRes = True
    End If
    
    OnVisionAct04Door = bRes
End Function

Public Function VisionAct04DoorTest(ByVal nCameraNo As Integer, ByVal bOpen As Boolean) As Boolean
    Dim lBkColor As Long
    Dim bRes As Boolean
    Dim bSignal As Boolean
    Dim lpCmd As String
    
    bRes = False
    
    If RunVar.nVisionAct04DoorPos = POS_INIT Then
        RunVar.bVisionAct04DoorRunning = True
        RunVar.nVisionAct04DoorReCount = 0
        RunVar.bVisionAct04DoorAck = False
        RunVar.nVisionAct04DoorPos = POS_RUN_INIT
    End If
    
    If RunVar.nVisionAct04DoorPos = POS_RUN_INIT Then
        If bOpen Then
            lpCmd = "OP" & Format(SetupVar.nOpenCameraNo(nCameraNo) + 1, "000")
        Else
            lpCmd = "CL" & Format(SetupVar.nCloseCameraNo(nCameraNo) + 1, "000")
        End If
        
        Call VisionDoorSend(lpCmd)
        Call OnLog("VISION ACT 04 DOOR SEND:" & lpCmd)
        
        RunVar.nVisionAct04DoorPos = POS_RUN
        Call SetTime(TM_VISION)
    End If
    
    If RunVar.nVisionAct04DoorPos = POS_RUN Then
        If ElapseTime(TM_VISION) > 10 Then
            Call OnLog("[ERROR] VISION Act04 DOOR TEST TIME OVER !!!")
            lBkColor = vbRed
            RunVar.nVisionAct04DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "RUNNING" Then
            bSignal = RunVar.bVisionAct04DoorAck
            RunVar.bVisionAct04DoorAck = True
            
            If bSignal <> RunVar.bVisionAct04DoorAck Then
                lBkColor = vbWhite
                Call OnLog("VISION Act04 DOOR ACK...")
            End If
        End If
        
        If Trim$(lpVisionCom) = ("RM" & Format(nCameraNo, "000")) Then
            lBkColor = vbRed
            RunVar.nVisionAct04DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "EMOK" Then
            lBkColor = vbGreen
            RunVar.nVisionAct04DoorPos = POS_END
        End If
        
        If Trim$(lpVisionCom) = "EMNG" Then
            lBkColor = vbRed
            RunVar.nVisionAct04DoorPos = POS_END
        End If
    End If
    
    If RunVar.nVisionAct04DoorPos = POS_END Then
        Call VisionDoorResult(nCameraNo, bOpen, lBkColor, SetupVar.nActBoardNo(3))
        
        RunVar.nVisionAct04DoorPos = 0
        RunVar.bVisionAct04DoorRunning = False
        
        If bOpen = True Then RunVar.nVisionAct04DoorReturn = &H1
        If bOpen = False Then RunVar.nVisionAct04DoorReturn = &H2
        
        RunVar.bVisionDoorManualOpenUse(nCameraNo) = False
        RunVar.bVisionDoorManualCloseUse(nCameraNo) = False
        
        RunVar.bVisionAct04DoorExit = True
        bRes = True
    End If
    
    VisionAct04DoorTest = bRes
End Function

Public Sub VisionDoorResult(ByVal nCameraNo As Integer, ByVal bOpen As Boolean, ByVal lResult As Long, ByVal nActNo As Integer)
    If lResult = vbGreen Then
        If bOpen Then
            frmRun.pnlVisionDoorOpen(nCameraNo).Caption = "OK"
            frmRun.pnlVisionDoorOpen(nCameraNo).BackColor = lResult
            frmRun.pnlVisionDoorOpenNum(nCameraNo).BackColor = lResult
        Else
            frmRun.pnlVisionDoorClose(nCameraNo).Caption = "OK"
            frmRun.pnlVisionDoorClose(nCameraNo).BackColor = lResult
            frmRun.pnlVisionDoorCloseNum(nCameraNo).BackColor = lResult
        End If
    Else
        If bOpen Then
            frmRun.pnlVisionDoorOpen(nCameraNo).Caption = "NG"
            frmRun.pnlVisionDoorOpen(nCameraNo).BackColor = lResult
            frmRun.pnlVisionDoorOpenNum(nCameraNo).BackColor = lResult
        Else
            frmRun.pnlVisionDoorClose(nCameraNo).Caption = "NG"
            frmRun.pnlVisionDoorClose(nCameraNo).BackColor = lResult
            frmRun.pnlVisionDoorCloseNum(nCameraNo).BackColor = lResult
        End If
        
        Select Case nActNo
            Case 1: RunVar.bReAct01Use = True
            Case 2: RunVar.bReAct02Use = True
            Case 3: RunVar.bReAct03Use = True
            Case 4: RunVar.bReAct04Use = True
        End Select
        
        RunVar.bReVisionUse = True
        RunVar.bFinal = False
    End If
End Sub

