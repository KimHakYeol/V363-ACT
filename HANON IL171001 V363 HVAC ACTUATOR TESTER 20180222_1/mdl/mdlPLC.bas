Attribute VB_Name = "mdlPLC"
Option Explicit


Public Const PLC_OK             As Long = &H1
Public Const PLC_NG             As Long = &H2
Public Const PLC_PASS           As Long = &H4

Public Function PLC_Open() As Boolean
    Dim lRes As Long
    Dim bRes As Boolean
    
    If PLCUSE = False Then Exit Function
    
    bRes = True
    
    lRes = frmMain.ActPlc.Close
    
    Call Delay(500)
    
    lRes = frmMain.ActPlc.Open
    
    If lRes <> 0 Then bRes = False
    
    PLC_Open = bRes
End Function

Public Function PLC_Close()
    Dim lRes As Long
    
    If PLCUSE = False Then Exit Function
    
    lRes = frmMain.ActPlc.Close
End Function

Public Function SetPlc(ByVal nValue As Long)
    If PLCUSE = False Or nValue = &H0 Then
        Exit Function
    End If
    
    If nValue < &H10000 Then
        PlcVar.lTotalResult(0) = PlcVar.lTotalResult(0) Or nValue
    Else
        PlcVar.lTotalResult(1) = PlcVar.lTotalResult(1) Or nValue \ 2 ^ (4 * 4) ' 4 bit right
    End If
End Function

Public Function SetClearPlc()
    If PLCUSE = False Then Exit Function
    
    Erase PlcVar.lTotalResult
End Function

Public Function BitChk(ByVal lData As Long, ByVal lChkBit As Long) As Boolean
    If lData And (2 ^ lChkBit) Then
        BitChk = True
    Else
        BitChk = False
    End If
End Function

Public Function PLC_Proc() As Boolean
    Const MAXARRAY As Integer = 11
    Const PLCRECOUNT As Integer = 5
    
    Dim i As Integer
    Dim lRes(MAXARRAY) As Long
    Dim nRes As Integer
    Dim nCount As Integer
    Dim lWrite As Long
    Dim lData As Long
    Dim lpErrorStr As String
    Dim lpErrorData As String
    
    Erase lRes
    Erase PlcData
    
    If bPlcStartSig Then Exit Function
    
    ' PLC로 받는 데이터
    lRes(MAXARRAY) = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrReady, 1, PlcVar.lGetPlcData(0))
    
    If PlcVar.lGetPlcData(0) = 1 Then
        lWrite = &HFFFF
    Else
        lWrite = &H0
    End If
    
    ' 테스트 시작
    If lWrite = &HFFFF Then
        
        lRes(0) = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrStart, 1, lData)
        
        If lData = 1 Then
            
PLCRELOAD01:
            
            ' 파렛트
            lRes(1) = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrPallet, 1, PlcData(0))
            
            If lRes(1) = 0 Then
                frmRun.pnlCarType(3).Caption = Trim$(PlcData(0))
                DataVar.lpPallet = frmRun.pnlCarType(3).Caption
            Else
                nCount = nCount + 1
                
                If nCount > PLCRECOUNT Then
                    lpErrorStr = "Pallet"
                    lpErrorData = PlcData(0)
                    
                    GoTo PLCLOADERROR
                Else
                    GoTo PLCRELOAD01
                End If
            End If
            
PLCRELOAD02:
            
            ' 차종 번호
            lRes(2) = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrCarType, 1, PlcData(1))
            
            If lRes(2) = 0 Then
                frmRun.pnlCarType(0).Caption = Trim$(PlcData(1))
                DataVar.lpModelNo = frmRun.pnlCarType(0).Caption
            Else
                nCount = nCount + 1
                
                If nCount > PLCRECOUNT Then
                    lpErrorStr = "CarType"
                    lpErrorData = PlcData(1)
                    
                    GoTo PLCLOADERROR
                Else
                    GoTo PLCRELOAD02
                End If
            End If
            
PLCRELOAD03:
            
            ' 서열 번호
            lRes(3) = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrCarRank, 1, PlcData(2))
            
            If lRes(3) = 0 Then
                frmRun.pnlCarType(1).Caption = Trim$(PlcData(2))
                DataVar.lpModelRank = frmRun.pnlCarType(1).Caption
            Else
                nCount = nCount + 1
                
                If nCount > PLCRECOUNT Then
                    lpErrorStr = "CarRank"
                    lpErrorData = PlcData(2)
                    
                    GoTo PLCLOADERROR
                Else
                    GoTo PLCRELOAD03
                End If
            End If
            
PLCRELOAD04:
            
            ' 그룹
            lRes(4) = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrCarGroup, 1, PlcData(3))
            
            If lRes(4) = 0 Then
                frmRun.pnlCarType(2).Caption = Trim$(PlcData(3))
                DataVar.lpModelGroup = frmRun.pnlCarType(2).Caption
            Else
                nCount = nCount + 1
                
                If nCount > PLCRECOUNT Then
                    lpErrorStr = "Group"
                    lpErrorData = PlcData(3)
                    
                    GoTo PLCLOADERROR
                Else
                    GoTo PLCRELOAD04
                End If
            End If
            
PLCRELOAD05:
            
            ' 시리얼
            lRes(5) = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrSerial(0), 1, PlcData(10))
            lRes(6) = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrSerial(1), 1, PlcData(11))
            lRes(7) = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrSerial(2), 1, PlcData(12))
            lRes(8) = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrSerial(3), 1, PlcData(13))
            lRes(9) = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrSerial(4), 1, PlcData(14))
            
            If lRes(5) = 0 And lRes(6) = 0 And lRes(7) = 0 And lRes(8) = 0 And lRes(9) = 0 Then
                frmRun.pnlSerial.Caption = Format(PlcData(10), "0000") & Format(PlcData(11), "00") & Format(PlcData(12), "00") & Format(PlcData(14), "000") & Format(PlcData(13), "0000")
                DataVar.lpSerialNo = frmRun.pnlSerial.Caption
            Else
                nCount = nCount + 1
                
                If nCount > PLCRECOUNT Then
                    lpErrorStr = "Serial"
                    lpErrorData = Format(PlcData(10), "0000") & Format(PlcData(11), "00") & Format(PlcData(12), "00") & Format(PlcData(14), "000") & Format(PlcData(13), "0000")
                    
                    GoTo PLCLOADERROR
                Else
                    GoTo PLCRELOAD05
                End If
            End If
            
            ' FINISH
            nRes = 2
            nCount = SUBSAVECOUNT
            
            PlcData(1) = PlcData(1) - 1
            PlcData(nRes) = PlcData(nRes) - 1
            nNowModelNo = (PlcData(1) * (nCount + 1)) + PlcData(nRes)
            
            If nNowModelNo < 0 Then
                lpErrorStr = "ModelNo"
                lpErrorData = nNowModelNo
                
                GoTo PLCLOADERROR
            End If
            
            lpNowModel = Format(nNowModelNo, "000") & "_" & SelectCar(PlcData(1)).ModelName & "_" & SelectCar(PlcData(1)).ModelNameSub(PlcData(nRes))
            
            bPlcStartSig = True
        End If
    End If
    
    Exit Function

PLCLOADERROR:

    If DEBUGMODE Then
        Debug.Print "PLC LOAD ERROR"
        
        For i = 0 To MAXARRAY
            Debug.Print "PLC LOAD RESULT : " & lRes(i)
        Next
        
        For i = 0 To UBound(PlcData)
            If PlcData(i) > 0 Then
                Debug.Print "PLC ARRAY " & Format(i, "00") & vbTab & vbTab & "DATA : " & PlcData(i)
            End If
        Next
    Else
        Call OnLog("PLC LOAD ERROR... [" & lpErrorStr & " : " & lpErrorData & "]")
    End If
    
    PLC_Proc = False
    bPlcStartSig = False
End Function

Public Sub PLC_TestRun()
    Dim i As Integer
    Dim nCount As Integer
    Dim lRes As Long
    Dim bRes(9) As Boolean
    
    If PLCUSE = False Then Exit Sub
    If nStartSignal <> 1 Then Exit Sub
    
    Erase bRes
    
    nCount = 0
    
    Call SetClearPlc
    
PLCRELOAD1:
    
    lRes = frmMain.ActPlc.WriteDeviceBlock(PlcVar.lpAddrStart, 1, 1)
    
    If lRes = 0 Then
        bRes(0) = True
        nCount = 0
    Else
        nCount = nCount + 1
        
        If nCount > 5 Then
            GoTo PLCRUNERROR
        Else
            GoTo PLCRELOAD1
        End If
    End If

PLCRELOAD2:
    
    lRes = frmMain.ActPlc.WriteDeviceBlock(PlcVar.lpAddrReady, 1, 1)
    
    If lRes = 0 Then
        bRes(1) = True
        nCount = 0
    Else
        nCount = nCount + 1
        
        If nCount > 5 Then
            GoTo PLCRUNERROR
        Else
            GoTo PLCRELOAD2
        End If
    End If

PLCRELOAD3:
    
    lRes = frmMain.ActPlc.WriteDeviceBlock(PlcVar.lpAddrRunning, 1, 1)
    
    If lRes = 0 Then
        bRes(2) = True
        nCount = 0
    Else
        nCount = nCount + 1
        
        If nCount > 5 Then
            GoTo PLCRUNERROR
        Else
            GoTo PLCRELOAD3
        End If
    End If
    
    For i = 0 To 2
        If bRes(i) = False Then
            GoTo PLCRUNERROR
            
            Exit For
        End If
    Next
    
    frmRun.pnlPlcStatus(2).BackColor = CO_GRASS ' plc test run
    
    frmRun.pnlPlcStatus(5).BackColor = CO_BNONE
    frmRun.pnlPlcStatus(5).Caption = " "
    
    Exit Sub
    
PLCRUNERROR:
    
    Call OnLog("[PLC] TEST RUN ERROR...")
End Sub

Public Sub PLC_TestEnd()
    Dim i As Integer
    Dim nCount As Integer
    Dim lRes As Long
    Dim bRes(9) As Boolean
    
    If PLCUSE = False Then Exit Sub
    If nStartSignal <> 1 Then Exit Sub
    
    Erase bRes
    
    nCount = 0

PLCRELOAD1:
    
    lRes = frmMain.ActPlc.WriteDeviceBlock(PlcVar.lpAddrResult, 1, IIf(RunVar.bFinal, 1, 2))
    
    If lRes = 0 Then
        bRes(0) = True
        nCount = 0
    Else
        nCount = nCount + 1
        
        If nCount > 5 Then
            GoTo PLCRUNERROR
        Else
            GoTo PLCRELOAD1
        End If
    End If

PLCRELOAD2:
    
    lRes = frmMain.ActPlc.WriteDeviceBlock(PlcVar.lpAddrReady, 1, 0)
    
    If lRes = 0 Then
        bRes(1) = True
        nCount = 0
    Else
        nCount = nCount + 1
        
        If nCount > 5 Then
            GoTo PLCRUNERROR
        Else
            GoTo PLCRELOAD2
        End If
    End If

PLCRELOAD3:
    
    lRes = frmMain.ActPlc.WriteDeviceBlock(PlcVar.lpAddrNGList, 2, PlcVar.lTotalResult(0))
    
    If lRes = 0 Then
        bRes(2) = True
        nCount = 0
    Else
        nCount = nCount + 1
        
        If nCount > 5 Then
            GoTo PLCRUNERROR
        Else
            GoTo PLCRELOAD3
        End If
    End If

PLCRELOAD4:
    
    lRes = frmMain.ActPlc.WriteDeviceBlock(PlcVar.lpAddrRunning, 1, 0)
    
    If lRes = 0 Then
        bRes(3) = True
        nCount = 0
    Else
        nCount = nCount + 1
        
        If nCount > 5 Then
            GoTo PLCRUNERROR
        Else
            GoTo PLCRELOAD4
        End If
    End If
    
    For i = 0 To 2
        If bRes(i) = False Then
            GoTo PLCRUNERROR
            
            Exit For
        End If
    Next
    
    frmRun.pnlPlcStatus(2).BackColor = CO_BNONE ' plc test run
    
    frmRun.pnlPlcStatus(5).BackColor = IIf(RunVar.bFinal, CO_GRASS, CO_ORANGE)
    frmRun.pnlPlcStatus(5).Caption = IIf(RunVar.bFinal, "OK  ", "NG  ")
    
    Call ClearPLC
    
    Exit Sub

PLCRUNERROR:
    
    Call OnLog("[PLC] TEST END ERROR...")
End Sub

Public Sub ClearPLC()
    Dim lRes As Long
    Dim bRes(5) As Boolean
    Dim bTotalRes As Boolean
    Dim i As Integer
    
    Erase bRes
    
    bTotalRes = True
    
    lRes = frmMain.ActPlc.WriteDeviceBlock(PlcVar.lpAddrStart, 1, 0)
    If lRes <> 0 Then bRes(0) = False
    
    For i = 0 To 5
        If bRes(i) = False Then
            bTotalRes = False
            Exit For
        End If
    Next
    
    If bTotalRes Then
        bPlcStartSig = False
    End If
End Sub

Public Sub PlcStatus()
    Dim nCount As Integer
    Dim lRes As Long
    Dim lData(1) As Long
    
    If PLCUSE And SysVar.bPlcCommUse Then
        
PLCREADYRELOAD:
        
        lRes = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrReady, 1, lData(0))
        
        If lRes = 0 Then
            frmRun.pnlPlcStatus(1).BackColor = IIf(lData(0) = 0, CO_BNONE, CO_ORANGE) ' plc ready
        Else
            nCount = nCount + 1
            
            If nCount > 10 Then
                frmRun.pnlPlcStatus(1).BackColor = CO_ORANGE
                
                GoTo PLCSTATUS_ERROR
            Else
                GoTo PLCREADYRELOAD
            End If
        End If
        
        lToggle = IIf(lToggle = 0, 1, 0)
        
PLCSTATUSRELOAD:
        
        lRes = frmMain.ActPlc.WriteDeviceBlock(PlcVar.lpAddrStatus, 1, lToggle)
        
        If lRes = 0 Then
            frmRun.pnlPlcStatus(0).BackColor = IIf(lToggle = 0, CO_BNONE, CO_GRASS) ' plc status
        Else
            nCount = nCount + 1
            
            If nCount > 10 Then
                frmRun.pnlPlcStatus(0).BackColor = CO_ORANGE
                
                GoTo PLCSTATUS_ERROR
            Else
                GoTo PLCSTATUSRELOAD
            End If
        End If
    End If
    
    If lData(0) = 0 And RunVar.bRun And bPlcStartSig Then
        bPlcStopSig = True
    Else
        bPlcStopSig = False
    End If
    
    Exit Sub
    
PLCSTATUS_ERROR:
    
    Call PLC_Close
    Call Delay(1000)
    Call PLC_Open
    Call Delay(1000)
End Sub

Public Function PlcDataTrackingSave()
    ' 사용안함
End Function

Public Function ExtractNumber(ByVal InputString As String)
    Dim i           As Integer
    Dim Num         As String
    
    For i = Len(InputString) To 1 Step -1
        If IsNumeric(Mid(InputString, i, 1)) Then
            Num = Mid(InputString, i, 1) & Num
        End If
    Next
    
    ExtractNumber = Num
End Function

Public Function AsciiChange(ByVal TempString As Long) As String
    Dim txtModel As String
    
    txtModel = ""
    
    If TempString <> 0 Then
        txtModel = txtModel + Chr(TempString And &HFF)
        txtModel = txtModel + Chr((TempString And &HFF00) / 256)
    End If
    
    AsciiChange = txtModel
End Function

Public Sub PlcLeakGet()
    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer
    Dim lRes As Long
    Dim bRes(5) As Boolean
    Dim bTotalRes As Boolean
    Dim lData1(19) As Long
    Dim txtData1(9) As String
    Dim txtData2(9) As String
    
    For i = 0 To 5
        bRes(i) = True
    Next
    
    bTotalRes = True
    nCount = 0
    
    If PLCUSE = False Then Exit Sub
    If nStartSignal <> 1 Then Exit Sub
    
RELOAD:
    
    lRes = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrLeak(0), 12, lData1(0))
    
    If lRes = 0 Then
        txtData1(0) = ""
        
        For j = 1 To 4: txtData1(0) = txtData1(0) & AsciiChange(lData1(j)): Next
        For j = 7 To 10: txtData2(0) = txtData2(0) & AsciiChange(lData1(j)): Next
        
        DataVar.lpLeak(0) = txtData1(0)
        If DataVar.lpLeak(0) = "" Then DataVar.lpLeak(0) = "NG"
        If lData1(0) = 2 Then DataVar.lpLeak(0) = "#" & DataVar.lpLeak(0)
        
        DataVar.lpLeak(1) = txtData2(0)
        If DataVar.lpLeak(1) = "" Then DataVar.lpLeak(1) = "NG"
        If lData1(6) = 2 Then DataVar.lpLeak(1) = "#" & DataVar.lpLeak(1)
    Else
        nCount = nCount + 1
        
        If nCount > 5 Then
            Call OnLog("[PLC] LEAK LOADING ERROR...")
            
            Exit Sub
        Else
            GoTo RELOAD
        End If
    End If
End Sub

Public Sub PlcBarcodeGet()
    Dim i As Integer
    Dim nCount As Integer
    Dim lRes As Long
    Dim lData(19) As Long
    Dim txtData(19) As String
    
    nCount = 0
    
    If PLCUSE = False Then Exit Sub
    If nStartSignal <> 1 Then Exit Sub
    
    Erase lData
    Erase txtData
    
RELOAD1:
    
    lRes = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrBarcode(0), 10, lData(0))
    
    If lRes = 0 Then
        txtData(0) = ""
        
        For i = 0 To 9
            txtData(0) = txtData(0) & AsciiChange(lData(i))
        Next
        
        DataVar.lpBarcode(0) = txtData(0)
    Else
        nCount = nCount + 1
        
        If nCount > 5 Then
            Call OnLog("[PLC] LEAK01 LOADING ERROR...")
            
            Exit Sub
        Else
            GoTo RELOAD1
        End If
    End If

    Erase lData
    Erase txtData
    
RELOAD2:
    
    lRes = frmMain.ActPlc.ReadDeviceBlock(PlcVar.lpAddrBarcode(1), 12, lData(0))
    
    If lRes = 0 Then
        txtData(1) = ""
        
        For i = 0 To 11
            txtData(1) = txtData(1) & AsciiChange(lData(i))
        Next
        
        DataVar.lpBarcode(1) = txtData(1)
    Else
        nCount = nCount + 1
        
        If nCount > 5 Then
            Call OnLog("[PLC] LEAK02 LOADING ERROR...")
            
            Exit Sub
        Else
            GoTo RELOAD2
        End If
    End If
End Sub

