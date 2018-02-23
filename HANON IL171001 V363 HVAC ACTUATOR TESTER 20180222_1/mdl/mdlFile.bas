Attribute VB_Name = "mdlFile"
Option Explicit


Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub INIModelListRead(ByRef cboRef As ComboBox)
    Dim i                   As Long
    Dim ret                 As Long
    Dim s                   As String * 8192
    Dim tmp()               As String
    
    ret = GetPrivateProfileSectionNames(s, 8192, lpPath & INISETUPFILE)
    
    If ret > 1 Then
        tmp = Split(Left$(s, ret - 1), Chr(0))
        
        For i = 0 To UBound(tmp)
            cboRef.AddItem Trim$(tmp(i))
        Next
    End If
End Sub

Public Sub INIModelNameRead()
    Dim i                   As Long
    Dim ret                 As Long
    Dim s                   As String * 8192
    Dim tmp()               As String
    
    ret = GetPrivateProfileSectionNames(s, 8192, lpPath & INIMODELFILE)
    
    If ret > 1 Then
        tmp = Split(Left$(s, ret - 1), Chr(0))
        
        For i = 0 To UBound(tmp)
            SelectCar(i).ModelName = Trim$(tmp(i))
        Next
    End If
End Sub

Public Sub INIModelDelete(m_strSection As String)
    Dim lngRet          As Long
    Dim strKeyName      As String
    Dim strlpString     As String
    
    lngRet = WritePrivateProfileString(m_strSection, strKeyName, strlpString, lpPath & INISETUPFILE)
End Sub

Public Function INIWrite(Section As String, KeyValue As String, Data As String, INIFile As String) As String
    Dim lngRet As Long
    
    lngRet = WritePrivateProfileString(Section, KeyValue, Data, INIFile)
End Function

Public Function INIRead(Section As String, KeyValue As String, INIFile As String) As String
    Dim lngRet      As Long
    Dim strValue    As String * 256
    
    lngRet = GetPrivateProfileString(Section, KeyValue, "", strValue, 256, INIFile)
    INIRead = Left$(strValue, InStr(strValue, Chr(0)) - 1)
End Function

Public Sub SetupDisp2Mem()
    Dim i As Integer
    
    ' General
    SetupVar.dTestVolt = Val(frmSetup.txtTestVolt.Text)
    SetupVar.lpFileName = Trim$(frmSetup.txtFileName.Text)
    SetupVar.bDataSave = IIf(frmSetup.chkSaveUse.Value = 1, True, False)
    SetupVar.bScannerUse = IIf(frmSetup.chkScannerUse.Value = 1, True, False)
    SetupVar.nScannerValue = Trim$(frmSetup.txtScannerValue)
    SetupVar.bNvhUse = IIf(frmSetup.chkNvh.Value = 1, True, False)
    
    For i = 0 To frmSetup.optModelType.UBound
        If frmSetup.optModelType(i).Value = True Then
            SetupVar.nModelType = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.txtActName.UBound
        SetupVar.lpActName(i) = Trim$(frmSetup.txtActName(i).Text)
        SetupVar.nActBoardNo(i) = Trim$(frmSetup.txtActBoardNo(i).Text)
    Next
    
    ' Blower
    SetupVar.bBlowerUse = IIf(frmSetup.chkBlower.Value = 1, True, False)
    
    For i = 0 To frmSetup.optBlowerType.UBound
        If frmSetup.optBlowerType(i).Value = True Then
            SetupVar.nBlowerType = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.optBlowerDirection.UBound
        If frmSetup.optBlowerDirection(i).Value = True Then
            SetupVar.nBlowerDirection = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.txtBlowerName.UBound
        SetupVar.lpBlowerName(i) = Trim$(frmSetup.txtBlowerName(i).Text)
        SetupVar.dBlowerCurrLo(i) = Val(frmSetup.txtBlowerCurrLo(i).Text)
        SetupVar.dBlowerCurrHi(i) = Val(frmSetup.txtBlowerCurrHi(i).Text)
        SetupVar.dBlowerTime(i) = Val(frmSetup.txtBlowerTime(i).Text)
        SetupVar.nLinSpeed(i) = Val(frmSetup.txtLinSpeed(i).Text)
    Next
    
    ' RPM
    SetupVar.lpRpmName = Trim$(frmSetup.txtRpmName.Text)
    SetupVar.dRpmCurrLo = Val(frmSetup.txtRpmCurrLo.Text)
    SetupVar.dRpmCurrHi = Val(frmSetup.txtRpmCurrHi.Text)
    
    ' Vibration
    SetupVar.bVibUse = IIf(frmSetup.chkVib.Value = 1, True, False)
    SetupVar.lpVibName = Trim$(frmSetup.txtVibName.Text)
    SetupVar.dVibCurrLo = Val(frmSetup.txtVibCurrLo.Text)
    SetupVar.dVibCurrHi = Val(frmSetup.txtVibCurrHi.Text)
    
    For i = 0 To frmSetup.optVibResultType.UBound
        If frmSetup.optVibResultType(i).Value = True Then
            SetupVar.nVibResultType = i                         ' Peak / RMS
            Exit For
        End If
    Next
    
    SetupVar.dVibStart = Val(frmSetup.txtVibStart.Text)
    SetupVar.dVibEnd = Val(frmSetup.txtVibEnd.Text)
    
    For i = 0 To frmSetup.optVibMethod.UBound
        If frmSetup.optVibMethod(i).Value = True Then
            SetupVar.nVibMethod = i                              ' With Hi / After Hi
            Exit For
        End If
    Next
    
    SetupVar.dVibVolt = Val(frmSetup.txtVibVolt.Text)
    SetupVar.dVibTime = Val(frmSetup.txtVibTime.Text)
    
    ' Act 01
    SetupVar.bAct01Use = IIf(frmSetup.chkAct01.Value = 1, True, False)
    SetupVar.nAct01TestType = IIf(frmSetup.optAct01TestType(0).Value, 0, 1)
    
    For i = 0 To frmSetup.optAct01Direction.UBound
        If frmSetup.optAct01Direction(i).Value = True Then
            SetupVar.nAct01Direction = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.txtAct01Name.UBound
        SetupVar.lpAct01Name(i) = Trim$(frmSetup.txtAct01Name(i).Text)
        SetupVar.dAct01SetVolt(i) = Val(frmSetup.txtAct01SetVolt(i).Text)
        SetupVar.dAct01CurrLo(i) = Val(frmSetup.txtAct01CurrLo(i).Text)
        SetupVar.dAct01CurrHi(i) = Val(frmSetup.txtAct01CurrHi(i).Text)
        SetupVar.dAct01VoltLo(i) = Val(frmSetup.txtAct01VoltLo(i).Text)
        SetupVar.dAct01VoltHi(i) = Val(frmSetup.txtAct01VoltHi(i).Text)
        SetupVar.dAct01TimeLo(i) = Val(frmSetup.txtAct01TimeLo(i).Text)
        SetupVar.dAct01TimeHi(i) = Val(frmSetup.txtAct01TimeHi(i).Text)
    Next
    
    SetupVar.dAct01StallDeltaVoltLo = Val(frmSetup.txtAct01StallDeltaMinVolt.Text)
    SetupVar.dAct01StallDeltaVoltHi = Val(frmSetup.txtAct01StallDeltaMaxVolt.Text)
    
    ' Act 02
    SetupVar.bAct02Use = IIf(frmSetup.chkAct02.Value = 1, True, False)
    SetupVar.nAct02TestType = IIf(frmSetup.optAct02TestType(0).Value, 0, 1)
    
    For i = 0 To frmSetup.optAct02Direction.UBound
        If frmSetup.optAct02Direction(i).Value = True Then
            SetupVar.nAct02Direction = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.txtAct02Name.UBound
        SetupVar.lpAct02Name(i) = Trim$(frmSetup.txtAct02Name(i).Text)
        SetupVar.dAct02SetVolt(i) = Val(frmSetup.txtAct02SetVolt(i).Text)
        SetupVar.dAct02CurrLo(i) = Val(frmSetup.txtAct02CurrLo(i).Text)
        SetupVar.dAct02CurrHi(i) = Val(frmSetup.txtAct02CurrHi(i).Text)
        SetupVar.dAct02VoltLo(i) = Val(frmSetup.txtAct02VoltLo(i).Text)
        SetupVar.dAct02VoltHi(i) = Val(frmSetup.txtAct02VoltHi(i).Text)
        SetupVar.dAct02TimeLo(i) = Val(frmSetup.txtAct02TimeLo(i).Text)
        SetupVar.dAct02TimeHi(i) = Val(frmSetup.txtAct02TimeHi(i).Text)
    Next
    
    SetupVar.dAct02StallDeltaVoltLo = Val(frmSetup.txtAct02StallDeltaMinVolt.Text)
    SetupVar.dAct02StallDeltaVoltHi = Val(frmSetup.txtAct02StallDeltaMaxVolt.Text)
    
    ' Act 03
    SetupVar.bAct03Use = IIf(frmSetup.chkAct03.Value = 1, True, False)
    SetupVar.nAct03TestType = IIf(frmSetup.optAct03TestType(0).Value, 0, 1)
    
    For i = 0 To frmSetup.optAct03Direction.UBound
        If frmSetup.optAct03Direction(i).Value = True Then
            SetupVar.nAct03Direction = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.txtAct03Name.UBound
        SetupVar.lpAct03Name(i) = Trim$(frmSetup.txtAct03Name(i).Text)
        SetupVar.dAct03SetVolt(i) = Val(frmSetup.txtAct03SetVolt(i).Text)
        SetupVar.dAct03CurrLo(i) = Val(frmSetup.txtAct03CurrLo(i).Text)
        SetupVar.dAct03CurrHi(i) = Val(frmSetup.txtAct03CurrHi(i).Text)
        SetupVar.dAct03VoltLo(i) = Val(frmSetup.txtAct03VoltLo(i).Text)
        SetupVar.dAct03VoltHi(i) = Val(frmSetup.txtAct03VoltHi(i).Text)
        SetupVar.dAct03TimeLo(i) = Val(frmSetup.txtAct03TimeLo(i).Text)
        SetupVar.dAct03TimeHi(i) = Val(frmSetup.txtAct03TimeHi(i).Text)
    Next
    
    SetupVar.dAct03StallDeltaVoltLo = Val(frmSetup.txtAct03StallDeltaMinVolt.Text)
    SetupVar.dAct03StallDeltaVoltHi = Val(frmSetup.txtAct03StallDeltaMaxVolt.Text)
    
    ' Act 04
    SetupVar.bAct04Use = IIf(frmSetup.chkAct04.Value = 1, True, False)
    SetupVar.nAct04TestType = IIf(frmSetup.optAct04TestType(0).Value, 0, 1)
    
    For i = 0 To frmSetup.opt2Pin.UBound
        If frmSetup.opt2Pin(i).Value = True Then
            SetupVar.nAct042Pin = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.opt2PinPos.UBound
        If frmSetup.opt2PinPos(i).Value = True Then
            SetupVar.nAct042PinPos = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.optAct04Direction.UBound
        If frmSetup.optAct04Direction(i).Value = True Then
            SetupVar.nAct04Direction = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.txtAct04Name.UBound
        SetupVar.lpAct04Name(i) = Trim$(frmSetup.txtAct04Name(i).Text)
        SetupVar.dAct04SetVolt(i) = Val(frmSetup.txtAct04SetVolt(i).Text)
        SetupVar.dAct04CurrLo(i) = Val(frmSetup.txtAct04CurrLo(i).Text)
        SetupVar.dAct04CurrHi(i) = Val(frmSetup.txtAct04CurrHi(i).Text)
        SetupVar.dAct04VoltLo(i) = Val(frmSetup.txtAct04VoltLo(i).Text)
        SetupVar.dAct04VoltHi(i) = Val(frmSetup.txtAct04VoltHi(i).Text)
        SetupVar.dAct04TimeLo(i) = Val(frmSetup.txtAct04TimeLo(i).Text)
        SetupVar.dAct04TimeHi(i) = Val(frmSetup.txtAct04TimeHi(i).Text)
    Next
    
    SetupVar.dAct04StallDeltaVoltLo = Val(frmSetup.txtAct04StallDeltaMinVolt.Text)
    SetupVar.dAct04StallDeltaVoltHi = Val(frmSetup.txtAct04StallDeltaMaxVolt.Text)
    
    ' Ion
    SetupVar.lpIonName = Trim$(frmSetup.txtIonName.Text)
    SetupVar.bIonUse = IIf(frmSetup.chkIonUse.Value = 1, True, False)
    
    For i = 0 To frmSetup.txtIonSubName.UBound
        SetupVar.lpIonSubName(i) = Trim$(frmSetup.txtIonSubName(i).Text)
        SetupVar.dIonHi(i) = Val(frmSetup.txtIonHi(i).Text)
        SetupVar.dIonLo(i) = Val(frmSetup.txtIonLo(i).Text)
    Next
    
    ' Sensor
    For i = 0 To frmSetup.chkSensor.UBound
        SetupVar.bSensorUse(i) = IIf(frmSetup.chkSensor(i).Value = 1, True, False)
        SetupVar.lpSensorName(i) = Trim$(frmSetup.txtSensorName(i).Text)
        SetupVar.dSensorCurrLo(i) = Val(frmSetup.txtSensorCurrLo(i).Text)
        SetupVar.dSensorCurrHi(i) = Val(frmSetup.txtSensorCurrHi(i).Text)
        SetupVar.dSensorTime(i) = Val(frmSetup.txtSensorTime(i).Text)
    Next
    
    ' PTC
    SetupVar.bPTCUse = IIf(frmSetup.chkPTC.Value = 1, True, False)
    SetupVar.lpPTCName = Trim$(frmSetup.txtPTCName.Text)
    SetupVar.dPTCCurrLo = Val(frmSetup.txtPTCCurrLo.Text)
    SetupVar.dPTCCurrHi = Val(frmSetup.txtPTCCurrHi.Text)
    SetupVar.dPTCTime = Val(frmSetup.txtPTCTime.Text)
    
    ' Leak
    For i = 0 To frmSetup.chkLeak.UBound
        SetupVar.bLeakUse(i) = IIf(frmSetup.chkLeak(i).Value = 1, True, False)
        SetupVar.lpLeakName(i) = Trim$(frmSetup.txtLeakName(i).Text)
    Next
    
    SetupVar.nLeakModel = Val(frmSetup.txtLeakModel.Text)
    
    ' Vision
    SetupVar.bVisionUse = IIf(frmSetup.chkVision.Value = 1, True, False)
    SetupVar.lpVisionName(0) = Trim$(frmSetup.txtVisionName(0).Text)
    
    For i = 0 To frmSetup.chkVisionDoor.UBound
        SetupVar.bVisionDoorUse(i) = IIf(frmSetup.chkVisionDoor(i).Value = 1, True, False)
    Next
    
    For i = 0 To frmSetup.txtVisionDoorName.UBound
        SetupVar.lpVisionDoorName(i) = frmSetup.txtVisionDoorName(i).Text
    Next
    
    SetupVar.nOpenCameraNo(0) = frmSetup.cboOpenCameraNo1.ListIndex
    For i = 0 To frmSetup.optOpenCameraPos1.UBound
        If frmSetup.optOpenCameraPos1(i).Value = True Then
            SetupVar.nOpenCameraPos(0) = i
            Exit For
        End If
    Next
    
    SetupVar.nOpenCameraNo(1) = frmSetup.cboOpenCameraNo2.ListIndex
    For i = 0 To frmSetup.optOpenCameraPos2.UBound
        If frmSetup.optOpenCameraPos2(i).Value = True Then
            SetupVar.nOpenCameraPos(1) = i
            Exit For
        End If
    Next
    
    SetupVar.nOpenCameraNo(2) = frmSetup.cboOpenCameraNo3.ListIndex
    For i = 0 To frmSetup.optOpenCameraPos3.UBound
        If frmSetup.optOpenCameraPos3(i).Value = True Then
            SetupVar.nOpenCameraPos(2) = i
            Exit For
        End If
    Next
    
    SetupVar.nOpenCameraNo(3) = frmSetup.cboOpenCameraNo4.ListIndex
    For i = 0 To frmSetup.optOpenCameraPos4.UBound
        If frmSetup.optOpenCameraPos4(i).Value = True Then
            SetupVar.nOpenCameraPos(3) = i
            Exit For
        End If
    Next
    
    SetupVar.nCloseCameraNo(0) = frmSetup.cboCloseCameraNo1.ListIndex
    For i = 0 To frmSetup.optCloseCameraPos1.UBound
        If frmSetup.optCloseCameraPos1(i).Value = True Then
            SetupVar.nCloseCameraPos(0) = i
            Exit For
        End If
    Next
    
    SetupVar.nCloseCameraNo(1) = frmSetup.cboCloseCameraNo2.ListIndex
    For i = 0 To frmSetup.optCloseCameraPos2.UBound
        If frmSetup.optCloseCameraPos2(i).Value = True Then
            SetupVar.nCloseCameraPos(1) = i
            Exit For
        End If
    Next
    
    SetupVar.nCloseCameraNo(2) = frmSetup.cboCloseCameraNo3.ListIndex
    For i = 0 To frmSetup.optCloseCameraPos3.UBound
        If frmSetup.optCloseCameraPos3(i).Value = True Then
            SetupVar.nCloseCameraPos(2) = i
            Exit For
        End If
    Next
    
    SetupVar.nCloseCameraNo(3) = frmSetup.cboCloseCameraNo4.ListIndex
    For i = 0 To frmSetup.optCloseCameraPos4.UBound
        If frmSetup.optCloseCameraPos4(i).Value = True Then
            SetupVar.nCloseCameraPos(3) = i
            Exit For
        End If
    Next
    
    ' Barcode Print
    SetupVar.bBarCodeUse = IIf(frmSetup.chkBarCodePrintUse.Value = 1, True, False)
    
    For i = 0 To 1
        If frmSetup.optBarcodeType(i).Value = True Then
            SetupVar.nBarcodeType = i
            Exit For
        End If
    Next
    
    For i = 0 To 4
        SetupVar.lpBarcode(i) = Trim$(frmSetup.txtBarCode(i).Text)
    Next
    
    ' Marking
    For i = 0 To 0
        SetupVar.bMarkingUse(i) = IIf(frmSetup.chkMarking(i).Value = 1, True, False)
        SetupVar.dMarkingTime(i) = Val(frmSetup.txtMarkingTime(i).Text)
    Next
    
    ' Part
    For i = 0 To MAX_DIO_CHANNEL
        SetupVar.bPartUse(i) = IIf(frmSetup.chkPart(i).Value = 1, True, False)
        SetupVar.bPartStatus(i) = IIf(frmSetup.lblPart(i).BackColor = vbGreen, True, False)
        SetupVar.lpPartName(i) = Trim$(frmSetup.txtPart(i).Text)
    Next
    
    SetupVar.bProductUse = IIf(frmSetup.chkProductUse.Value = 1, True, False)
    SetupVar.lpProductList = frmSetup.txtProductList.Text
    SetupVar.lpProductName = frmSetup.txtProductName.Text
    SetupVar.bModelTypeUse = IIf(frmSetup.chkModelTypeUse.Value = 1, True, False)
    SetupVar.lpModelLHDList = frmSetup.txtModelLHDList.Text
    SetupVar.lpLHDPartName = frmSetup.txtLHDPartName.Text
    SetupVar.lpModelRHDList = frmSetup.txtModelRHDList.Text
    SetupVar.lpRHDPartName = frmSetup.txtRHDPartName.Text
    
    ' Lin
    For i = 0 To 3
        SetupVar.bLinActUse(i) = IIf(frmSetup.chkLinAct(i).Value = 1, True, False)
        SetupVar.lpLinActName(i) = Trim$(frmSetup.txtLinActName(i).Text)
        SetupVar.dLinActLo(i) = frmSetup.txtLinActLo(i).Text
        SetupVar.dLinActHi(i) = frmSetup.txtLinActHi(i).Text
        SetupVar.nLinActMove(i) = frmSetup.txtLinActMove(i).Text
        SetupVar.nLinActFinal(i) = frmSetup.txtLinActFinal(i).Text
        SetupVar.lLinActAngle(i) = frmSetup.txtLinActAngle(i).Text
        SetupVar.dLinActTime(i) = frmSetup.txtLinActTime(i).Text
        SetupVar.dLinActCurrLo(i) = frmSetup.txtLinActCurrLo(i).Text
        SetupVar.dLinActCurrHi(i) = frmSetup.txtLinActCurrHi(i).Text
    Next
    
    For i = 0 To 4
        SetupVar.nLinAct01Check(i) = frmSetup.txtLinAct01Check(i).Text
        SetupVar.nLinAct02Check(i) = frmSetup.txtLinAct02Check(i).Text
        SetupVar.nLinAct03Check(i) = frmSetup.txtLinAct03Check(i).Text
        SetupVar.nLinAct04Check(i) = frmSetup.txtLinAct04Check(i).Text
        SetupVar.dLinAct01CheckTime(i) = frmSetup.txtLinAct01CheckTime(i).Text
        SetupVar.dLinAct02CheckTime(i) = frmSetup.txtLinAct02CheckTime(i).Text
        SetupVar.dLinAct03CheckTime(i) = frmSetup.txtLinAct03CheckTime(i).Text
        SetupVar.dLinAct04CheckTime(i) = frmSetup.txtLinAct04CheckTime(i).Text
    Next
    
    For i = 0 To frmSetup.optLinAct01FirstMove.UBound
        If frmSetup.optLinAct01FirstMove(i).Value = True Then
            SetupVar.nLinActFirstMove(0) = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.optLinAct02FirstMove.UBound
        If frmSetup.optLinAct02FirstMove(i).Value = True Then
            SetupVar.nLinActFirstMove(1) = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.optLinAct03FirstMove.UBound
        If frmSetup.optLinAct03FirstMove(i).Value = True Then
            SetupVar.nLinActFirstMove(2) = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.optLinAct04FirstMove.UBound
        If frmSetup.optLinAct04FirstMove(i).Value = True Then
            SetupVar.nLinActFirstMove(3) = i
            Exit For
        End If
    Next
    
    SetupVar.bAutoAddressUse = IIf(frmSetup.chkAutoAddress.Value = 1, True, False)
    
    For i = 0 To frmSetup.optLinTestType.UBound
        If frmSetup.optLinTestType(i).Value = True Then
            SetupVar.nLinTestType = i
            Exit For
        End If
    Next
    
    For i = 0 To frmSetup.optAct01RefPos.UBound
        If frmSetup.optAct01RefPos(i).Value = True Then
            SetupVar.nLinAct01RefPos = i
            Exit For
        End If
    Next
    
    SetupVar.bCheckPointUse = IIf(frmSetup.chkCheckPoint.Value = 1, True, False)
    SetupVar.bStallUse = IIf(frmSetup.chkStallUse.Value = 1, True, False)
    
    ' Curr Count
    For i = 0 To frmSetup.txtActPeakCurrCount.UBound
        SetupVar.nActPeakCurrCount(i) = Val(frmSetup.txtActPeakCurrCount(i).Text)
        SetupVar.nActEndPosCount(i) = Val(frmSetup.txtActEndPosCount(i).Text)
    Next
    
    ' Stepping Set
    For i = 0 To frmSetup.txtStepAct01Arr.UBound
        SetupVar.nSteppingSet(0, i) = Trim$(frmSetup.txtStepAct01Arr(i).Text)
        SetupVar.nSteppingSet(1, i) = Trim$(frmSetup.txtStepAct02Arr(i).Text)
        SetupVar.nSteppingSet(2, i) = Trim$(frmSetup.txtStepAct03Arr(i).Text)
        SetupVar.nSteppingSet(3, i) = Trim$(frmSetup.txtStepAct04Arr(i).Text)
    Next
    
    ' Adjustment
    For i = 0 To frmSetup.chkAdjustUse.UBound
        SetupVar.bAdjustUse(i) = IIf(frmSetup.chkAdjustUse(i).Value = 1, True, False)
        SetupVar.dAdd(i) = Val(frmSetup.txtAdd(i).Text)
        SetupVar.dMulti(i) = Val(frmSetup.txtMulti(i).Text)
    Next
End Sub

Public Sub SetupMem2Disp()
    Dim i As Integer
    
    ' General
    frmSetup.txtTestVolt.Text = Format(Trim$(SetupVar.dTestVolt), "#0.0")
    frmSetup.txtFileName.Text = Trim$(SetupVar.lpFileName)
    frmSetup.chkSaveUse.Value = IIf(SetupVar.bDataSave, 1, 0)
    frmSetup.chkScannerUse.Value = IIf(SetupVar.bScannerUse, 1, 0)
    frmSetup.txtScannerValue = Trim$(SetupVar.nScannerValue)
    frmSetup.chkNvh.Value = IIf(SetupVar.bNvhUse, 1, 0)
    frmSetup.optModelType(SetupVar.nModelType).Value = 1
    
    For i = 0 To frmSetup.txtActName.UBound
        frmSetup.txtActName(i).Text = Trim$(SetupVar.lpActName(i))
        frmSetup.txtActBoardNo(i).Text = Trim$(SetupVar.nActBoardNo(i))
    Next
    
    ' Blower
    frmSetup.chkBlower.Value = IIf(SetupVar.bBlowerUse, 1, 0)
    frmSetup.optBlowerType(SetupVar.nBlowerType).Value = 1
    frmSetup.optBlowerDirection(SetupVar.nBlowerDirection).Value = 1
    
    For i = 0 To frmSetup.txtBlowerName.UBound
        frmSetup.txtBlowerName(i).Text = Trim$(SetupVar.lpBlowerName(i))
        frmSetup.txtBlowerCurrLo(i).Text = Trim$(SetupVar.dBlowerCurrLo(i))
        frmSetup.txtBlowerCurrHi(i).Text = Trim$(SetupVar.dBlowerCurrHi(i))
        frmSetup.txtBlowerTime(i).Text = Trim$(SetupVar.dBlowerTime(i))
        frmSetup.txtLinSpeed(i).Text = Trim$(SetupVar.nLinSpeed(i))
    Next
    
    ' RPM
    frmSetup.txtRpmName.Text = Trim$(SetupVar.lpRpmName)
    frmSetup.txtRpmCurrLo.Text = Trim$(SetupVar.dRpmCurrLo)
    frmSetup.txtRpmCurrHi.Text = Trim$(SetupVar.dRpmCurrHi)
    
    ' Vibration
    frmSetup.chkVib.Value = IIf(SetupVar.bVibUse = True, 1, 0)
    frmSetup.txtVibName.Text = Trim$(SetupVar.lpVibName)
    frmSetup.txtVibCurrLo.Text = Trim$(SetupVar.dVibCurrLo)
    frmSetup.txtVibCurrHi.Text = Trim$(SetupVar.dVibCurrHi)
    frmSetup.optVibResultType(SetupVar.nVibResultType).Value = 1
    frmSetup.txtVibStart.Text = Trim$(SetupVar.dVibStart)
    frmSetup.txtVibEnd.Text = Trim$(SetupVar.dVibEnd)
    frmSetup.optVibMethod(SetupVar.nVibMethod).Value = True
    frmSetup.txtVibVolt.Text = Trim$(SetupVar.dVibVolt)
    frmSetup.txtVibTime.Text = Trim$(SetupVar.dVibTime)
    
    ' Act 01
    frmSetup.chkAct01.Value = IIf(SetupVar.bAct01Use = True, 1, 0)
    frmSetup.optAct01Direction(SetupVar.nAct01Direction).Value = True
    frmSetup.optAct01TestType(SetupVar.nAct01TestType).Value = 1
    
    For i = 0 To frmSetup.txtAct01Name.UBound
        frmSetup.txtAct01Name(i).Text = Trim$(SetupVar.lpAct01Name(i))
        frmSetup.txtAct01SetVolt(i).Text = Trim$(SetupVar.dAct01SetVolt(i))
        frmSetup.txtAct01CurrLo(i).Text = Trim$(SetupVar.dAct01CurrLo(i))
        frmSetup.txtAct01CurrHi(i).Text = Trim$(SetupVar.dAct01CurrHi(i))
        frmSetup.txtAct01VoltLo(i).Text = Trim$(SetupVar.dAct01VoltLo(i))
        frmSetup.txtAct01VoltHi(i).Text = Trim$(SetupVar.dAct01VoltHi(i))
        frmSetup.txtAct01TimeLo(i).Text = Trim$(SetupVar.dAct01TimeLo(i))
        frmSetup.txtAct01TimeHi(i).Text = Trim$(SetupVar.dAct01TimeHi(i))
    Next
    
    frmSetup.txtAct01StallDeltaMinVolt.Text = Trim$(SetupVar.dAct01StallDeltaVoltLo)
    frmSetup.txtAct01StallDeltaMaxVolt.Text = Trim$(SetupVar.dAct01StallDeltaVoltHi)
    
    ' Act 02
    frmSetup.chkAct02.Value = IIf(SetupVar.bAct02Use = True, 1, 0)
    frmSetup.optAct02Direction(SetupVar.nAct02Direction).Value = True
    frmSetup.optAct02TestType(SetupVar.nAct02TestType).Value = 1
    
    For i = 0 To frmSetup.txtAct02Name.UBound
        frmSetup.txtAct02Name(i).Text = Trim$(SetupVar.lpAct02Name(i))
        frmSetup.txtAct02SetVolt(i).Text = Trim$(SetupVar.dAct02SetVolt(i))
        frmSetup.txtAct02CurrLo(i).Text = Trim$(SetupVar.dAct02CurrLo(i))
        frmSetup.txtAct02CurrHi(i).Text = Trim$(SetupVar.dAct02CurrHi(i))
        frmSetup.txtAct02VoltLo(i).Text = Trim$(SetupVar.dAct02VoltLo(i))
        frmSetup.txtAct02VoltHi(i).Text = Trim$(SetupVar.dAct02VoltHi(i))
        frmSetup.txtAct02TimeLo(i).Text = Trim$(SetupVar.dAct02TimeLo(i))
        frmSetup.txtAct02TimeHi(i).Text = Trim$(SetupVar.dAct02TimeHi(i))
    Next
    
    frmSetup.txtAct02StallDeltaMinVolt.Text = Trim$(SetupVar.dAct02StallDeltaVoltLo)
    frmSetup.txtAct02StallDeltaMaxVolt.Text = Trim$(SetupVar.dAct02StallDeltaVoltHi)
    
    ' Act 03
    frmSetup.chkAct03.Value = IIf(SetupVar.bAct03Use = True, 1, 0)
    frmSetup.optAct03Direction(SetupVar.nAct03Direction).Value = True
    frmSetup.optAct03TestType(SetupVar.nAct03TestType).Value = 1
    
    For i = 0 To frmSetup.txtAct03Name.UBound
        frmSetup.txtAct03Name(i).Text = Trim$(SetupVar.lpAct03Name(i))
        frmSetup.txtAct03SetVolt(i).Text = Trim$(SetupVar.dAct03SetVolt(i))
        frmSetup.txtAct03CurrLo(i).Text = Trim$(SetupVar.dAct03CurrLo(i))
        frmSetup.txtAct03CurrHi(i).Text = Trim$(SetupVar.dAct03CurrHi(i))
        frmSetup.txtAct03VoltLo(i).Text = Trim$(SetupVar.dAct03VoltLo(i))
        frmSetup.txtAct03VoltHi(i).Text = Trim$(SetupVar.dAct03VoltHi(i))
        frmSetup.txtAct03TimeLo(i).Text = Trim$(SetupVar.dAct03TimeLo(i))
        frmSetup.txtAct03TimeHi(i).Text = Trim$(SetupVar.dAct03TimeHi(i))
    Next
    
    frmSetup.txtAct03StallDeltaMinVolt.Text = Trim$(SetupVar.dAct03StallDeltaVoltLo)
    frmSetup.txtAct03StallDeltaMaxVolt.Text = Trim$(SetupVar.dAct03StallDeltaVoltHi)
    
    ' Act 04
    frmSetup.chkAct04.Value = IIf(SetupVar.bAct04Use = True, 1, 0)
    frmSetup.opt2Pin(SetupVar.nAct042Pin).Value = True
    frmSetup.opt2PinPos(SetupVar.nAct042PinPos).Value = True
    frmSetup.optAct04Direction(SetupVar.nAct04Direction).Value = True
    frmSetup.optAct04TestType(SetupVar.nAct04TestType).Value = 1
    
    For i = 0 To frmSetup.txtAct04Name.UBound
        frmSetup.txtAct04Name(i).Text = Trim$(SetupVar.lpAct04Name(i))
        frmSetup.txtAct04SetVolt(i).Text = Trim$(SetupVar.dAct04SetVolt(i))
        frmSetup.txtAct04CurrLo(i).Text = Trim$(SetupVar.dAct04CurrLo(i))
        frmSetup.txtAct04CurrHi(i).Text = Trim$(SetupVar.dAct04CurrHi(i))
        frmSetup.txtAct04VoltLo(i).Text = Trim$(SetupVar.dAct04VoltLo(i))
        frmSetup.txtAct04VoltHi(i).Text = Trim$(SetupVar.dAct04VoltHi(i))
        frmSetup.txtAct04TimeLo(i).Text = Trim$(SetupVar.dAct04TimeLo(i))
        frmSetup.txtAct04TimeHi(i).Text = Trim$(SetupVar.dAct04TimeHi(i))
    Next
    
    frmSetup.txtAct04StallDeltaMinVolt.Text = Trim$(SetupVar.dAct04StallDeltaVoltLo)
    frmSetup.txtAct04StallDeltaMaxVolt.Text = Trim$(SetupVar.dAct04StallDeltaVoltHi)
    
    ' Ion
    frmSetup.txtIonName.Text = Trim$(SetupVar.lpIonName)
    frmSetup.chkIonUse.Value = IIf(SetupVar.bIonUse, 1, 0)
    
    For i = 0 To frmSetup.txtIonSubName.UBound
        frmSetup.txtIonSubName(i).Text = Trim$(SetupVar.lpIonSubName(i))
        frmSetup.txtIonHi(i).Text = Trim$(SetupVar.dIonHi(i))
        frmSetup.txtIonLo(i).Text = Trim$(SetupVar.dIonLo(i))
    Next
    
    ' Sensor
    For i = 0 To frmSetup.chkSensor.UBound
        frmSetup.chkSensor(i).Value = IIf(SetupVar.bSensorUse(i), 1, 0)
        frmSetup.txtSensorName(i).Text = Trim$(SetupVar.lpSensorName(i))
        frmSetup.txtSensorCurrLo(i).Text = Trim$(SetupVar.dSensorCurrLo(i))
        frmSetup.txtSensorCurrHi(i).Text = Trim$(SetupVar.dSensorCurrHi(i))
        frmSetup.txtSensorTime(i).Text = Trim$(SetupVar.dSensorTime(i))
    Next
    
    ' PTC
    frmSetup.chkPTC.Value = IIf(SetupVar.bPTCUse, 1, 0)
    frmSetup.txtPTCName.Text = Trim$(SetupVar.lpPTCName)
    frmSetup.txtPTCCurrLo.Text = Trim$(SetupVar.dPTCCurrLo)
    frmSetup.txtPTCCurrHi.Text = Trim$(SetupVar.dPTCCurrHi)
    frmSetup.txtPTCTime.Text = Trim$(SetupVar.dPTCTime)
    
    ' Leak
    For i = 0 To frmSetup.chkLeak.UBound
        frmSetup.chkLeak(i).Value = IIf(SetupVar.bLeakUse(i), 1, 0)
        frmSetup.txtLeakName(i).Text = Trim$(SetupVar.lpLeakName(i))
    Next
    
    frmSetup.txtLeakModel.Text = Trim$(SetupVar.nLeakModel)
    
    ' Vision
    frmSetup.chkVision.Value = IIf(SetupVar.bVisionUse = True, 1, 0)
    frmSetup.txtVisionName(0).Text = Trim$(SetupVar.lpVisionName(0))
    
    For i = 0 To frmSetup.chkVisionDoor.UBound
        frmSetup.chkVisionDoor(i).Value = IIf(SetupVar.bVisionDoorUse(i) = True, 1, 0)
    Next
    
    For i = 0 To frmSetup.txtVisionDoorName.UBound
        frmSetup.txtVisionDoorName(i).Text = SetupVar.lpVisionDoorName(i)
    Next
    
    frmSetup.cboOpenCameraNo1.ListIndex = SetupVar.nOpenCameraNo(0)
    frmSetup.optOpenCameraPos1(SetupVar.nOpenCameraPos(0)).Value = True
    frmSetup.cboOpenCameraNo2.ListIndex = SetupVar.nOpenCameraNo(1)
    frmSetup.optOpenCameraPos2(SetupVar.nOpenCameraPos(1)).Value = True
    frmSetup.cboOpenCameraNo3.ListIndex = SetupVar.nOpenCameraNo(2)
    frmSetup.optOpenCameraPos3(SetupVar.nOpenCameraPos(2)).Value = True
    frmSetup.cboOpenCameraNo4.ListIndex = SetupVar.nOpenCameraNo(3)
    frmSetup.optOpenCameraPos4(SetupVar.nOpenCameraPos(3)).Value = True
    
    frmSetup.cboCloseCameraNo1.ListIndex = SetupVar.nCloseCameraNo(0)
    frmSetup.optCloseCameraPos1(SetupVar.nCloseCameraPos(0)).Value = True
    frmSetup.cboCloseCameraNo2.ListIndex = SetupVar.nCloseCameraNo(1)
    frmSetup.optCloseCameraPos2(SetupVar.nCloseCameraPos(1)).Value = True
    frmSetup.cboCloseCameraNo3.ListIndex = SetupVar.nCloseCameraNo(2)
    frmSetup.optCloseCameraPos3(SetupVar.nCloseCameraPos(2)).Value = True
    frmSetup.cboCloseCameraNo4.ListIndex = SetupVar.nCloseCameraNo(3)
    frmSetup.optCloseCameraPos4(SetupVar.nCloseCameraPos(3)).Value = True
    
    ' Barcode Print
    frmSetup.chkBarCodePrintUse.Value = IIf(SetupVar.bBarCodeUse = True, 1, 0)
    frmSetup.optBarcodeType(SetupVar.nBarcodeType).Value = True
    
    For i = 0 To 4
        frmSetup.txtBarCode(i).Text = Trim$(SetupVar.lpBarcode(i))
    Next
    
    ' Marking
    For i = 0 To 0
        frmSetup.chkMarking(i).Value = IIf(SetupVar.bMarkingUse(i) = True, 1, 0)
        frmSetup.txtMarkingTime(i).Text = Trim$(SetupVar.dMarkingTime(i))
    Next
    
    ' Part
    For i = 0 To MAX_DIO_CHANNEL
        frmSetup.chkPart(i).Value = IIf(SetupVar.bPartUse(i) = True, 1, 0)
        
        If SetupVar.bPartUse(i) Then
            frmSetup.lblPart(i).BackColor = IIf(SetupVar.bPartStatus(i) = True, vbGreen, vbRed)
            
            If SetupVar.bPartStatus(i) Then
                frmSetup.lblPart(i).Caption = "ON"
            Else
                frmSetup.lblPart(i).Caption = "OFF"
            End If
        Else
            frmSetup.lblPart(i).BackColor = CO_NONE
            frmSetup.lblPart(i).Caption = "--"
        End If
    
        frmSetup.txtPart(i).Text = Trim$(SetupVar.lpPartName(i))
    Next
    
    frmSetup.chkProductUse.Value = IIf(SetupVar.bProductUse = True, 1, 0)
    frmSetup.txtProductList.Text = Trim$(SetupVar.lpProductList)
    frmSetup.txtProductName.Text = Trim$(SetupVar.lpProductName)
    frmSetup.chkModelTypeUse.Value = IIf(SetupVar.bModelTypeUse = True, 1, 0)
    frmSetup.txtModelLHDList.Text = Trim$(SetupVar.lpModelLHDList)
    frmSetup.txtLHDPartName.Text = Trim$(SetupVar.lpLHDPartName)
    frmSetup.txtModelRHDList.Text = Trim$(SetupVar.lpModelRHDList)
    frmSetup.txtRHDPartName.Text = Trim$(SetupVar.lpRHDPartName)
    
    ' Lin
    For i = 0 To 3
        frmSetup.chkLinAct(i).Value = IIf(SetupVar.bLinActUse(i), 1, 0)
        frmSetup.txtLinActName(i).Text = Trim$(SetupVar.lpLinActName(i))
        frmSetup.txtLinActLo(i).Text = Trim$(SetupVar.dLinActLo(i))
        frmSetup.txtLinActHi(i).Text = Trim$(SetupVar.dLinActHi(i))
        frmSetup.txtLinActFinal(i).Text = Trim$(SetupVar.nLinActFinal(i))
        frmSetup.txtLinActMove(i).Text = Trim$(SetupVar.nLinActMove(i))
        frmSetup.txtLinActAngle(i).Text = Trim$(SetupVar.lLinActAngle(i))
        frmSetup.txtLinActTime(i).Text = Trim$(SetupVar.dLinActTime(i))
        frmSetup.txtLinActCurrLo(i).Text = Trim$(SetupVar.dLinActCurrLo(i))
        frmSetup.txtLinActCurrHi(i).Text = Trim$(SetupVar.dLinActCurrHi(i))
    Next
    
    For i = 0 To 4
        frmSetup.txtLinAct01Check(i).Text = Trim$(SetupVar.nLinAct01Check(i))
        frmSetup.txtLinAct02Check(i).Text = Trim$(SetupVar.nLinAct02Check(i))
        frmSetup.txtLinAct03Check(i).Text = Trim$(SetupVar.nLinAct03Check(i))
        frmSetup.txtLinAct04Check(i).Text = Trim$(SetupVar.nLinAct04Check(i))
        frmSetup.txtLinAct01CheckTime(i).Text = Trim$(SetupVar.dLinAct01CheckTime(i))
        frmSetup.txtLinAct02CheckTime(i).Text = Trim$(SetupVar.dLinAct02CheckTime(i))
        frmSetup.txtLinAct03CheckTime(i).Text = Trim$(SetupVar.dLinAct03CheckTime(i))
        frmSetup.txtLinAct04CheckTime(i).Text = Trim$(SetupVar.dLinAct04CheckTime(i))
    Next
    
    frmSetup.chkAutoAddress.Value = IIf(SetupVar.bAutoAddressUse, 1, 0)
    frmSetup.optLinTestType(SetupVar.nLinTestType).Value = 1
    
    frmSetup.optLinAct01FirstMove(SetupVar.nLinActFirstMove(0)).Value = 1
    frmSetup.optLinAct02FirstMove(SetupVar.nLinActFirstMove(1)).Value = 1
    frmSetup.optLinAct03FirstMove(SetupVar.nLinActFirstMove(2)).Value = 1
    frmSetup.optLinAct04FirstMove(SetupVar.nLinActFirstMove(3)).Value = 1
    
    frmSetup.optAct01RefPos(SetupVar.nLinAct01RefPos).Value = 1
    
    frmSetup.chkCheckPoint.Value = IIf(SetupVar.bCheckPointUse, 1, 0)
    frmSetup.chkStallUse.Value = IIf(SetupVar.bStallUse, 1, 0)
    
    ' Curr Count
    For i = 0 To frmSetup.txtActPeakCurrCount.UBound
        frmSetup.txtActPeakCurrCount(i).Text = Trim$(SetupVar.nActPeakCurrCount(i))
        frmSetup.txtActEndPosCount(i).Text = Trim$(SetupVar.nActEndPosCount(i))
    Next
    
    ' Stepping Set
    For i = 0 To frmSetup.txtStepAct01Arr.UBound
        frmSetup.txtStepAct01Arr(i).Text = Trim$(SetupVar.nSteppingSet(0, i))
        frmSetup.txtStepAct02Arr(i).Text = Trim$(SetupVar.nSteppingSet(1, i))
        frmSetup.txtStepAct03Arr(i).Text = Trim$(SetupVar.nSteppingSet(2, i))
        frmSetup.txtStepAct04Arr(i).Text = Trim$(SetupVar.nSteppingSet(3, i))
    Next
    
    ' Adjustment
    For i = 0 To frmSetup.chkAdjustUse.UBound
        frmSetup.chkAdjustUse(i).Value = IIf(SetupVar.bAdjustUse(i) = True, 1, 0)
        frmSetup.txtAdd(i).Text = Format(SetupVar.dAdd(i), "#0.00")
        frmSetup.txtMulti(i).Text = Format(SetupVar.dMulti(i), "#0.00")
    Next
End Sub

Public Function LoadSetupFile(ByVal lpModelName As String) As Boolean
    Dim i               As Integer
    Dim lpFileName      As String
    
    On Error Resume Next
    
    lpFileName = lpPath & INISETUPFILE
    LoadSetupFile = SearchFile(lpFileName)
    
    ' General
    SetupVar.dTestVolt = INIRead(lpModelName, "TestVolt", lpFileName)
    SetupVar.lpFileName = INIRead(lpModelName, "FileName", lpFileName)
    SetupVar.bDataSave = IIf(INIRead(lpModelName, "NoDataSave", lpFileName) = "True", True, False)
    SetupVar.bScannerUse = IIf(INIRead(lpModelName, "ScannerUse", lpFileName) = "True", True, False)
    SetupVar.nScannerValue = INIRead(lpModelName, "ScannerValue", lpFileName)
    SetupVar.bNvhUse = IIf(INIRead(lpModelName, "NvhUse", lpFileName) = "True", True, False)
    SetupVar.nModelType = INIRead(lpModelName, "ModelType", lpFileName)
    
    For i = 0 To frmSetup.txtActName.UBound
        SetupVar.lpActName(i) = INIRead(lpModelName, "ActName" & Format(i, "00"), lpFileName)
        SetupVar.nActBoardNo(i) = INIRead(lpModelName, "ActBoardNo" & Format(i, "00"), lpFileName)
    Next
    
    ' Blower
    SetupVar.bBlowerUse = IIf(INIRead(lpModelName, "BlowerUse", lpFileName) = "True", True, False)
    SetupVar.nBlowerType = INIRead(lpModelName, "BlowerType", lpFileName)
    SetupVar.nBlowerDirection = INIRead(lpModelName, "BlowerDirection", lpFileName)
    
    For i = 0 To frmSetup.txtBlowerName.UBound
        SetupVar.lpBlowerName(i) = INIRead(lpModelName, "BlowerName" & Format(i, "00"), lpFileName)
        SetupVar.dBlowerCurrLo(i) = INIRead(lpModelName, "BlowerCurrLo" & Format(i, "00"), lpFileName)
        SetupVar.dBlowerCurrHi(i) = INIRead(lpModelName, "BlowerCurrHi" & Format(i, "00"), lpFileName)
        SetupVar.dBlowerTime(i) = INIRead(lpModelName, "BlowerTime" & Format(i, "00"), lpFileName)
        SetupVar.nLinSpeed(i) = INIRead(lpModelName, "LinSpeed" & Format(i, "00"), lpFileName)
    Next
    
    ' RPM
    SetupVar.lpRpmName = INIRead(lpModelName, "RpmName", lpFileName)
    SetupVar.dRpmCurrLo = INIRead(lpModelName, "RpmCurrLo", lpFileName)
    SetupVar.dRpmCurrHi = INIRead(lpModelName, "RpmCurrHi", lpFileName)
    
    ' Vibration
    SetupVar.bVibUse = IIf(INIRead(lpModelName, "VibUse", lpFileName) = "True", True, False)
    SetupVar.lpVibName = INIRead(lpModelName, "VibName", lpFileName)
    SetupVar.dVibCurrLo = INIRead(lpModelName, "VibCurrLo", lpFileName)
    SetupVar.dVibCurrHi = INIRead(lpModelName, "VibCurrHi", lpFileName)
    SetupVar.nVibResultType = INIRead(lpModelName, "VibResultType", lpFileName)
    SetupVar.dVibStart = INIRead(lpModelName, "VibStart", lpFileName)
    SetupVar.dVibEnd = INIRead(lpModelName, "VibEnd", lpFileName)
    SetupVar.nVibMethod = INIRead(lpModelName, "VibMethod", lpFileName)
    SetupVar.dVibVolt = INIRead(lpModelName, "VibVolt", lpFileName)
    SetupVar.dVibTime = INIRead(lpModelName, "VibTime", lpFileName)
    
    ' Act 01
    SetupVar.bAct01Use = INIRead(lpModelName, "Act01Use", lpFileName)
    SetupVar.nAct01Direction = INIRead(lpModelName, "Act01Direction", lpFileName)
    SetupVar.nAct01TestType = INIRead(lpModelName, "Act01TestType", lpFileName)
    
    For i = 0 To frmSetup.txtAct01Name.UBound
        SetupVar.lpAct01Name(i) = INIRead(lpModelName, "Act01Name" & Format(i, "00"), lpFileName)
        SetupVar.dAct01SetVolt(i) = INIRead(lpModelName, "Act01SetVolt" & Format(i, "00"), lpFileName)
        SetupVar.dAct01CurrLo(i) = INIRead(lpModelName, "Act01CurrLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct01CurrHi(i) = INIRead(lpModelName, "Act01CurrHi" & Format(i, "00"), lpFileName)
        SetupVar.dAct01VoltLo(i) = INIRead(lpModelName, "Act01VoltLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct01VoltHi(i) = INIRead(lpModelName, "Act01VoltHi" & Format(i, "00"), lpFileName)
        SetupVar.dAct01TimeLo(i) = INIRead(lpModelName, "Act01TimeLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct01TimeHi(i) = INIRead(lpModelName, "Act01TimeHi" & Format(i, "00"), lpFileName)
    Next
    
    SetupVar.dAct01StallDeltaVoltLo = INIRead(lpModelName, "Act01StallDeltaVoltLo", lpFileName)
    SetupVar.dAct01StallDeltaVoltHi = INIRead(lpModelName, "Act01StallDeltaVoltHi", lpFileName)
    
    ' Act 02
    SetupVar.bAct02Use = INIRead(lpModelName, "Act02Use", lpFileName)
    SetupVar.nAct02Direction = INIRead(lpModelName, "Act02Direction", lpFileName)
    SetupVar.nAct02TestType = INIRead(lpModelName, "Act02TestType", lpFileName)
    
    For i = 0 To frmSetup.txtAct02Name.UBound
        SetupVar.lpAct02Name(i) = INIRead(lpModelName, "Act02Name" & Format(i, "00"), lpFileName)
        SetupVar.dAct02SetVolt(i) = INIRead(lpModelName, "Act02SetVolt" & Format(i, "00"), lpFileName)
        SetupVar.dAct02CurrLo(i) = INIRead(lpModelName, "Act02CurrLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct02CurrHi(i) = INIRead(lpModelName, "Act02CurrHi" & Format(i, "00"), lpFileName)
        SetupVar.dAct02VoltLo(i) = INIRead(lpModelName, "Act02VoltLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct02VoltHi(i) = INIRead(lpModelName, "Act02VoltHi" & Format(i, "00"), lpFileName)
        SetupVar.dAct02TimeLo(i) = INIRead(lpModelName, "Act02TimeLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct02TimeHi(i) = INIRead(lpModelName, "Act02TimeHi" & Format(i, "00"), lpFileName)
    Next
    
    SetupVar.dAct02StallDeltaVoltLo = INIRead(lpModelName, "Act02StallDeltaVoltLo", lpFileName)
    SetupVar.dAct02StallDeltaVoltHi = INIRead(lpModelName, "Act02StallDeltaVoltHi", lpFileName)
    
    ' Act 03
    SetupVar.bAct03Use = INIRead(lpModelName, "Act03Use", lpFileName)
    SetupVar.nAct03Direction = INIRead(lpModelName, "Act03Direction", lpFileName)
    SetupVar.nAct03TestType = INIRead(lpModelName, "Act03TestType", lpFileName)
    
    For i = 0 To frmSetup.txtAct03Name.UBound
        SetupVar.lpAct03Name(i) = INIRead(lpModelName, "Act03Name" & Format(i, "00"), lpFileName)
        SetupVar.dAct03SetVolt(i) = INIRead(lpModelName, "Act03SetVolt" & Format(i, "00"), lpFileName)
        SetupVar.dAct03CurrLo(i) = INIRead(lpModelName, "Act03CurrLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct03CurrHi(i) = INIRead(lpModelName, "Act03CurrHi" & Format(i, "00"), lpFileName)
        SetupVar.dAct03VoltLo(i) = INIRead(lpModelName, "Act03VoltLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct03VoltHi(i) = INIRead(lpModelName, "Act03VoltHi" & Format(i, "00"), lpFileName)
        SetupVar.dAct03TimeLo(i) = INIRead(lpModelName, "Act03TimeLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct03TimeHi(i) = INIRead(lpModelName, "Act03TimeHi" & Format(i, "00"), lpFileName)
    Next
    
    SetupVar.dAct03StallDeltaVoltLo = INIRead(lpModelName, "Act03StallDeltaVoltLo", lpFileName)
    SetupVar.dAct03StallDeltaVoltHi = INIRead(lpModelName, "Act03StallDeltaVoltHi", lpFileName)
    
    ' Act 04
    SetupVar.bAct04Use = INIRead(lpModelName, "Act04Use", lpFileName)
    SetupVar.nAct042Pin = INIRead(lpModelName, "Act042Pin", lpFileName)
    SetupVar.nAct042PinPos = INIRead(lpModelName, "Act042PinPos", lpFileName)
    SetupVar.nAct04Direction = INIRead(lpModelName, "Act04Direction", lpFileName)
    SetupVar.nAct04TestType = INIRead(lpModelName, "Act04TestType", lpFileName)
    
    For i = 0 To frmSetup.txtAct04Name.UBound
        SetupVar.lpAct04Name(i) = INIRead(lpModelName, "Act04Name" & Format(i, "00"), lpFileName)
        SetupVar.dAct04SetVolt(i) = INIRead(lpModelName, "Act04SetVolt" & Format(i, "00"), lpFileName)
        SetupVar.dAct04CurrLo(i) = INIRead(lpModelName, "Act04CurrLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct04CurrHi(i) = INIRead(lpModelName, "Act04CurrHi" & Format(i, "00"), lpFileName)
        SetupVar.dAct04VoltLo(i) = INIRead(lpModelName, "Act04voltLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct04VoltHi(i) = INIRead(lpModelName, "Act04voltHi" & Format(i, "00"), lpFileName)
        SetupVar.dAct04TimeLo(i) = INIRead(lpModelName, "Act04timeLo" & Format(i, "00"), lpFileName)
        SetupVar.dAct04TimeHi(i) = INIRead(lpModelName, "Act04timeHi" & Format(i, "00"), lpFileName)
    Next
    
    SetupVar.dAct04StallDeltaVoltLo = INIRead(lpModelName, "act04StallDeltaVoltLo", lpFileName)
    SetupVar.dAct04StallDeltaVoltHi = INIRead(lpModelName, "act04StallDeltaVoltHi", lpFileName)
    
    ' Ion
    SetupVar.lpIonName = INIRead(lpModelName, "IonName", lpFileName)
    SetupVar.bIonUse = IIf(INIRead(lpModelName, "IonUse", lpFileName) = "True", True, False)
    
    For i = 0 To UBound(SetupVar.lpIonSubName)
        SetupVar.lpIonSubName(i) = INIRead(lpModelName, "IonSubName" & Format(i, "00"), lpFileName)
        SetupVar.dIonHi(i) = INIRead(lpModelName, "IonHi" & Format(i, "00"), lpFileName)
        SetupVar.dIonLo(i) = INIRead(lpModelName, "IonLo" & Format(i, "00"), lpFileName)
    Next
    
    ' Sensor
    For i = 0 To frmSetup.txtSensorName.UBound
        SetupVar.bSensorUse(i) = IIf(INIRead(lpModelName, "SensorUse" & Format(i, "00"), lpFileName) = "True", True, False)
        SetupVar.lpSensorName(i) = INIRead(lpModelName, "SensorName" & Format(i, "00"), lpFileName)
        SetupVar.dSensorCurrLo(i) = INIRead(lpModelName, "SensorCurrLo" & Format(i, "00"), lpFileName)
        SetupVar.dSensorCurrHi(i) = INIRead(lpModelName, "SensorCurrHi" & Format(i, "00"), lpFileName)
        SetupVar.dSensorTime(i) = INIRead(lpModelName, "SensorTime" & Format(i, "00"), lpFileName)
    Next
    
    ' PTC
    SetupVar.bPTCUse = IIf(INIRead(lpModelName, "PTCUse", lpFileName) = "True", True, False)
    SetupVar.lpPTCName = INIRead(lpModelName, "PTCName", lpFileName)
    SetupVar.dPTCCurrLo = INIRead(lpModelName, "PTCCurrLo", lpFileName)
    SetupVar.dPTCCurrHi = INIRead(lpModelName, "PTCCurrHi", lpFileName)
    SetupVar.dPTCTime = INIRead(lpModelName, "PTCTime", lpFileName)
    
    ' Leak
    For i = 0 To frmSetup.chkLeak.UBound
        SetupVar.bLeakUse(i) = IIf(INIRead(lpModelName, "LeakUse" & Format(i, "00"), lpFileName) = "True", True, False)
        SetupVar.lpLeakName(i) = INIRead(lpModelName, "LeakName" & Format(i, "00"), lpFileName)
    Next
    
    SetupVar.nLeakModel = INIRead(lpModelName, "LeakModel", lpFileName)
    
    ' Vision
    SetupVar.bVisionUse = IIf(INIRead(lpModelName, "VisionUse", lpFileName) = "True", True, False)
    SetupVar.lpVisionName(0) = INIRead(lpModelName, "VisionName00", lpFileName)
    
    For i = 0 To 3
        SetupVar.bVisionDoorUse(i) = IIf(INIRead(lpModelName, "VisionDoorUse" & Format(i, "00"), lpFileName) = "True", True, False)
    Next
    
    For i = 0 To 3
        SetupVar.lpVisionDoorName(i) = INIRead(lpModelName, "VisionDoorName" & Format(i, "00"), lpFileName)
    Next
    
    For i = 0 To 3
        SetupVar.nOpenCameraNo(i) = INIRead(lpModelName, "OpenCameraNo" & i, lpFileName)
        SetupVar.nOpenCameraPos(i) = INIRead(lpModelName, "OpenCameraPos" & i, lpFileName)
        SetupVar.nCloseCameraNo(i) = INIRead(lpModelName, "CloseCameraNo" & i, lpFileName)
        SetupVar.nCloseCameraPos(i) = INIRead(lpModelName, "CloseCameraPos" & i, lpFileName)
    Next
    
    ' Barcode
    SetupVar.bBarCodeUse = IIf(INIRead(lpModelName, "BarcodeUse", lpFileName) = "True", True, False)
    SetupVar.nBarcodeType = INIRead(lpModelName, "BarcodeType", lpFileName)
    
    For i = 0 To 9
        SetupVar.lpBarcode(i) = INIRead(lpModelName, "BarcodeStr" & Format(i, "00"), lpFileName)
    Next
    
    ' Marking
    For i = 0 To 1
        SetupVar.bMarkingUse(i) = IIf(INIRead(lpModelName, "MarkingUse" & Format(i, "00"), lpFileName) = "True", True, False)
        SetupVar.dMarkingTime(i) = INIRead(lpModelName, "MarkingTime" & Format(i, "00"), lpFileName)
    Next
    
    ' Part
    For i = 0 To MAX_DIO_CHANNEL
        SetupVar.bPartUse(i) = IIf(INIRead(lpModelName, "PartUse" & Format(i, "000"), lpFileName) = "True", True, False)
        SetupVar.bPartStatus(i) = IIf(INIRead(lpModelName, "PartStatus" & Format(i, "000"), lpFileName) = "True", True, False)
        SetupVar.lpPartName(i) = INIRead(lpModelName, "PartName" & Format(i, "000"), lpFileName)
    Next
    
    SetupVar.bProductUse = IIf(INIRead(lpModelName, "ProductUse", lpFileName) = "True", True, False)
    SetupVar.lpProductList = INIRead(lpModelName, "ProductList", lpFileName)
    SetupVar.lpProductName = INIRead(lpModelName, "ProductName", lpFileName)
    SetupVar.bModelTypeUse = IIf(INIRead(lpModelName, "ModelTypeUse", lpFileName) = "True", True, False)
    SetupVar.lpModelLHDList = INIRead(lpModelName, "ModelLHDList", lpFileName)
    SetupVar.lpLHDPartName = INIRead(lpModelName, "LHDPartName", lpFileName)
    SetupVar.lpModelRHDList = INIRead(lpModelName, "ModelRHDList", lpFileName)
    SetupVar.lpRHDPartName = INIRead(lpModelName, "RHDPartName", lpFileName)
    
    ' Lin
    For i = 0 To 3
        SetupVar.bLinActUse(i) = IIf(INIRead(lpModelName, "LinActUse" & Format(i, "00"), lpFileName) = "True", True, False)
        SetupVar.lpLinActName(i) = INIRead(lpModelName, "LinActName" & Format(i, "00"), lpFileName)
        SetupVar.dLinActLo(i) = INIRead(lpModelName, "LinActLo" & Format(i, "00"), lpFileName)
        SetupVar.dLinActHi(i) = INIRead(lpModelName, "LinActHi" & Format(i, "00"), lpFileName)
        SetupVar.nLinActFinal(i) = INIRead(lpModelName, "LinActRange" & Format(i, "00"), lpFileName)
        SetupVar.nLinActMove(i) = INIRead(lpModelName, "LinActMove" & Format(i, "00"), lpFileName)
        SetupVar.lLinActAngle(i) = INIRead(lpModelName, "LinActAngle" & Format(i, "00"), lpFileName)
        SetupVar.dLinActTime(i) = INIRead(lpModelName, "LinActTime" & Format(i, "00"), lpFileName)
        SetupVar.dLinActCurrLo(i) = INIRead(lpModelName, "LinActCurrLo" & Format(i, "00"), lpFileName)
        SetupVar.dLinActCurrHi(i) = INIRead(lpModelName, "LinActCurrHi" & Format(i, "00"), lpFileName)
    Next
    
    For i = 0 To 4
        SetupVar.nLinAct01Check(i) = INIRead(lpModelName, "LinAct01Check" & Format(i, "00"), lpFileName)
        SetupVar.nLinAct02Check(i) = INIRead(lpModelName, "LinAct02Check" & Format(i, "00"), lpFileName)
        SetupVar.nLinAct03Check(i) = INIRead(lpModelName, "LinAct03Check" & Format(i, "00"), lpFileName)
        SetupVar.nLinAct04Check(i) = INIRead(lpModelName, "LinAct04Check" & Format(i, "00"), lpFileName)
        SetupVar.dLinAct01CheckTime(i) = INIRead(lpModelName, "LinAct01CheckTime" & Format(i, "00"), lpFileName)
        SetupVar.dLinAct02CheckTime(i) = INIRead(lpModelName, "LinAct02CheckTime" & Format(i, "00"), lpFileName)
        SetupVar.dLinAct03CheckTime(i) = INIRead(lpModelName, "LinAct03CheckTime" & Format(i, "00"), lpFileName)
        SetupVar.dLinAct04CheckTime(i) = INIRead(lpModelName, "LinAct04CheckTime" & Format(i, "00"), lpFileName)
    Next
    
    SetupVar.nLinActFirstMove(0) = INIRead(lpModelName, "LinAct01FirstMove", lpFileName)
    SetupVar.nLinActFirstMove(1) = INIRead(lpModelName, "LinAct02FirstMove", lpFileName)
    SetupVar.nLinActFirstMove(2) = INIRead(lpModelName, "LinAct03FirstMove", lpFileName)
    SetupVar.nLinActFirstMove(3) = INIRead(lpModelName, "LinAct04FirstMove", lpFileName)
    
    SetupVar.bAutoAddressUse = IIf(INIRead(lpModelName, "AutoAddressUse", lpFileName) = "True", True, False)
    SetupVar.nLinTestType = INIRead(lpModelName, "LinCurrType", lpFileName)
    
    SetupVar.nLinAct01RefPos = INIRead(lpModelName, "Act01RefPos", lpFileName)
    
    SetupVar.bCheckPointUse = IIf(INIRead(lpModelName, "CheckPointUse", lpFileName) = "True", True, False)
    SetupVar.bStallUse = IIf(INIRead(lpModelName, "StallUse", lpFileName) = "True", True, False)
    SetupVar.bLinBlowerUse = IIf(INIRead(lpModelName, "LinBlowerUse", lpFileName) = "True", True, False)
    SetupVar.lpLinBlowerName = INIRead(lpModelName, "LinBlowerName", lpFileName)
    SetupVar.dLinBlowerTime = INIRead(lpModelName, "LinBlowerTime", lpFileName)
    
    ' Offset
    For i = 0 To MAX_AD_CHANNEL
        SetupVar.bAdjustUse(i) = IIf(INIRead(lpModelName, "AdjustUse" & Format(i, "00"), lpFileName) = "True", True, False)
        SetupVar.dAdd(i) = INIRead(lpModelName, "Add" & Format(i, "00"), lpFileName)
        SetupVar.dMulti(i) = INIRead(lpModelName, "Multi" & Format(i, "00"), lpFileName)
    Next
    
    ' Curr Count
    For i = 0 To 3
        SetupVar.nActPeakCurrCount(i) = INIRead(lpModelName, "ActPeakCurrCount" & Format(i, "00"), lpFileName)
        SetupVar.nActEndPosCount(i) = INIRead(lpModelName, "ActEndPosCount" & Format(i, "00"), lpFileName)
    Next
    
    ' Stepping Set
    For i = 0 To frmSetup.txtStepAct01Arr.UBound
        SetupVar.nSteppingSet(0, i) = INIRead(lpModelName, "nSteppingSet01" & Format(i, "00"), lpFileName)
        SetupVar.nSteppingSet(1, i) = INIRead(lpModelName, "nSteppingSet02" & Format(i, "00"), lpFileName)
        SetupVar.nSteppingSet(2, i) = INIRead(lpModelName, "nSteppingSet03" & Format(i, "00"), lpFileName)
        SetupVar.nSteppingSet(3, i) = INIRead(lpModelName, "nSteppingSet04" & Format(i, "00"), lpFileName)
    Next
    
    ' label name
    For i = 0 To MAX_AD_CHANNEL
        frmSetup.pnlAdChName(i).Caption = SysVar.lpName(i)
    Next
End Function

Public Function SaveSetupFile(ByVal lpModelName As String) As Boolean
    Dim i               As Integer
    Dim lpFileName      As String
    
    On Error Resume Next
    
    lpFileName = lpPath & INISETUPFILE
    SaveSetupFile = SearchFile(lpFileName)
    
    ' Hidden
    Call INIWrite(lpModelName, "LinActUse", CStr(SetupVar.dTestVolt), lpFileName)
    
    ' General
    Call INIWrite(lpModelName, "TestVolt", CStr(SetupVar.dTestVolt), lpFileName)
    Call INIWrite(lpModelName, "FileName", CStr(SetupVar.lpFileName), lpFileName)
    Call INIWrite(lpModelName, "NoDataSave", CStr(SetupVar.bDataSave), lpFileName)
    Call INIWrite(lpModelName, "ScannerUse", CStr(SetupVar.bScannerUse), lpFileName)
    Call INIWrite(lpModelName, "ScannerValue", CStr(SetupVar.nScannerValue), lpFileName)
    Call INIWrite(lpModelName, "NvhUse", CStr(SetupVar.bNvhUse), lpFileName)
    Call INIWrite(lpModelName, "ModelType", CStr(SetupVar.nModelType), lpFileName)
    
    For i = 0 To frmSetup.txtActName.UBound
        Call INIWrite(lpModelName, "ActName" & Format(i, "00"), CStr(SetupVar.lpActName(i)), lpFileName)
        Call INIWrite(lpModelName, "ActBoardNo" & Format(i, "00"), CStr(SetupVar.nActBoardNo(i)), lpFileName)
    Next
    
    ' Blower
    Call INIWrite(lpModelName, "BlowerUse", CStr(SetupVar.bBlowerUse), lpFileName)
    Call INIWrite(lpModelName, "BlowerType", CStr(SetupVar.nBlowerType), lpFileName)
    Call INIWrite(lpModelName, "BlowerDirection", CStr(SetupVar.nBlowerDirection), lpFileName)
    
    For i = 0 To frmSetup.txtBlowerName.UBound
        Call INIWrite(lpModelName, "BlowerName" & Format(i, "00"), CStr(SetupVar.lpBlowerName(i)), lpFileName)
        Call INIWrite(lpModelName, "BlowerCurrLo" & Format(i, "00"), CStr(SetupVar.dBlowerCurrLo(i)), lpFileName)
        Call INIWrite(lpModelName, "BlowerCurrHi" & Format(i, "00"), CStr(SetupVar.dBlowerCurrHi(i)), lpFileName)
        Call INIWrite(lpModelName, "BlowerTime" & Format(i, "00"), CStr(SetupVar.dBlowerTime(i)), lpFileName)
        Call INIWrite(lpModelName, "LinSpeed" & Format(i, "00"), CStr(SetupVar.nLinSpeed(i)), lpFileName)
    Next
    
    ' RPM
    Call INIWrite(lpModelName, "RpmName", CStr(SetupVar.lpRpmName), lpFileName)
    Call INIWrite(lpModelName, "RpmCurrLo", CStr(SetupVar.dRpmCurrLo), lpFileName)
    Call INIWrite(lpModelName, "RpmCurrHi", CStr(SetupVar.dRpmCurrHi), lpFileName)
    
    ' Vibration
    Call INIWrite(lpModelName, "VibUse", CStr(SetupVar.bVibUse), lpFileName)
    Call INIWrite(lpModelName, "VibName", CStr(SetupVar.lpVibName), lpFileName)
    Call INIWrite(lpModelName, "VibCurrLo", CStr(SetupVar.dVibCurrLo), lpFileName)
    Call INIWrite(lpModelName, "VibCurrHi", CStr(SetupVar.dVibCurrHi), lpFileName)
    Call INIWrite(lpModelName, "VibResultType", CStr(SetupVar.nVibResultType), lpFileName)
    Call INIWrite(lpModelName, "VibStart", CStr(SetupVar.dVibStart), lpFileName)
    Call INIWrite(lpModelName, "VibEnd", CStr(SetupVar.dVibEnd), lpFileName)
    Call INIWrite(lpModelName, "VibMethod", CStr(SetupVar.nVibMethod), lpFileName)
    Call INIWrite(lpModelName, "VibVolt", CStr(SetupVar.dVibVolt), lpFileName)
    Call INIWrite(lpModelName, "VibTime", CStr(SetupVar.dVibTime), lpFileName)
    
    ' Act 01
    Call INIWrite(lpModelName, "Act01Use", CStr(SetupVar.bAct01Use), lpFileName)
    Call INIWrite(lpModelName, "Act01Direction", CStr(SetupVar.nAct01Direction), lpFileName)
    Call INIWrite(lpModelName, "Act01TestType", CStr(SetupVar.nAct01TestType), lpFileName)
    
    For i = 0 To frmSetup.txtAct01Name.UBound
        Call INIWrite(lpModelName, "Act01Name" & Format(i, "00"), CStr(SetupVar.lpAct01Name(i)), lpFileName)
        Call INIWrite(lpModelName, "Act01SetVolt" & Format(i, "00"), CStr(SetupVar.dAct01SetVolt(i)), lpFileName)
        Call INIWrite(lpModelName, "Act01CurrLo" & Format(i, "00"), CStr(SetupVar.dAct01CurrLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act01CurrHi" & Format(i, "00"), CStr(SetupVar.dAct01CurrHi(i)), lpFileName)
        Call INIWrite(lpModelName, "Act01VoltLo" & Format(i, "00"), CStr(SetupVar.dAct01VoltLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act01VoltHi" & Format(i, "00"), CStr(SetupVar.dAct01VoltHi(i)), lpFileName)
        Call INIWrite(lpModelName, "Act01TimeLo" & Format(i, "00"), CStr(SetupVar.dAct01TimeLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act01TimeHi" & Format(i, "00"), CStr(SetupVar.dAct01TimeHi(i)), lpFileName)
    Next
    
    Call INIWrite(lpModelName, "Act01StallDeltaVoltLo", CStr(SetupVar.dAct01StallDeltaVoltLo), lpFileName)
    Call INIWrite(lpModelName, "Act01StallDeltaVoltHi", CStr(SetupVar.dAct01StallDeltaVoltHi), lpFileName)
    
    ' Act 02
    Call INIWrite(lpModelName, "Act02Use", CStr(SetupVar.bAct02Use), lpFileName)
    Call INIWrite(lpModelName, "Act02Direction", CStr(SetupVar.nAct02Direction), lpFileName)
    Call INIWrite(lpModelName, "Act02TestType", CStr(SetupVar.nAct02TestType), lpFileName)
    
    For i = 0 To frmSetup.txtAct02Name.UBound
        Call INIWrite(lpModelName, "Act02Name" & Format(i, "00"), CStr(SetupVar.lpAct02Name(i)), lpFileName)
        Call INIWrite(lpModelName, "Act02SetVolt" & Format(i, "00"), CStr(SetupVar.dAct02SetVolt(i)), lpFileName)
        Call INIWrite(lpModelName, "Act02CurrLo" & Format(i, "00"), CStr(SetupVar.dAct02CurrLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act02CurrHi" & Format(i, "00"), CStr(SetupVar.dAct02CurrHi(i)), lpFileName)
        Call INIWrite(lpModelName, "Act02VoltLo" & Format(i, "00"), CStr(SetupVar.dAct02VoltLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act02VoltHi" & Format(i, "00"), CStr(SetupVar.dAct02VoltHi(i)), lpFileName)
        Call INIWrite(lpModelName, "Act02TimeLo" & Format(i, "00"), CStr(SetupVar.dAct02TimeLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act02TimeHi" & Format(i, "00"), CStr(SetupVar.dAct02TimeHi(i)), lpFileName)
    Next
    
    Call INIWrite(lpModelName, "Act02StallDeltaVoltLo", CStr(SetupVar.dAct02StallDeltaVoltLo), lpFileName)
    Call INIWrite(lpModelName, "Act02StallDeltaVoltHi", CStr(SetupVar.dAct02StallDeltaVoltHi), lpFileName)
    
    ' Act 03
    Call INIWrite(lpModelName, "Act03Use", CStr(SetupVar.bAct03Use), lpFileName)
    Call INIWrite(lpModelName, "Act03Direction", CStr(SetupVar.nAct03Direction), lpFileName)
    Call INIWrite(lpModelName, "Act03TestType", CStr(SetupVar.nAct03TestType), lpFileName)
    
    For i = 0 To frmSetup.txtAct03Name.UBound
        Call INIWrite(lpModelName, "Act03Name" & Format(i, "00"), CStr(SetupVar.lpAct03Name(i)), lpFileName)
        Call INIWrite(lpModelName, "Act03SetVolt" & Format(i, "00"), CStr(SetupVar.dAct03SetVolt(i)), lpFileName)
        Call INIWrite(lpModelName, "Act03CurrLo" & Format(i, "00"), CStr(SetupVar.dAct03CurrLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act03CurrHi" & Format(i, "00"), CStr(SetupVar.dAct03CurrHi(i)), lpFileName)
        Call INIWrite(lpModelName, "Act03VoltLo" & Format(i, "00"), CStr(SetupVar.dAct03VoltLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act03VoltHi" & Format(i, "00"), CStr(SetupVar.dAct03VoltHi(i)), lpFileName)
        Call INIWrite(lpModelName, "Act03TimeLo" & Format(i, "00"), CStr(SetupVar.dAct03TimeLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act03TimeHi" & Format(i, "00"), CStr(SetupVar.dAct03TimeHi(i)), lpFileName)
    Next
    
    Call INIWrite(lpModelName, "Act03StallDeltaVoltLo", CStr(SetupVar.dAct03StallDeltaVoltLo), lpFileName)
    Call INIWrite(lpModelName, "Act03StallDeltaVoltHi", CStr(SetupVar.dAct03StallDeltaVoltHi), lpFileName)
    
    ' Act 04
    Call INIWrite(lpModelName, "Act04Use", CStr(SetupVar.bAct04Use), lpFileName)
    Call INIWrite(lpModelName, "Act042Pin", CStr(SetupVar.nAct042Pin), lpFileName)
    Call INIWrite(lpModelName, "Act042PinPos", CStr(SetupVar.nAct042PinPos), lpFileName)
    Call INIWrite(lpModelName, "Act04Direction", CStr(SetupVar.nAct04Direction), lpFileName)
    Call INIWrite(lpModelName, "Act04TestType", CStr(SetupVar.nAct04TestType), lpFileName)
    
    For i = 0 To frmSetup.txtAct04Name.UBound
        Call INIWrite(lpModelName, "Act04Name" & Format(i, "00"), CStr(SetupVar.lpAct04Name(i)), lpFileName)
        Call INIWrite(lpModelName, "Act04SetVolt" & Format(i, "00"), CStr(SetupVar.dAct04SetVolt(i)), lpFileName)
        Call INIWrite(lpModelName, "Act04CurrLo" & Format(i, "00"), CStr(SetupVar.dAct04CurrLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act04CurrHi" & Format(i, "00"), CStr(SetupVar.dAct04CurrHi(i)), lpFileName)
        Call INIWrite(lpModelName, "Act04VoltLo" & Format(i, "00"), CStr(SetupVar.dAct04VoltLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act04VoltHi" & Format(i, "00"), CStr(SetupVar.dAct04VoltHi(i)), lpFileName)
        Call INIWrite(lpModelName, "Act04TimeLo" & Format(i, "00"), CStr(SetupVar.dAct04TimeLo(i)), lpFileName)
        Call INIWrite(lpModelName, "Act04TimeHi" & Format(i, "00"), CStr(SetupVar.dAct04TimeHi(i)), lpFileName)
    Next
    
    Call INIWrite(lpModelName, "Act04StallDeltaVoltLo", CStr(SetupVar.dAct04StallDeltaVoltLo), lpFileName)
    Call INIWrite(lpModelName, "Act04StallDeltaVoltHi", CStr(SetupVar.dAct04StallDeltaVoltHi), lpFileName)
    
    ' Ion
    Call INIWrite(lpModelName, "IonName", CStr(SetupVar.lpIonName), lpFileName)
    Call INIWrite(lpModelName, "IonUse", CStr(SetupVar.bIonUse), lpFileName)
    
    For i = 0 To 1
        Call INIWrite(lpModelName, "IonSubName" & Format(i, "00"), CStr(SetupVar.lpIonSubName(i)), lpFileName)
        Call INIWrite(lpModelName, "IonHi" & Format(i, "00"), CStr(SetupVar.dIonHi(i)), lpFileName)
        Call INIWrite(lpModelName, "IonLo" & Format(i, "00"), CStr(SetupVar.dIonLo(i)), lpFileName)
    Next
    
    ' Sensor
    For i = 0 To frmSetup.txtSensorName.UBound
        Call INIWrite(lpModelName, "SensorUse" & Format(i, "00"), CStr(SetupVar.bSensorUse(i)), lpFileName)
        Call INIWrite(lpModelName, "SensorName" & Format(i, "00"), CStr(SetupVar.lpSensorName(i)), lpFileName)
        Call INIWrite(lpModelName, "SensorCurrLo" & Format(i, "00"), CStr(SetupVar.dSensorCurrLo(i)), lpFileName)
        Call INIWrite(lpModelName, "SensorCurrHi" & Format(i, "00"), CStr(SetupVar.dSensorCurrHi(i)), lpFileName)
        Call INIWrite(lpModelName, "SensorTime" & Format(i, "00"), CStr(SetupVar.dSensorTime(i)), lpFileName)
    Next
    
    ' PTC
    Call INIWrite(lpModelName, "PTCUse", CStr(SetupVar.bPTCUse), lpFileName)
    Call INIWrite(lpModelName, "PTCName", CStr(SetupVar.lpPTCName), lpFileName)
    Call INIWrite(lpModelName, "PTCCurrLo", CStr(SetupVar.dPTCCurrLo), lpFileName)
    Call INIWrite(lpModelName, "PTCCurrHi", CStr(SetupVar.dPTCCurrHi), lpFileName)
    Call INIWrite(lpModelName, "PTCTime", CStr(SetupVar.dPTCTime), lpFileName)
    
    ' Leak
    For i = 0 To frmSetup.chkLeak.UBound
        Call INIWrite(lpModelName, "LeakUse" & Format(i, "00"), CStr(SetupVar.bLeakUse(i)), lpFileName)
        Call INIWrite(lpModelName, "LeakName" & Format(i, "00"), CStr(SetupVar.lpLeakName(i)), lpFileName)
    Next
    
    Call INIWrite(lpModelName, "LeakModel", CStr(SetupVar.nLeakModel), lpFileName)
    
    ' Vision
    Call INIWrite(lpModelName, "VisionUse", CStr(SetupVar.bVisionUse), lpFileName)
    Call INIWrite(lpModelName, "VisionName00", CStr(SetupVar.lpVisionName(0)), lpFileName)
    
    For i = 0 To 3
        Call INIWrite(lpModelName, "VisionDoorUse" & Format(i, "00"), CStr(SetupVar.bVisionDoorUse(i)), lpFileName)
    Next
    
    For i = 0 To 3
        Call INIWrite(lpModelName, "VisionDoorName" & Format(i, "00"), CStr(SetupVar.lpVisionDoorName(i)), lpFileName)
    Next
    
    For i = 0 To 3
        Call INIWrite(lpModelName, "OpenCameraNo" & i, CStr(SetupVar.nOpenCameraNo(i)), lpFileName)
        Call INIWrite(lpModelName, "OpenCameraPos" & i, CStr(SetupVar.nOpenCameraPos(i)), lpFileName)
        Call INIWrite(lpModelName, "CloseCameraNo" & i, CStr(SetupVar.nCloseCameraNo(i)), lpFileName)
        Call INIWrite(lpModelName, "CloseCameraPos" & i, CStr(SetupVar.nCloseCameraPos(i)), lpFileName)
    Next
    
    ' Barcode
    Call INIWrite(lpModelName, "BarcodeUse", CStr(SetupVar.bBarCodeUse), lpFileName)
    Call INIWrite(lpModelName, "BarcodeType", CStr(SetupVar.nBarcodeType), lpFileName)
    
    For i = 0 To 9
        Call INIWrite(lpModelName, "BarcodeStr" & Format(i, "00"), CStr(SetupVar.lpBarcode(i)), lpFileName)
    Next
    
    ' Marking
    For i = 0 To 1
        Call INIWrite(lpModelName, "MarkingUse" & Format(i, "00"), CStr(SetupVar.bMarkingUse(i)), lpFileName)
        Call INIWrite(lpModelName, "MarkingTime" & Format(i, "00"), CStr(SetupVar.dMarkingTime(i)), lpFileName)
    Next
    
    ' Part
    For i = 0 To MAX_DIO_CHANNEL
        Call INIWrite(lpModelName, "PartUse" & Format(i, "000"), CStr(SetupVar.bPartUse(i)), lpFileName)
        Call INIWrite(lpModelName, "PartStatus" & Format(i, "000"), CStr(SetupVar.bPartStatus(i)), lpFileName)
        Call INIWrite(lpModelName, "PartName" & Format(i, "000"), CStr(SetupVar.lpPartName(i)), lpFileName)
    Next
    
    Call INIWrite(lpModelName, "ProductUse", CStr(SetupVar.bProductUse), lpFileName)
    Call INIWrite(lpModelName, "ProductList", CStr(SetupVar.lpProductList), lpFileName)
    Call INIWrite(lpModelName, "ProductName", CStr(SetupVar.lpProductName), lpFileName)
    Call INIWrite(lpModelName, "ModelTypeUse", CStr(SetupVar.bModelTypeUse), lpFileName)
    Call INIWrite(lpModelName, "ModelLHDList", CStr(SetupVar.lpModelLHDList), lpFileName)
    Call INIWrite(lpModelName, "LHDPartName", CStr(SetupVar.lpLHDPartName), lpFileName)
    Call INIWrite(lpModelName, "ModelRHDList", CStr(SetupVar.lpModelRHDList), lpFileName)
    Call INIWrite(lpModelName, "RHDPartName", CStr(SetupVar.lpRHDPartName), lpFileName)
    
    ' Lin
    For i = 0 To 3
        Call INIWrite(lpModelName, "LinActUse" & Format(i, "00"), CStr(SetupVar.bLinActUse(i)), lpFileName)
        Call INIWrite(lpModelName, "LinActName" & Format(i, "00"), CStr(SetupVar.lpLinActName(i)), lpFileName)
        Call INIWrite(lpModelName, "LinActLo" & Format(i, "00"), CStr(SetupVar.dLinActLo(i)), lpFileName)
        Call INIWrite(lpModelName, "LinActHi" & Format(i, "00"), CStr(SetupVar.dLinActHi(i)), lpFileName)
        Call INIWrite(lpModelName, "LinActRange" & Format(i, "00"), CStr(SetupVar.nLinActFinal(i)), lpFileName)
        Call INIWrite(lpModelName, "LinActMove" & Format(i, "00"), CStr(SetupVar.nLinActMove(i)), lpFileName)
        Call INIWrite(lpModelName, "LinActAngle" & Format(i, "00"), CStr(SetupVar.lLinActAngle(i)), lpFileName)
        Call INIWrite(lpModelName, "LinActTime" & Format(i, "00"), CStr(SetupVar.dLinActTime(i)), lpFileName)
        Call INIWrite(lpModelName, "LinActCurrLo" & Format(i, "00"), CStr(SetupVar.dLinActCurrLo(i)), lpFileName)
        Call INIWrite(lpModelName, "LinActCurrHi" & Format(i, "00"), CStr(SetupVar.dLinActCurrHi(i)), lpFileName)
    Next
    
    For i = 0 To 4
        Call INIWrite(lpModelName, "LinAct01Check" & Format(i, "00"), CStr(SetupVar.nLinAct01Check(i)), lpFileName)
        Call INIWrite(lpModelName, "LinAct02Check" & Format(i, "00"), CStr(SetupVar.nLinAct02Check(i)), lpFileName)
        Call INIWrite(lpModelName, "LinAct03Check" & Format(i, "00"), CStr(SetupVar.nLinAct03Check(i)), lpFileName)
        Call INIWrite(lpModelName, "LinAct04Check" & Format(i, "00"), CStr(SetupVar.nLinAct04Check(i)), lpFileName)
        Call INIWrite(lpModelName, "LinAct01CheckTime" & Format(i, "00"), CStr(SetupVar.dLinAct01CheckTime(i)), lpFileName)
        Call INIWrite(lpModelName, "LinAct02CheckTime" & Format(i, "00"), CStr(SetupVar.dLinAct02CheckTime(i)), lpFileName)
        Call INIWrite(lpModelName, "LinAct03CheckTime" & Format(i, "00"), CStr(SetupVar.dLinAct03CheckTime(i)), lpFileName)
        Call INIWrite(lpModelName, "LinAct04CheckTime" & Format(i, "00"), CStr(SetupVar.dLinAct04CheckTime(i)), lpFileName)
    Next
    
    Call INIWrite(lpModelName, "LinAct01FirstMove", CStr(SetupVar.nLinActFirstMove(0)), lpFileName)
    Call INIWrite(lpModelName, "LinAct02FirstMove", CStr(SetupVar.nLinActFirstMove(1)), lpFileName)
    Call INIWrite(lpModelName, "LinAct03FirstMove", CStr(SetupVar.nLinActFirstMove(2)), lpFileName)
    Call INIWrite(lpModelName, "LinAct04FirstMove", CStr(SetupVar.nLinActFirstMove(3)), lpFileName)
    
    Call INIWrite(lpModelName, "AutoAddressUse", CStr(SetupVar.bAutoAddressUse), lpFileName)
    Call INIWrite(lpModelName, "LinCurrType", CStr(SetupVar.nLinTestType), lpFileName)
    
    Call INIWrite(lpModelName, "Act01RefPos", CStr(SetupVar.nLinAct01RefPos), lpFileName)
    
    Call INIWrite(lpModelName, "CheckPointUse", CStr(SetupVar.bCheckPointUse), lpFileName)
    Call INIWrite(lpModelName, "StallUse", CStr(SetupVar.bStallUse), lpFileName)
    Call INIWrite(lpModelName, "LinBlowerUse", CStr(SetupVar.bLinBlowerUse), lpFileName)
    Call INIWrite(lpModelName, "LinBlowerName", CStr(SetupVar.lpLinBlowerName), lpFileName)
    Call INIWrite(lpModelName, "LinBlowerTime", CStr(SetupVar.dLinBlowerTime), lpFileName)
    
    ' Curr Count
    For i = 0 To 3
        Call INIWrite(lpModelName, "ActPeakCurrCount" & Format(i, "00"), CStr(SetupVar.nActPeakCurrCount(i)), lpFileName)
        Call INIWrite(lpModelName, "ActEndPosCount" & Format(i, "00"), CStr(SetupVar.nActEndPosCount(i)), lpFileName)
    Next
    
    ' Stepping Set
    For i = 0 To frmSetup.txtStepAct01Arr.UBound
        Call INIWrite(lpModelName, "nSteppingSet01" & Format(i, "00"), CStr(SetupVar.nSteppingSet(0, i)), lpFileName)
        Call INIWrite(lpModelName, "nSteppingSet02" & Format(i, "00"), CStr(SetupVar.nSteppingSet(1, i)), lpFileName)
        Call INIWrite(lpModelName, "nSteppingSet03" & Format(i, "00"), CStr(SetupVar.nSteppingSet(2, i)), lpFileName)
        Call INIWrite(lpModelName, "nSteppingSet04" & Format(i, "00"), CStr(SetupVar.nSteppingSet(3, i)), lpFileName)
    Next
    
    ' Offset
    For i = 0 To MAX_AD_CHANNEL
        Call INIWrite(lpModelName, "AdjustUse" & Format(i, "00"), CStr(SetupVar.bAdjustUse(i)), lpFileName)
        Call INIWrite(lpModelName, "Add" & Format(i, "00"), CStr(SetupVar.dAdd(i)), lpFileName)
        Call INIWrite(lpModelName, "Multi" & Format(i, "00"), CStr(SetupVar.dMulti(i)), lpFileName)
    Next
End Function

Public Sub SystemDisp2Mem()
    Dim i As Integer

    ' General
    SysVar.nPowerPort = frmSystem.txtPowerPort.Text
    SysVar.nNvhPort = frmSystem.txtNvhPort.Text
    SysVar.nStepPort(0) = frmSystem.txtSteppingPort(0).Text
    SysVar.nStepPort(1) = frmSystem.txtSteppingPort(1).Text
    SysVar.bPlcCommUse = IIf(frmSystem.chkPLCUse.Value = 1, True, False)
    SysVar.bScreenKeyboardUse = IIf(frmSystem.chkScreenKeyboardUse.Value = 1, True, False)
    SysVar.bCaptureSend = IIf(frmSystem.chkCaptureSend.Value = 1, True, False)
    SysVar.nReTest = frmSystem.txtRetest.Text
    SysVar.bNgTableUse = IIf(frmSystem.chkNgTable.Value = 1, True, False)
    SysVar.lpNgTableList = frmSystem.txtNgTableList.Text
    SysVar.bSideDoor = IIf(frmSystem.chkSideDoor.Value = 1, True, False)
    
    ' Hidden
    SysVar.nDBCol = frmSystem.txtDBCol.Text
    SysVar.nDBRow = frmSystem.txtDBRow.Text
    SysVar.lpPrintName = frmSystem.txtPrintName.Text
    SysVar.lpFtpIp = frmSystem.txtFtpInfo(0).Text
    SysVar.lpFtpPort = frmSystem.txtFtpInfo(1).Text
    SysVar.lpFtpId = frmSystem.txtFtpInfo(2).Text
    SysVar.lpFtpPw = frmSystem.txtFtpInfo(3).Text
    
    ' Correlation
    SysVar.nFiltering = Trim$(frmSystem.txtFiltering.Text)
    
    For i = 0 To MAX_AD_CHANNEL
        SysVar.lpName(i) = Trim$(frmSystem.txtName(i).Text)
        SysVar.dNaive(i) = Val(frmSystem.txtNaive(i).Text)
        SysVar.dAdd(i) = Val(frmSystem.txtAdd(i).Text)
        SysVar.dMulti(i) = Val(frmSystem.txtMulti(i).Text)
        SysVar.bMinus(i) = IIf(frmSystem.chkMinus(i).Value = 1, True, False)
        SysVar.bZero(i) = IIf(frmSystem.chkZero(i).Value = 1, True, False)
        SysVar.lpUnit(i) = Trim$(frmSystem.txtUnit(i).Text)
        SysVar.bPercent(i) = IIf(frmSystem.chkPercent(i).Value = 1, True, False)
    Next
    
    ' Mastersample
    SysVar.dMSVolt = Val(frmSystem.txtMSVolt.Text)
    
    For i = 0 To 1
        SysVar.bMSTest(i) = IIf(frmSystem.chkMSTEST(i).Value = 1, True, False)
    Next
    
    For i = 0 To frmSystem.optMSUse.UBound
        If frmSystem.optMSUse(i).Value = True Then
            SysVar.nMSUse = i
            Exit For
        End If
    Next
    
    SysVar.nMS2Time = Val(frmSystem.txtMS2Time.Text)
    
    For i = 0 To 3
        SysVar.nMS4Time(i) = Format(frmSystem.txtMS4Time(i).Text, "00")
    Next
    
    SysVar.nMSAfterDelay = Val(frmSystem.txtMSAfterDelay.Text)
    
    ' Leak
    SysVar.bLeakTest = IIf(frmSystem.chkLEAKTEST.Value = 1, True, False)
    SysVar.nLeakGroup = Val(frmSystem.txtLeakGroup.Text)
    
    ' Offset
    For i = 0 To MAX_AD_CHANNEL
        SysVar.bOffsetMSUse(i) = IIf(frmSystem.chkCalUse(i).Value = 1, True, False)
        SysVar.dOffsetMSMin(i) = Format(frmSystem.txtCalMin(i).Text, "#0.00")
        SysVar.dOffsetMSMax(i) = Format(frmSystem.txtCalMax(i).Text, "#0.00")
        SysVar.dOffsetMSDelay(i) = Format(frmSystem.txtCalDelay(i).Text, "#0.00")
    Next
End Sub

Public Sub SystemMem2Disp()
    Dim i As Integer
    
    ' General
    frmSystem.txtPowerPort.Text = Trim$(SysVar.nPowerPort)
    frmSystem.txtNvhPort.Text = Trim$(SysVar.nNvhPort)
    frmSystem.txtSteppingPort(0).Text = Trim$(SysVar.nStepPort(0))
    frmSystem.txtSteppingPort(1).Text = Trim$(SysVar.nStepPort(1))
    frmSystem.chkPLCUse.Value = IIf(SysVar.bPlcCommUse, 1, 0)
    frmSystem.chkScreenKeyboardUse.Value = IIf(SysVar.bScreenKeyboardUse, 1, 0)
    frmSystem.chkCaptureSend.Value = IIf(SysVar.bCaptureSend, 1, 0)
    frmSystem.txtRetest.Text = Trim$(SysVar.nReTest)
    frmSystem.chkNgTable.Value = IIf(SysVar.bNgTableUse, 1, 0)
    frmSystem.txtNgTableList = Trim$(SysVar.lpNgTableList)
    frmSystem.chkSideDoor.Value = IIf(SysVar.bSideDoor, 1, 0)
    
    ' Hidden
    frmSystem.txtDBCol.Text = Trim$(SysVar.nDBCol)
    frmSystem.txtDBRow.Text = Trim$(SysVar.nDBRow)
    frmSystem.txtPrintName.Text = Trim$(SysVar.lpPrintName)
    frmSystem.txtFtpInfo(0).Text = Trim$(SysVar.lpFtpIp)
    frmSystem.txtFtpInfo(1).Text = Trim$(SysVar.lpFtpPort)
    frmSystem.txtFtpInfo(2).Text = Trim$(SysVar.lpFtpId)
    frmSystem.txtFtpInfo(3).Text = Trim$(SysVar.lpFtpPw)
    
    ' Corellation
    frmSystem.txtFiltering.Text = Trim$(SysVar.nFiltering)
    
    For i = 0 To MAX_AD_CHANNEL
        frmSystem.txtName(i).Text = Trim$(SysVar.lpName(i))
        frmSystem.txtNaive(i).Text = Format(SysVar.dNaive(i), "#0.00")
        frmSystem.txtAdd(i).Text = Format(SysVar.dAdd(i), "#0.00")
        frmSystem.txtMulti(i).Text = Format(SysVar.dMulti(i), "#0.00")
        frmSystem.chkMinus(i).Value = IIf(SysVar.bMinus(i), 1, 0)
        frmSystem.chkZero(i).Value = IIf(SysVar.bZero(i), 1, 0)
        frmSystem.txtUnit(i).Text = Trim$(SysVar.lpUnit(i))
        frmSystem.txtDisp(i).Text = Format(0, SysVar.lpUnit(i))
        frmSystem.chkPercent(i).Value = IIf(SysVar.bPercent(i), 1, 0)
    Next
    
    ' Mastersample
    frmSystem.txtMSVolt.Text = Trim$(SysVar.dMSVolt)
    
    For i = 0 To 1
        frmSystem.chkMSTEST(i).Value = IIf(SysVar.bMSTest(i), 1, 0)
    Next
    
    frmSystem.optMSUse(SysVar.nMSUse).Value = True
    frmSystem.txtMS2Time.Text = Trim$(SysVar.nMS2Time)
    
    For i = 0 To 3
        frmSystem.txtMS4Time(i).Text = Format(SysVar.nMS4Time(i), "00")
    Next
    
    frmSystem.txtMSAfterDelay.Text = Trim$(SysVar.nMSAfterDelay)
    
    ' Leak
    frmSystem.chkLEAKTEST.Value = IIf(SysVar.bLeakTest, 1, 0)
    frmSystem.txtLeakGroup.Text = Trim$(SysVar.nLeakGroup)
    
    ' Offset
    For i = 0 To MAX_AD_CHANNEL
        frmSystem.chkCalUse(i).Value = IIf(SysVar.bOffsetMSUse(i), 1, 0)
        frmSystem.txtCalMin(i).Text = Format(SysVar.dOffsetMSMin(i), "#0.00")
        frmSystem.txtCalMax(i).Text = Format(SysVar.dOffsetMSMax(i), "#0.00")
        frmSystem.txtCalDelay(i).Text = Format(SysVar.dOffsetMSDelay(i), "#0.00")
    Next
End Sub

Public Function LoadSystemFile() As Boolean
    Dim i               As Integer
    Dim lpFileName      As String
    
    On Error Resume Next
    
    lpFileName = lpPath & INISYSTEMFILE
    LoadSystemFile = SearchFile(lpFileName)
    
    ' Hidden
    SysVar.lpModel = INIRead("SYSTEM", "Model", lpFileName)
    SysVar.lpSaveDate = INIRead("SYSTEM", "SaveDate", lpFileName)
    SysVar.lpSaveFileName = INIRead("SYSTEM", "SaveFileName", lpFileName)
    SysVar.lpSaveNgFileName = INIRead("SYSTEM", "SaveNgFileName", lpFileName)
    SysVar.lTotalCounter = INIRead("SYSTEM", "TotalCounter", lpFileName)
    SysVar.lOkCounter = INIRead("SYSTEM", "OkCounter", lpFileName)
    SysVar.lNgCounter = INIRead("SYSTEM", "NGCounter", lpFileName)
    SysVar.nDBCol = INIRead("SYSTEM", "DBCol", lpFileName)
    SysVar.nDBRow = INIRead("SYSTEM", "DBRow", lpFileName)
    SysVar.lpPrintName = INIRead("SYSTEM", "PrintName", lpFileName)
    SysVar.lpFtpIp = INIRead("SYSTEM", "FtpIp", lpFileName)
    SysVar.lpFtpPort = INIRead("SYSTEM", "FtpPort", lpFileName)
    SysVar.lpFtpId = INIRead("SYSTEM", "FtpId", lpFileName)
    SysVar.lpFtpPw = INIRead("SYSTEM", "FtpPw", lpFileName)
    SysVar.nContinueNGQty = INIRead("SYSTEM", "ContinueNGQty", lpFileName)
    SysVar.lpOldModelSave = INIRead("SYSTEM", "OldModelSave", lpFileName)
    SysVar.bSideDoor = INIRead("SYSTEM", "SideDoor", lpFileName)
    
    ' General
    SysVar.nPowerPort = INIRead("SYSTEM", "PowerPort", lpFileName)
    SysVar.nNvhPort = INIRead("SYSTEM", "NvhPort", lpFileName)
    SysVar.nStepPort(0) = INIRead("SYSTEM", "StepPort00", lpFileName)
    SysVar.nStepPort(1) = INIRead("SYSTEM", "StepPort01", lpFileName)
    SysVar.bPlcCommUse = IIf(INIRead("SYSTEM", "PlcCommUse", lpFileName) = "True", True, False)
    SysVar.bScreenKeyboardUse = IIf(INIRead("SYSTEM", "ScreenKeyboardUse", lpFileName) = "True", True, False)
    SysVar.bCaptureSend = IIf(INIRead("SYSTEM", "CaptureSend", lpFileName) = "True", True, False)
    SysVar.lpPassword = INIRead("SYSTEM", "Password", lpFileName)
    SysVar.nReTest = INIRead("SYSTEM", "ReTest", lpFileName)
    SysVar.bNgTableUse = IIf(INIRead("SYSTEM", "NgTableUse", lpFileName) = "True", True, False)
    SysVar.lpNgTableList = INIRead("SYSTEM", "NgTableList", lpFileName)
    
    ' Correlation
    SysVar.nFiltering = INIRead("SYSTEM", "Filtering", lpFileName)
    
    For i = 0 To MAX_AD_CHANNEL
        SysVar.lpName(i) = INIRead("SYSTEM", "AD-Name" & Format(i, "00"), lpFileName)
        SysVar.dNaive(i) = INIRead("SYSTEM", "AD-Naive" & Format(i, "00"), lpFileName)
        SysVar.dAdd(i) = INIRead("SYSTEM", "AD-Add" & Format(i, "00"), lpFileName)
        SysVar.dMulti(i) = INIRead("SYSTEM", "AD-Multi" & Format(i, "00"), lpFileName)
        SysVar.bMinus(i) = IIf(INIRead("SYSTEM", "AD-Minus" & Format(i, "00"), lpFileName) = "True", True, False)
        SysVar.bZero(i) = IIf(INIRead("SYSTEM", "AD-Zero" & Format(i, "00"), lpFileName) = "True", True, False)
        SysVar.lpUnit(i) = INIRead("SYSTEM", "AD-Unit" & Format(i, "00"), lpFileName)
        SysVar.bPercent(i) = IIf(INIRead("SYSTEM", "AD-Percent" & Format(i, "00"), lpFileName) = "True", True, False)
    Next
    
    ' Mastersample
    SysVar.dMSVolt = Val(INIRead("SYSTEM", "MSVolt", lpFileName))
    SysVar.bMSTest(0) = IIf(INIRead("SYSTEM", "MSTest00", lpFileName) = "True", True, False)
    SysVar.bMSTest(1) = IIf(INIRead("SYSTEM", "MSTest01", lpFileName) = "True", True, False)
    SysVar.nMSUse = Val(INIRead("SYSTEM", "MSUse", lpFileName))
    SysVar.nMS2Time = Val(INIRead("SYSTEM", "MS2Time", lpFileName))
    
    For i = 0 To 3
        SysVar.nMS4Time(i) = INIRead("SYSTEM", "MS4Time" & Format(i, "00"), lpFileName)
    Next
    
    SysVar.nMSAfterDelay = INIRead("SYSTEM", "MSAfterDelay", lpFileName)
    
    ' Leak
    SysVar.bLeakTest = IIf(INIRead("SYSTEM", "LEAKTest", lpFileName) = "True", True, False)
    SysVar.nLeakGroup = INIRead("SYSTEM", "LeakGroup", lpFileName)
    
    ' Offset
    For i = 0 To MAX_AD_CHANNEL
        SysVar.bOffsetMSUse(i) = IIf(INIRead("SYSTEM", "Offset-Use" & Format(i, "00"), lpFileName) = "True", True, False)
        SysVar.dOffsetMSMin(i) = INIRead("SYSTEM", "Offset-Min" & Format(i, "00"), lpFileName)
        SysVar.dOffsetMSMax(i) = INIRead("SYSTEM", "Offset-Max" & Format(i, "00"), lpFileName)
        SysVar.dOffsetMSDelay(i) = INIRead("SYSTEM", "Offset-Delay" & Format(i, "00"), lpFileName)
    Next
    
    ' label name
    For i = 0 To MAX_AD_CHANNEL
        frmSystem.pnlCalName(i).Caption = SysVar.lpName(i)
    Next
End Function

Public Function SaveSystemFile() As Boolean
    Dim i                   As Integer
    Dim lpFileName          As String
    
    On Error Resume Next
    
    lpFileName = lpPath & INISYSTEMFILE
    SaveSystemFile = SearchFile(lpFileName)
    
    ' Hidden
    Call INIWrite("SYSTEM", "Model", CStr(SysVar.lpModel), lpFileName)
    Call INIWrite("SYSTEM", "SaveDate", CStr(SysVar.lpSaveDate), lpFileName)
    Call INIWrite("SYSTEM", "SaveFileName", CStr(SysVar.lpSaveFileName), lpFileName)
    Call INIWrite("SYSTEM", "SaveNgFileName", CStr(SysVar.lpSaveNgFileName), lpFileName)
    Call INIWrite("SYSTEM", "TotalCounter", CStr(SysVar.lTotalCounter), lpFileName)
    Call INIWrite("SYSTEM", "OKCounter", CStr(SysVar.lOkCounter), lpFileName)
    Call INIWrite("SYSTEM", "NGCounter", CStr(SysVar.lNgCounter), lpFileName)
    Call INIWrite("SYSTEM", "DBCol", CStr(SysVar.nDBCol), lpFileName)
    Call INIWrite("SYSTEM", "DBRow", CStr(SysVar.nDBRow), lpFileName)
    Call INIWrite("SYSTEM", "PrintName", CStr(SysVar.lpPrintName), lpFileName)
    Call INIWrite("SYSTEM", "FtpIp", CStr(SysVar.lpFtpIp), lpFileName)
    Call INIWrite("SYSTEM", "FtpPort", CStr(SysVar.lpFtpPort), lpFileName)
    Call INIWrite("SYSTEM", "FtpId", CStr(SysVar.lpFtpId), lpFileName)
    Call INIWrite("SYSTEM", "FtpPw", CStr(SysVar.lpFtpPw), lpFileName)
    Call INIWrite("SYSTEM", "ContinueNGQty", CStr(SysVar.nContinueNGQty), lpFileName)
    Call INIWrite("SYSTEM", "OldModelSave", CStr(SysVar.lpOldModelSave), lpFileName)
    Call INIWrite("SYSTEM", "SideDoor", CStr(SysVar.bSideDoor), lpFileName)
    
    ' General
    Call INIWrite("SYSTEM", "PowerPort", CStr(SysVar.nPowerPort), lpFileName)
    Call INIWrite("SYSTEM", "NvhPort", CStr(SysVar.nNvhPort), lpFileName)
    Call INIWrite("SYSTEM", "StepPort00", CStr(SysVar.nStepPort(0)), lpFileName)
    Call INIWrite("SYSTEM", "StepPort01", CStr(SysVar.nStepPort(1)), lpFileName)
    Call INIWrite("SYSTEM", "PlcCommUse", CStr(SysVar.bPlcCommUse), lpFileName)
    Call INIWrite("SYSTEM", "ScreenKeyboardUse", CStr(SysVar.bScreenKeyboardUse), lpFileName)
    Call INIWrite("SYSTEM", "CaptureSend", CStr(SysVar.bCaptureSend), lpFileName)
    Call INIWrite("SYSTEM", "Password", CStr(SysVar.lpPassword), lpFileName)
    Call INIWrite("SYSTEM", "ReTest", CStr(SysVar.nReTest), lpFileName)
    Call INIWrite("SYSTEM", "NgTableUse", CStr(SysVar.bNgTableUse), lpFileName)
    Call INIWrite("SYSTEM", "NgTableList", CStr(SysVar.lpNgTableList), lpFileName)
    
    ' Correlation
    Call INIWrite("SYSTEM", "Filtering", CStr(SysVar.nFiltering), lpFileName)
    
    For i = 0 To MAX_AD_CHANNEL
        Call INIWrite("SYSTEM", "AD-Name" & Format(i, "00"), CStr(SysVar.lpName(i)), lpFileName)
        Call INIWrite("SYSTEM", "AD-Naive" & Format(i, "00"), CStr(SysVar.dNaive(i)), lpFileName)
        Call INIWrite("SYSTEM", "AD-Add" & Format(i, "00"), CStr(SysVar.dAdd(i)), lpFileName)
        Call INIWrite("SYSTEM", "AD-Multi" & Format(i, "00"), CStr(SysVar.dMulti(i)), lpFileName)
        Call INIWrite("SYSTEM", "AD-Minus" & Format(i, "00"), CStr(SysVar.bMinus(i)), lpFileName)
        Call INIWrite("SYSTEM", "AD-Zero" & Format(i, "00"), CStr(SysVar.bZero(i)), lpFileName)
        Call INIWrite("SYSTEM", "AD-Unit" & Format(i, "00"), CStr(SysVar.lpUnit(i)), lpFileName)
        Call INIWrite("SYSTEM", "AD-Percent" & Format(i, "00"), CStr(SysVar.bPercent(i)), lpFileName)
    Next
    
    ' Mastersample
    Call INIWrite("SYSTEM", "MSVolt", CStr(SysVar.dMSVolt), lpFileName)
    
    For i = 0 To 1
        Call INIWrite("SYSTEM", "MSTest" & Format(i, "00"), CStr(SysVar.bMSTest(i)), lpFileName)
    Next
    
    Call INIWrite("SYSTEM", "MSUse", CStr(SysVar.nMSUse), lpFileName)
    Call INIWrite("SYSTEM", "MS2Time", CStr(SysVar.nMS2Time), lpFileName)
    
    For i = 0 To 3
        Call INIWrite("SYSTEM", "MS4Time" & Format(i, "00"), Format(CStr(SysVar.nMS4Time(i)), "00"), lpFileName)
    Next
    
    Call INIWrite("SYSTEM", "MSAfterDelay", CStr(SysVar.nMSAfterDelay), lpFileName)
    
    ' Leak
    Call INIWrite("SYSTEM", "LEAKTest", CStr(SysVar.bLeakTest), lpFileName)
    Call INIWrite("SYSTEM", "LeakGroup", CStr(SysVar.nLeakGroup), lpFileName)
    
    ' Offset
    For i = 0 To MAX_AD_CHANNEL
        Call INIWrite("SYSTEM", "Offset-Use" & Format(i, "00"), CStr(SysVar.bOffsetMSUse(i)), lpFileName)
        Call INIWrite("SYSTEM", "Offset-Min" & Format(i, "00"), CStr(SysVar.dOffsetMSMin(i)), lpFileName)
        Call INIWrite("SYSTEM", "Offset-Max" & Format(i, "00"), CStr(SysVar.dOffsetMSMax(i)), lpFileName)
        Call INIWrite("SYSTEM", "Offset-Delay" & Format(i, "00"), CStr(SysVar.dOffsetMSDelay(i)), lpFileName)
    Next
End Function

Public Function LoadModelName() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim lpFileName As String
    
    On Error Resume Next
    
    lpFileName = lpPath & INIMODELFILE
    LoadModelName = SearchFile(lpFileName)
    
    Call INIModelNameRead
    
    For i = 0 To CARSAVECOUNT - 1
        SelectCar(i).ModelName = INIRead(Format(i, "00"), "ModelName", lpFileName)
        
        For j = 0 To SUBSAVECOUNT - 1
            SelectCar(i).ModelNameSub(j) = INIRead(Format(i, "00"), "ModelNameSub" & Format(j, "00"), lpFileName)
        Next
    Next
End Function

Public Function SaveModelName() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim lpFileName As String
    
    On Error Resume Next
    
    lpFileName = lpPath & INIMODELFILE
    SaveModelName = SearchFile(lpFileName)
    
    For i = 0 To CARSAVECOUNT - 1
        Call INIWrite(Format(i, "00"), "ModelName", Trim$(SelectCar(i).ModelName), lpFileName)
        
        For j = 0 To SUBSAVECOUNT - 1
            Call INIWrite(Format(i, "00"), "ModelNameSub" & Format(j, "00"), Trim$(SelectCar(i).ModelNameSub(j)), lpFileName)
        Next
    Next
End Function

Public Function LoadLangFile(ByVal nForm As Integer) As Boolean
    Dim i As Integer
    Dim lpFileName As String
    Dim lpTmp(9) As String
    
    On Error Resume Next
    
    lpFileName = lpPath & INILANGFILE
    LoadLangFile = SearchFile(lpFileName)
    
    lpLang = INIRead("language", "main", lpFileName)
    LangVar.Main = Trim$(lpLang)
    
    LangVar.MsgYes = INIRead(lpLang, "All-Yes", lpFileName)
    LangVar.MsgNo = INIRead(lpLang, "All-No", lpFileName)
    
    Select Case nForm
        Case FM_MAIN:
            For i = 0 To 5
                frmMain.btnMenu(i).Caption = INIRead(lpLang, "frmMain-Menu" & Format(i, "00"), lpFileName)
            Next
        
        Case FM_LOGIN:
            frmLogin.lblPassword.Caption = INIRead(lpLang, "frmLogin-Label-Password", lpFileName)
            frmLogin.btnLogin.Caption = INIRead(lpLang, "frmLogin-Button-OK", lpFileName)
            frmLogin.btnCancel.Caption = INIRead(lpLang, "frmLogin-Button-Cancel", lpFileName)
        
        Case FM_DATABASE:
            For i = 1 To 4
                frmDatabase.lblTemp(i).Caption = INIRead(lpLang, "frmDatabase-Label" & Format(i - 1, "00"), lpFileName)
            Next
            
            frmDatabase.btnOpen.Caption = INIRead(lpLang, "frmDatabase-Button-Open", lpFileName)
            frmDatabase.btnReturn.Caption = INIRead(lpLang, "frmDatabase-Button-Return", lpFileName)
        
        Case FM_RUN:
            frmRun.btnReturn.Caption = INIRead(lpLang, "frmRun-Button-Return", lpFileName)
            frmRun.btnClearCounter.Caption = INIRead(lpLang, "frmRun-Button-CountReset", lpFileName)
            frmRun.btnStatistical.Caption = INIRead(lpLang, "frmRun-Button-Statistical", lpFileName)
            
            For i = 0 To 2
                lpTmp(i) = INIRead(lpLang, "frmRun-Text-Mastersample" & Format(i, "00"), lpFileName)
            Next
            
            frmRun.pnlMS.Caption = vbCrLf & lpTmp(0) & vbCrLf & lpTmp(1) & vbCrLf & lpTmp(2)
            
            LangVar.Msg(MSG_STOP) = INIRead(lpLang, "frmRun-MSG-STOP", lpFileName)
            LangVar.Msg(MSG_RUN) = INIRead(lpLang, "frmRun-MSG-RUN", lpFileName)
            LangVar.Msg(MSG_READY) = INIRead(lpLang, "frmRun-MSG-READY", lpFileName)
            LangVar.Msg(MSG_OK) = INIRead(lpLang, "frmRun-MSG-OK", lpFileName)
            LangVar.Msg(MSG_NG) = INIRead(lpLang, "frmRun-MSG-NG", lpFileName)
            LangVar.Msg(MSG_CAL) = INIRead(lpLang, "frmRun-MSG-CAL", lpFileName)
            LangVar.Msg(MSG_ERR) = INIRead(lpLang, "frmRun-MSG-ERR", lpFileName)
            LangVar.Msg(MSG_OKPASS) = INIRead(lpLang, "frmRun-MSG-OKPASS", lpFileName)
            LangVar.Msg(MSG_OKFAIL) = INIRead(lpLang, "frmRun-MSG-OKFAIL", lpFileName)
            LangVar.Msg(MSG_NGPASS) = INIRead(lpLang, "frmRun-MSG-NGPASS", lpFileName)
            LangVar.Msg(MSG_NGFAIL) = INIRead(lpLang, "frmRun-MSG-NGFAIL", lpFileName)
            LangVar.Msg(MSG_SCANNER) = INIRead(lpLang, "frmRun-MSG-SCANNER", lpFileName)
            LangVar.Msg(MSG_NGTABLE) = INIRead(lpLang, "frmRun-MSG-NGTABLE", lpFileName)
            
            frmRun.pnlSerialName.Caption = INIRead(lpLang, "frmRun-Panel-Information-Serial", lpFileName)
            
            frmRun.pnlCarName(0).Caption = INIRead(lpLang, "frmRun-Panel-Information-CarType", lpFileName)
            frmRun.pnlCarName(1).Caption = INIRead(lpLang, "frmRun-Panel-Information-CarRank", lpFileName)
            frmRun.pnlCarName(2).Caption = INIRead(lpLang, "frmRun-Panel-Information-Group", lpFileName)
            frmRun.pnlCarName(3).Caption = INIRead(lpLang, "frmRun-Panel-Information-Pallet", lpFileName)
            
            frmRun.lblCounterName(0).Caption = INIRead(lpLang, "frmRun-Panel-Information-Total", lpFileName)
            frmRun.lblCounterName(1).Caption = INIRead(lpLang, "frmRun-Panel-Information-OK", lpFileName)
            frmRun.lblCounterName(2).Caption = INIRead(lpLang, "frmRun-Panel-Information-NG", lpFileName)
            frmRun.lblCounterName(3).Caption = INIRead(lpLang, "frmRun-Panel-Information-RATIO", lpFileName)
            
            frmRun.pnlInfoTitle.Caption = INIRead(lpLang, "frmRun-fraInformation", lpFileName)
            
            frmRun.pnlAct01Item(1).Caption = INIRead(lpLang, "frmRun-Act-Curr", lpFileName)
            frmRun.pnlAct01Item(2).Caption = INIRead(lpLang, "frmRun-Act-Volt", lpFileName)
            frmRun.pnlAct01Item(3).Caption = INIRead(lpLang, "frmRun-Act-Time", lpFileName)
            frmRun.pnlAct01Item(4).Caption = INIRead(lpLang, "frmRun-Act-Result", lpFileName)
            
            frmRun.pnlAct02Item(1).Caption = INIRead(lpLang, "frmRun-Act-Curr", lpFileName)
            frmRun.pnlAct02Item(2).Caption = INIRead(lpLang, "frmRun-Act-Volt", lpFileName)
            frmRun.pnlAct02Item(3).Caption = INIRead(lpLang, "frmRun-Act-Time", lpFileName)
            frmRun.pnlAct02Item(4).Caption = INIRead(lpLang, "frmRun-Act-Result", lpFileName)
            
            frmRun.pnlAct03Item(1).Caption = INIRead(lpLang, "frmRun-Act-Curr", lpFileName)
            frmRun.pnlAct03Item(2).Caption = INIRead(lpLang, "frmRun-Act-Volt", lpFileName)
            frmRun.pnlAct03Item(3).Caption = INIRead(lpLang, "frmRun-Act-Time", lpFileName)
            frmRun.pnlAct03Item(4).Caption = INIRead(lpLang, "frmRun-Act-Result", lpFileName)
            
            frmRun.pnlAct04Item(1).Caption = INIRead(lpLang, "frmRun-Act-Curr", lpFileName)
            frmRun.pnlAct04Item(2).Caption = INIRead(lpLang, "frmRun-Act-Volt", lpFileName)
            frmRun.pnlAct04Item(3).Caption = INIRead(lpLang, "frmRun-Act-Time", lpFileName)
            frmRun.pnlAct04Item(4).Caption = INIRead(lpLang, "frmRun-Act-Result", lpFileName)
            
            frmRun.pnlBlowerItem(1).Caption = INIRead(lpLang, "frmRun-Blw-Curr", lpFileName)
            frmRun.pnlBlowerItem(2).Caption = INIRead(lpLang, "frmRun-Blw-Time", lpFileName)
            frmRun.pnlBlowerItem(3).Caption = INIRead(lpLang, "frmRun-Blw-Result", lpFileName)
        
        Case FM_SETUP:
            frmSetup.btnReturn.Caption = INIRead(lpLang, "frmSetup-Button-Return", lpFileName)
            frmSetup.btnSave.Caption = INIRead(lpLang, "frmSetup-Button-Save", lpFileName)
            frmSetup.btnDelete.Caption = INIRead(lpLang, "frmSetup-Button-Delete", lpFileName)
            frmSetup.btnLoading.Caption = INIRead(lpLang, "frmSetup-Button-Loading", lpFileName)
        
        Case FM_SYSTEM:
            
            Select Case lpLang
                Case "korea": i = 0
                Case "english": i = 1
            End Select
            
            frmSystem.optLang(i).Value = True
            
            frmSystem.btnReturn.Caption = INIRead(lpLang, "frmSystem-Button-Return", lpFileName)
            frmSystem.btnSave.Caption = INIRead(lpLang, "frmSystem-Button-Save", lpFileName)
            
            frmSystem.fraPassword.Caption = INIRead(lpLang, "frmSystem-Frame-Password-Title", lpFileName)
            frmSystem.lblTemp(0).Caption = INIRead(lpLang, "frmSystem-Frame-Password-Current", lpFileName)
            frmSystem.lblTemp(1).Caption = INIRead(lpLang, "frmSystem-Frame-Password-New", lpFileName)
            frmSystem.lblTemp(2).Caption = INIRead(lpLang, "frmSystem-Frame-Password-Confirm", lpFileName)
            frmSystem.btnPassword.Caption = INIRead(lpLang, "frmSystem-Button-Password-Confirm", lpFileName)
        
        Case FM_EXIT:
            For i = 0 To 2
                frmExit.btnExitMenu(i).Caption = INIRead(lpLang, "frmExit-Button-Menu" & Format(i, "00"), lpFileName)
            Next
        
        Case FM_GRAPH:
            frmGraph.btnReturn.Caption = INIRead(lpLang, "frmGraph-Button-Return", lpFileName)
    
        Case FM_SETUPSELECT:
            frmSetupSelect.btnReturn.Caption = INIRead(lpLang, "frmSetupSelect-Button-Return", lpFileName)
            frmSetupSelect.btnSetup.Caption = INIRead(lpLang, "frmSetupSelect-Button-Setup", lpFileName)
            frmSetupSelect.btnNameSave.Caption = INIRead(lpLang, "frmSetupSelect-Button-NameSave", lpFileName)
            frmSetupSelect.btnModelLoading.Caption = INIRead(lpLang, "frmSetupSelect-Button-ModelLoading", lpFileName)
            frmSetupSelect.pnlTitle(0).Caption = INIRead(lpLang, "frmSetupSelect-Panel-Title00", lpFileName)
            frmSetupSelect.pnlTitle(1).Caption = INIRead(lpLang, "frmSetupSelect-Panel-Title01", lpFileName)
    
    End Select
End Function

Public Sub SaveLangFile()
    Dim lpFileName      As String
    Dim lpRes           As String
    Dim i               As Integer
    
    lpFileName = lpPath & INILANGFILE
    
    For i = 0 To frmSystem.optLang.UBound
        If frmSystem.optLang(i).Value = True Then
            Exit For
        End If
    Next
    
    Select Case i
        Case 0: lpRes = "korea"
        Case 1: lpRes = "english"
    End Select
    
    Call INIWrite("language", "main", lpRes, lpFileName)
    Call Sleep(100)
End Sub

Public Function LoadPlcFile() As Boolean
    Dim i               As Integer
    Dim lpFileName      As String
    
    On Error Resume Next
    
    lpFileName = lpPath & INIPLCFILE
    LoadPlcFile = SearchFile(lpFileName)
    
    If LoadPlcFile = False Then
        Call MsgBox("PLC ADDRESS FILE Nothing... ")
        
        Exit Function
    End If
    
    ' PLC
    PlcVar.lpAddrReady = INIRead("PLC", "AddrReady", lpFileName)
    PlcVar.lpAddrStart = INIRead("PLC", "AddrStart", lpFileName)
    PlcVar.lpAddrStatus = INIRead("PLC", "AddrStatus", lpFileName)
    PlcVar.lpAddrRunning = INIRead("PLC", "AddrRunning", lpFileName)
    PlcVar.lpAddrResult = INIRead("PLC", "AddrResult", lpFileName)
    PlcVar.lpAddrNGList = INIRead("PLC", "AddrNGList", lpFileName)
    PlcVar.lpAddrRunStart = INIRead("PLC", "AddrRunStart", lpFileName)
    
    ' INFO
    PlcVar.lpAddrCarType = INIRead("INFO", "AddrCarType", lpFileName)
    PlcVar.lpAddrCarRank = INIRead("INFO", "AddrCarRank", lpFileName)
    PlcVar.lpAddrCarGroup = INIRead("INFO", "AddrCarGroup", lpFileName)
    PlcVar.lpAddrPallet = INIRead("INFO", "AddrPallet", lpFileName)
    
    For i = 0 To 4
        PlcVar.lpAddrSerial(i) = INIRead("INFO", "AddrSerial" & Format(i, "00"), lpFileName)
    Next
    
    ' DATATRACKING
    For i = 0 To 10
        PlcVar.lpAddrDataTracking(i) = INIRead("DATATRACKING", "AddrDataTracking" & Format(i, "00"), lpFileName)
    Next
    
    ' OTHER
    For i = 0 To 2
        PlcVar.lpAddrLoadModel(i) = INIRead("OTHER", "AddrLoadModel" & Format(i, "00"), lpFileName)
        PlcVar.nNextLoadModel(i) = INIRead("OTHER", "NextLoadModel" & Format(i, "00"), lpFileName)
        PlcVar.nSizeLoadModel(i) = INIRead("OTHER", "SizeLoadModel" & Format(i, "00"), lpFileName)
    Next
    
    PlcVar.lpAddrLeak(0) = INIRead("OTHER", "AddrLeak00", lpFileName)
    
    For i = 0 To 1
        PlcVar.lpAddrBarcode(i) = INIRead("OTHER", "AddrBarcode" & Format(i, "00"), lpFileName)
    Next
    
    PlcVar.lpAddrVibResult = INIRead("OTHER", "AddrVibResult", lpFileName)
End Function

Public Sub SaveDataFile(ByVal lpFileName As String)
    Dim i As Integer
    Dim nFileNo As Integer
    
    On Error GoTo ErrHandler_SaveDataFile
    
    nFileNo = FreeFile
    
    If Len(Dir(lpFileName, vbReadOnly)) = 0 Then
        Open lpFileName For Output As #nFileNo
            Print #nFileNo, "MODEL,";
            Print #nFileNo, "SERIAL,";
            Print #nFileNo, "PALLET,";
            Print #nFileNo, "TIME,";
            
            Print #nFileNo, "RESULT,";
            
            ' ACT01
            Print #nFileNo, "ACT01 ST1 CURR,";
            Print #nFileNo, "ACT01 ST1 VOLT,";
            Print #nFileNo, "ACT01 ST1 TIME,";
            
            For i = 1 To 5
                Print #nFileNo, "ACT01 P" & i & " CURR,";
                Print #nFileNo, "ACT01 P" & i & " VOLT,";
                Print #nFileNo, "ACT01 P" & i & " TIME,";
            Next
            
            Print #nFileNo, "ACT01 ST2 CURR,";
            Print #nFileNo, "ACT01 ST2 VOLT,";
            Print #nFileNo, "ACT01 ST2 TIME,";
            
            Print #nFileNo, "ACT01 FINAL CURR,";
            Print #nFileNo, "ACT01 FINAL VOLT,";
            Print #nFileNo, "ACT01 FINAL TIME,";
            
            Print #nFileNo, "ACT01 STALL,";
            
            ' ACT02
            Print #nFileNo, "ACT02 ST1 CURR,";
            Print #nFileNo, "ACT02 ST1 VOLT,";
            Print #nFileNo, "ACT02 ST1 TIME,";
            
            For i = 1 To 2
                Print #nFileNo, "ACT02 P" & i & " CURR,";
                Print #nFileNo, "ACT02 P" & i & " VOLT,";
                Print #nFileNo, "ACT02 P" & i & " TIME,";
            Next
            
            Print #nFileNo, "ACT02 ST2 CURR,";
            Print #nFileNo, "ACT02 ST2 VOLT,";
            Print #nFileNo, "ACT02 ST2 TIME,";
            
            Print #nFileNo, "ACT02 FINAL CURR,";
            Print #nFileNo, "ACT02 FINAL VOLT,";
            Print #nFileNo, "ACT02 FINAL TIME,";
            
            Print #nFileNo, "ACT02 STALL,";
            
            ' ACT03
            Print #nFileNo, "ACT03 ST1 CURR,";
            Print #nFileNo, "ACT03 ST1 VOLT,";
            Print #nFileNo, "ACT03 ST1 TIME,";
            
            For i = 1 To 2
                Print #nFileNo, "ACT03 P" & i & " CURR,";
                Print #nFileNo, "ACT03 P" & i & " VOLT,";
                Print #nFileNo, "ACT03 P" & i & " TIME,";
            Next
            
            Print #nFileNo, "ACT03 ST2 CURR,";
            Print #nFileNo, "ACT03 ST2 VOLT,";
            Print #nFileNo, "ACT03 ST2 TIME,";
            
            Print #nFileNo, "ACT03 FINAL CURR,";
            Print #nFileNo, "ACT03 FINAL VOLT,";
            Print #nFileNo, "ACT03 FINAL TIME,";
            
            Print #nFileNo, "ACT03 STALL,";
            
            ' ACT04
            Print #nFileNo, "ACT04 ST1 CURR,";
            Print #nFileNo, "ACT04 ST1 VOLT,";
            Print #nFileNo, "ACT04 ST1 TIME,";
            
            For i = 1 To 2
                Print #nFileNo, "ACT04 P" & i & " CURR,";
                Print #nFileNo, "ACT04 P" & i & " VOLT,";
                Print #nFileNo, "ACT04 P" & i & " TIME,";
            Next
            
            Print #nFileNo, "ACT04 ST2 CURR,";
            Print #nFileNo, "ACT04 ST2 VOLT,";
            Print #nFileNo, "ACT04 ST2 TIME,";
            
            Print #nFileNo, "ACT04 FINAL CURR,";
            Print #nFileNo, "ACT04 FINAL VOLT,";
            Print #nFileNo, "ACT04 FINAL TIME,";
            
            Print #nFileNo, "ACT04 STALL,";
            
            ' SENSOR
            Print #nFileNo, "SENSOR 1,";
            Print #nFileNo, "SENSOR 2,";
            Print #nFileNo, "SENSOR 3,";
            Print #nFileNo, "SENSOR 4,";
            Print #nFileNo, "SENSOR 5,";
            Print #nFileNo, "SENSOR 6,";
            
            Print #nFileNo, "MEMO"
        Close #nFileNo
    End If
    
    Open lpFileName For Append As #nFileNo
        ' General
        Print #nFileNo, DataVar.lpModel; ",";
        Print #nFileNo, DataVar.lpSerialNo; ",";
        Print #nFileNo, DataVar.lpPallet; ",";
        Print #nFileNo, DataVar.lpTime; ",";
        Print #nFileNo, DataVar.lpResult; ",";
        
        ' ACT01
        For i = 0 To 7
            Print #nFileNo, DataVar.lpAct01Curr(i); ",";
            Print #nFileNo, DataVar.lpAct01Volt(i); ",";
            Print #nFileNo, DataVar.lpAct01Time(i); ",";
        Next
        
        Print #nFileNo, DataVar.lpAct01Stall(0); ",";
        
        ' ACT02
        For i = 0 To 4
            Print #nFileNo, DataVar.lpAct02Curr(i); ",";
            Print #nFileNo, DataVar.lpAct02Volt(i); ",";
            Print #nFileNo, DataVar.lpAct02Time(i); ",";
        Next
        
        Print #nFileNo, DataVar.lpAct02Stall(0); ",";
        
        ' ACT03
        For i = 0 To 4
            Print #nFileNo, DataVar.lpAct03Curr(i); ",";
            Print #nFileNo, DataVar.lpAct03Volt(i); ",";
            Print #nFileNo, DataVar.lpAct03Time(i); ",";
        Next
        
        Print #nFileNo, DataVar.lpAct03Stall(0); ",";
        
        ' ACT04
        For i = 0 To 4
            Print #nFileNo, DataVar.lpAct04Curr(i); ",";
            Print #nFileNo, DataVar.lpAct04Volt(i); ",";
            Print #nFileNo, DataVar.lpAct04Time(i); ",";
        Next
        
        Print #nFileNo, DataVar.lpAct04Stall(0); ",";
        
        ' SENSOR
        For i = 0 To 5
            Print #nFileNo, DataVar.lpSensor(i); ",";
        Next
        
        Print #nFileNo, DataVar.lpMemo
    Close #nFileNo
    
    Exit Sub
    
ErrHandler_SaveDataFile:
    
    Close #nFileNo
    
    Call MsgBox("Error Occurred : " + vbCrLf + "Function : SaveDataFile()", vbCritical + vbOKOnly, "Error")
End Sub

Private Sub MakeStatistical()
    Call SearchDir(lpPath + "\StatiDataFile")
    Call SearchDir(lpPath + "\StatiDataFile\" & Format(Date, "YYYY"))
    Call SearchDir(lpPath + "\StatiDataFile\" & Format(Date, "YYYY") & "\" & Format(Date, "YYYYMM"))
End Sub

Public Sub GetStatistical(ByRef RefStcalVar() As StatisticalVariables)
    Dim vFileNumber As Variant
    Dim lpFileName As String
    Dim lpString As String
    Dim i As Integer
    Dim nFileNo As Integer
    
    On Error Resume Next
    
    nFileNo = FreeFile
    
    Erase RefStcalVar
    
    For i = 0 To 50
        RefStcalVar(i).lpModelName = ""
    Next
    
    Call MakeStatistical
    
    lpFileName = lpPath & "\StatiDataFile\" & Format(Date, "YYYY") & "\" & Format(Date, "YYYYMM") & "\" + Format(Date, "YYYYMMDD") + ".csv"
    
    If SearchFile(lpFileName) Then
        Open lpFileName For Input As #nFileNo
        
        vFileNumber = LOF(nFileNo)
        
        If vFileNumber = 0 Then
            Close #nFileNo
        Else
            If vFileNumber > 36 Then
                For vFileNumber = 1 To 5
                    Input #nFileNo, lpString
                Next
                
                i = 0
                
                Do While Not EOF(nFileNo)
                    Input #nFileNo, lpString: RefStcalVar(i).lpModelName = Trim$(lpString)
                    Input #nFileNo, lpString: RefStcalVar(i).lAllCounter = Val(lpString)
                    Input #nFileNo, lpString: RefStcalVar(i).lOkCounter = Val(lpString)
                    Input #nFileNo, lpString: RefStcalVar(i).lNgCounter = Val(lpString)
                    Input #nFileNo, lpString: RefStcalVar(i).dPercent = Val(lpString)
                    
                    i = i + 1
                    
                    If EOF(nFileNo) = True Then
                        Exit Do
                    End If
                Loop
            End If
        End If
        Close #nFileNo
    Else
        Open lpFileName For Output As #nFileNo
            Print #nFileNo, "Model,Total,OK,NG,Ratio"
            Print #nFileNo, "ALL,0,0,0,0.0"
        Close #nFileNo
    End If
    
    On Error GoTo 0
End Sub

Public Sub PlusStatistical(ByVal MdN As String, ByVal bResult As Boolean)
    Dim lpFileName As String
    Dim i As Integer
    Dim nFileNo As Integer
    
    nFileNo = FreeFile
    
    For i = 1 To 50
        If Trim$(StcalVar(i).lpModelName) = "" Or Trim$(MdN) = Trim$(StcalVar(i).lpModelName) Then
            GoTo INSERT_D
        End If
    Next
    If i > 50 Then i = 1

INSERT_D:
    
    StcalVar(i).lpModelName = Trim$(MdN)
    
    If bResult Then
        StcalVar(i).lOkCounter = StcalVar(i).lOkCounter + 1
    Else
        StcalVar(i).lNgCounter = StcalVar(i).lNgCounter + 1
    End If
    
    StcalVar(i).lAllCounter = StcalVar(i).lOkCounter + StcalVar(i).lNgCounter
    StcalVar(i).dPercent = Format((1 - (StcalVar(i).lNgCounter / StcalVar(i).lAllCounter)) * 100, "#0.0")
    
    If bResult Then
        StcalVar(0).lOkCounter = StcalVar(0).lOkCounter + 1
    Else
        StcalVar(0).lNgCounter = StcalVar(0).lNgCounter + 1
    End If
    
    StcalVar(0).lAllCounter = StcalVar(0).lOkCounter + StcalVar(0).lNgCounter
    StcalVar(0).dPercent = Format((1 - (StcalVar(0).lNgCounter / StcalVar(0).lAllCounter)) * 100, "#0.0")
    
    Call MakeStatistical
    
    ' 일별데이터파일구함
    lpFileName = lpPath & "\StatiDataFile\" & Format(Date, "YYYY") & "\" & Format(Date, "YYYYMM") & "\" + Format(Date, "YYYYMMDD") + ".csv"
    
    Open lpFileName For Output As #nFileNo
        Print #nFileNo, "Model,Total,OK,NG,Ratio"
        Print #nFileNo, "ALL,"; Trim$(str(StcalVar(0).lAllCounter)); ","; Trim$(str(StcalVar(0).lOkCounter)); ","; Trim$(str(StcalVar(0).lNgCounter)); ","; Format(StcalVar(0).dPercent, "#0.0")
        
        For i = 1 To 50
            If Trim$(StcalVar(i).lpModelName) = "" Then Exit For
            
            Print #nFileNo, Trim$(StcalVar(i).lpModelName); ","; Trim$(str(StcalVar(i).lAllCounter)); ","; Trim$(str(StcalVar(i).lOkCounter)); ","; Trim$(str(StcalVar(i).lNgCounter)); ","; Format(StcalVar(i).dPercent, "#0.0")
        Next
    Close #nFileNo
    
    On Error GoTo 0
End Sub

