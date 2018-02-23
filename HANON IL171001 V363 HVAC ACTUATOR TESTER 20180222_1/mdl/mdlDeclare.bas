Attribute VB_Name = "mdlDeclare"
Option Explicit

Public Type TagSystemVariables
    ' Hidden
    lpModel As String
    lpSaveDate As String
    lpSaveFileName As String
    lpSaveNgFileName As String
    lTotalCounter As Long
    lOkCounter As Long
    lNgCounter As Long
    nDBCol As Integer
    nDBRow As Integer
    lpPrintName As String
    lpFtpIp As String
    lpFtpPort As String
    lpFtpId As String
    lpFtpPw As String
    nContinueNGQty As Integer
    lpOldModelSave As String
    
    ' General
    nPowerPort As Integer
    nStepPort(1) As Integer
    nNvhPort As Integer
    bPlcCommUse As Boolean
    bScreenKeyboardUse As Boolean
    bCaptureSend As Boolean
    lpPassword As String
    nReTest As Integer
    bNgTableUse As Boolean
    lpNgTableList As String
    bSideDoor As Boolean
    
    ' Correlation
    nFiltering As Integer
    lpName(MAX_AD_CHANNEL) As String
    
    dNaive(MAX_AD_CHANNEL) As Double
    dAdd(MAX_AD_CHANNEL) As Double
    dMulti(MAX_AD_CHANNEL) As Double
    
    bMinus(MAX_AD_CHANNEL) As Boolean
    bZero(MAX_AD_CHANNEL) As Boolean
    lpUnit(MAX_AD_CHANNEL) As String
    bPercent(MAX_AD_CHANNEL) As Boolean
    
    ' Mastersample
    dMSVolt As Double
    bMSTest(1) As Boolean
    nMSUse As Integer
    nMS2Time As Integer
    nMS4Time(3) As Integer
    nMSAfterDelay As Integer
    
    bOffsetMSUse(MAX_AD_CHANNEL) As Boolean
    dOffsetMSMin(MAX_AD_CHANNEL) As Double
    dOffsetMSMax(MAX_AD_CHANNEL) As Double
    dOffsetMSDelay(MAX_AD_CHANNEL) As Double
    
    ' Leak
    bLeakTest As Boolean
    nLeakGroup As Integer
End Type


Public Type TagSetupVariables
    ' General
    dTestVolt As Double
    lpFileName As String
    bDataSave As Boolean
    nModelType As Integer
    lpActName(9) As String
    nActBoardNo(9) As Integer
    bScannerUse As Boolean
    nScannerValue As Integer
    
    ' Blower
    bBlowerUse As Boolean
    nBlowerType As Integer  ' PWM / FET
    nBlowerDirection As Integer
    nLinSpeed(9) As Integer
    
    lpBlowerName(9) As String
    dBlowerCurrLo(9) As Double
    dBlowerCurrHi(9) As Double
    dBlowerTime(9) As Double
    
    ' RPM
    lpRpmName As String
    dRpmCurrLo As Double
    dRpmCurrHi As Double
    
    ' Vibration
    bVibUse As Boolean
    lpVibName As String
    dVibCurrLo As Double
    dVibCurrHi As Double
    nVibResultType As Integer  ' Peak / RMS
    dVibStart As Double
    dVibEnd As Double
    nVibMethod As Integer  ' With Hi / After Hi
    dVibVolt As Double
    dVibTime As Double
    
    ' Act 01
    bAct01Use As Boolean
    nAct01Direction As Integer
    nAct01TestType As Integer
    
    lpAct01Name(9) As String
    dAct01SetVolt(9) As Double
    dAct01CurrLo(9) As Double
    dAct01CurrHi(9) As Double
    dAct01VoltLo(9) As Double
    dAct01VoltHi(9) As Double
    dAct01TimeLo(9) As Double
    dAct01TimeHi(9) As Double
    dAct01StallDeltaVoltLo As Double
    dAct01StallDeltaVoltHi As Double
    
    ' Act 02
    bAct02Use As Boolean
    nAct02Direction As Integer
    nAct02TestType As Integer
    
    lpAct02Name(9) As String
    dAct02SetVolt(9) As Double
    dAct02CurrLo(9) As Double
    dAct02CurrHi(9) As Double
    dAct02VoltLo(9) As Double
    dAct02VoltHi(9) As Double
    dAct02TimeLo(9) As Double
    dAct02TimeHi(9) As Double
    dAct02StallDeltaVoltLo As Double
    dAct02StallDeltaVoltHi As Double
    
    ' Act 03
    bAct03Use As Boolean
    nAct03Direction As Integer
    nAct03TestType As Integer
    
    lpAct03Name(9) As String
    dAct03SetVolt(9) As Double
    dAct03CurrLo(9) As Double
    dAct03CurrHi(9) As Double
    dAct03VoltLo(9) As Double
    dAct03VoltHi(9) As Double
    dAct03TimeLo(9) As Double
    dAct03TimeHi(9) As Double
    dAct03StallDeltaVoltLo As Double
    dAct03StallDeltaVoltHi As Double
    
    ' Act 04
    bAct04Use As Boolean
    nAct04Direction As Integer
    nAct042Pin As Integer
    nAct042PinPos As Integer
    nAct04TestType As Integer
    
    lpAct04Name(9) As String
    dAct04SetVolt(9) As Double
    dAct04CurrLo(9) As Double
    dAct04CurrHi(9) As Double
    dAct04VoltLo(9) As Double
    dAct04VoltHi(9) As Double
    dAct04TimeLo(9) As Double
    dAct04TimeHi(9) As Double
    dAct04StallDeltaVoltLo As Double
    dAct04StallDeltaVoltHi As Double
    
    ' Curr Count
    nActPeakCurrCount(9) As Integer
    nActEndPosCount(9) As Integer
    
    ' Sensor
    bSensorUse(9) As Boolean
    lpSensorName(9) As String
    dSensorCurrLo(9) As Double
    dSensorCurrHi(9) As Double
    dSensorTime(9) As Double
    
    ' PTC
    bPTCUse As Boolean
    lpPTCName As String
    dPTCCurrLo As Double
    dPTCCurrHi As Double
    dPTCTime As Double
    
    ' Ion
    lpIonName As String
    bIonUse As Boolean
    lpIonSubName(1) As String
    dIonLo(1) As Double
    dIonHi(1) As Double
    
    ' Leak
    bLeakUse(1) As Boolean
    lpLeakName(1) As String
    nLeakModel As Integer
    
    ' Nvh
    bNvhUse As Boolean
    
    ' Vision
    bVisionUse As Boolean
    lpVisionName(9) As String
    bVisionDoorUse(9) As Boolean
    lpVisionDoorName(9) As String
    
    nOpenCameraNo(9) As Integer
    nOpenCameraPos(9) As Integer
    nCloseCameraNo(9) As Integer
    nCloseCameraPos(9) As Integer
    
    ' BarCode
    bBarCodeUse As Boolean
    nBarcodeType As Integer
    lpBarcode(9) As String
    
    ' Marking
    bMarkingUse(9) As Boolean
    dMarkingTime(9) As Double
    
    ' Parts
    bPartUse(256) As Boolean
    bPartStatus(256) As Boolean
    lpPartName(256) As String
    bProductUse As Boolean
    lpProductList As String
    lpProductName As String
    bModelTypeUse As Boolean
    lpModelLHDList As String
    lpLHDPartName As String
    lpModelRHDList As String
    lpRHDPartName As String
    
    ' Lin
    bLinActUse(3) As Boolean
    lpLinActName(3) As String
    dLinActLo(3) As Double
    dLinActHi(3) As Double
    dLinActCurrLo(3) As Double
    dLinActCurrHi(3) As Double
    nLinAct01Check(9) As Integer
    nLinAct02Check(9) As Integer
    nLinAct03Check(9) As Integer
    nLinAct04Check(9) As Integer
    dLinAct01CheckTime(9) As Double
    dLinAct02CheckTime(9) As Double
    dLinAct03CheckTime(9) As Double
    dLinAct04CheckTime(9) As Double
    nLinActMove(3) As Integer
    nLinActFinal(3) As Integer
    bAutoAddressUse As Boolean
    nLinTestType As Integer
    nLinActFirstMove(3) As Integer
    nLinAct01RefPos As Integer
    bCheckPointUse As Boolean
    bStallUse As Boolean
    lLinActAngle(9) As Long
    dLinActTime(9) As Double
    bLinBlowerUse As Boolean
    lpLinBlowerName As String
    dLinBlowerTime As Double
    
    ' Stepping
    nSteppingSet(9, 19) As Integer
    
    ' Adjustment
    bAdjustUse(99) As Boolean
    dMulti(99) As Double
    dAdd(99) As Double
End Type

Public Type TagDataVariables
    lpModel As String
    lpSerialNo As String
    lpModelNo As String
    lpModelRank As String
    lpModelGroup As String
    lpPallet As String
    lpTime As String
    lpResult As String
    
    lpBlowerCurr(9) As String
    lpBlowerTime(9) As String
    lpRpm As String
    lpVib As String
    
    lpAct01Curr(9) As String
    lpAct01Volt(9) As String
    lpAct01Time(9) As String
    lpAct01Stall(9) As String
    
    lpAct02Curr(9) As String
    lpAct02Volt(9) As String
    lpAct02Time(9) As String
    lpAct02Stall(9) As String
    
    lpAct03Curr(9) As String
    lpAct03Volt(9) As String
    lpAct03Time(9) As String
    lpAct03Stall(9) As String
    
    lpAct04Curr(9) As String
    lpAct04Volt(9) As String
    lpAct04Time(9) As String
    lpAct04Stall(9) As String
    
    lpIon(9) As String
    lpSensor(9) As String
    
    lpLeak(1) As String
    lpBarcode(1) As String
    
    lpLinActMove(9) As String
    lpLinActFinal(9) As String
    lpLinActTime(9) As String
    lpPtc As String
    
    lpVision As String
    
    lpPartOK As String
    lpPartNG As String
    
    lpDoor(9) As String
    
    lpMemo As String
End Type

Public Type StatisticalVariables
    lpModelName         As String
    lAllCounter         As Long
    lOkCounter          As Long
    lNgCounter          As Long
    dPercent            As Double
End Type

' 프로그램 마다 변경되기에 되도록 삭제는 자제...
Public Type TagPlcVariables
    ' PLC
    lpAddrReady As String
    lpAddrStart As String
    lpAddrStatus As String
    lpAddrRunning As String
    lpAddrResult As String
    lpAddrNGList As String
    lpAddrRunStart As String
    
    ' INFO
    lpAddrPallet As String ' 파렛트
    lpAddrCarType As String ' 차종
    lpAddrCarRank As String ' 서열
    lpAddrCarGroup As String ' 그룹
    lpAddrSerial(10) As String
    
    ' DATATRACKING
    lpAddrDataTracking(10) As String
    
    ' OTHER
    lpAddrLoadModel(2) As String
    nNextLoadModel(2) As Integer
    nSizeLoadModel(2) As Integer
    lpAddrVibResult As String
    lpAddrLeak(1) As String
    lpAddrBarcode(1) As String
    
    lGetPlcData(20) As Long
    lTotalResult(9) As Long
End Type

Type ModelSourceVariables
    ModelName As String
    ModelNameSub(99) As String
End Type

Type LangVariables
    Main As String
    MsgYes As String
    MsgNo As String
    Msg(20) As String
End Type

Public Type TagRunVariables
    bLoading As Boolean
    
    nDispCounter As Integer
    bDispFlash As Boolean
    
    bAutoManual As Boolean
    bRun As Boolean
    bRestart As Boolean
    nTestDelay As Integer
    
    bStopFlag As Boolean
    
    bBlowerUse As Boolean
    bAct01Use As Boolean ' mode 1
    bAct02Use As Boolean ' temp 1
    bAct03Use As Boolean ' temp 2
    bAct04Use As Boolean ' intake
    bSensorUse(9) As Boolean
    bIonUse As Boolean
    bVibUse As Boolean
    bRpmUse As Boolean
    bLinAct01Use As Boolean
    bLinAct02Use As Boolean
    bLinAct03Use As Boolean
    bLinAct04Use As Boolean
    bLinPtcUse As Boolean
    bLinBlowerUse As Boolean
    
    bTestEnd(99) As Boolean
    
    ' 각 위치
    nTestPos As Integer
    nBlowerPos As Integer
    nAct01Pos As Integer
    nAct02Pos As Integer
    nAct03Pos As Integer
    nAct04Pos As Integer
    nIonPos As Integer
    nSensorPos As Integer
    nNvhPos As Integer
    nLinAct01Pos As Integer
    nLinAct02Pos As Integer
    nLinAct03Pos As Integer
    nLinAct04Pos As Integer
    nLinPtcPos As Integer
    nLinBlowerPos As Integer
    
    ' Lin
    bLinDataResult(9, 9) As Boolean
    lLinDataCP(9, 9) As Long
    bLinCheckPoint(9, 9) As Boolean
    dLinCPTime(9, 9) As Double
    lLinDataFinal(9) As Long
    lLinDataMove(9, 9) As Long
    bLinReadRes(9, 1) As Boolean
    bLinRefPos(9) As Boolean
    nLinAct01CheckPos(9) As Long
    nLinAct02CheckPos(9) As Long
    nLinAct03CheckPos(9) As Long
    nLinAct04CheckPos(9) As Long
    
    ' Blower
    bBlowerAddr(POS_BLOWER_HI) As Boolean ' Stack Address Position
    
    nVibCount As Integer
    nBlowerCount As Integer
    
    sAct01Addr As Variant ' Stack Address Position
    sAct02Addr As Variant ' Stack Address Position
    sAct03Addr As Variant ' Stack Address Position
    sAct04Addr As Variant ' Stack Address Position
    
    nAct01MaxLoop As Integer
    nAct02MaxLoop As Integer
    nAct03MaxLoop As Integer
    nAct04MaxLoop As Integer
    
    nActPeakCurrCount(9) As Integer
    nActEndPosCount(9) As Integer
    nActCount(9) As Integer
    
    ' Nvh
    bNvhUse As Boolean
    
    ' Door
    bDoorUse(MAX_DIO_CHANNEL) As Boolean
    bDoorStatus(MAX_DIO_CHANNEL) As Boolean
    bDoorResult(MAX_DIO_CHANNEL) As Boolean
    
    ' Other
    bFinal As Boolean
    lpNowDate As String
    bUpdate As Boolean
    bUpdateDate As Boolean
    
    ' 재 테스트
    bReBlowerUse As Boolean
    bReAct01Use As Boolean
    bReAct02Use As Boolean
    bReAct03Use As Boolean
    bReAct04Use As Boolean
    bReSensorUse(9) As Boolean
    bReIonUse As Boolean
    bReLeakUse(1) As Boolean
    bReVisionUse As Boolean
    bReNvhUse As Boolean
    bReLinAct01Use As Boolean
    bReLinAct02Use As Boolean
    bReLinAct03Use As Boolean
    bReLinAct04Use As Boolean
    bReLinPtcUse As Boolean
    bReLinBlowerUse As Boolean
    
    nMSUse As Integer  ' Calibration
    bMSStart As Boolean
    bMSTest(1) As Boolean
    bMSJudge(1) As Boolean
    nMSPos As Integer
    bMSResult(MAX_AD_CHANNEL)   As Boolean
    bMSTotalResult  As Boolean
End Type

Public Type TagManualVariables
    nBlowerPos As Integer
    nAct01Pos As Integer
    nAct02Pos As Integer
    nAct03Pos As Integer
    nAct04Pos As Integer
    nLinAct01Pos As Integer
    nLinAct02Pos As Integer
    nLinAct03Pos As Integer
    nLinAct04Pos As Integer
    nSensorPos As Integer
    nNvhPos As Integer
    nLinPos As Integer
    nIonPos As Integer
End Type

Public Type TagGraphVariables
    XS As Double ' X축시작
    XE As Double ' X축끝
    
    YS As Double ' Y축시작
    YE As Double ' Y축끝
    
    OVY As Double
    OVX As Double
    OVZ As Double
    
    PS As Integer
    PE As Integer
    
    nPeak As Integer
End Type

Public Type TagActNoControlVariables
    AD_CURR As Integer
    AD_VOLT As Integer
    O_POWER As Integer
    DA_NO As Integer
End Type

Public ActNo(9) As TagActNoControlVariables
Public RunVar As TagRunVariables
Public ManuVar As TagManualVariables
Public GraphVar As TagGraphVariables
Public SysVar As TagSystemVariables
Public SetupVar As TagSetupVariables
Public DataVar As TagDataVariables
Public StcalVar(50) As StatisticalVariables
Public PlcVar As TagPlcVariables
Public SelectCar(99) As ModelSourceVariables
Public LangVar As LangVariables

Public ZeroCurr(MAX_AD_CHANNEL) As Double
Public bLogin As Boolean
Public lpPath As String

Public KeyBoardObj As TextBox ' 화면키보드에서 전달되어질 텍스트박스이름
Public bKeyBoardNum As Boolean ' 화면키보드에서 전달되어질값이 true 면 상수만 입력

Public bSplash As Boolean ' SPLASH 하기

Public bGlobalStartSw As Boolean
Public bGlobalStopSw As Boolean
Public bPlcStartSig As Boolean
Public bPlcStopSig As Boolean
Public nStartSignal As Integer

Public lpNowModel As String
Public nNowModelNo As Integer
Public lpNowModelName As String
Public lpOldModelName As String

Public bRunningGraphWin As Boolean ' 통계 버튼을 누르고 리턴하면 순서가 뒤바뀌어 Calibration  한다.

Public nNowForm As Integer
Public lpLang As String

Public lToggle As Long

Public nNowTime As Integer
Public nMSTime As Integer
Public bMSTemp(1) As Boolean
Public MSFlag As Integer

Public dAct01CurrBuf(MAX_INT) As Double
Public dAct02CurrBuf(MAX_INT) As Double
Public dAct03CurrBuf(MAX_INT) As Double
Public dAct04CurrBuf(MAX_INT) As Double
Public dBlowerCurrBuf(MAX_INT) As Double
Public dVibBuf(2, MAX_INT) As Double

Public PlcData(20) As Long

Public dSupplyVolt As Double ' Read from Serial
Public dSupplyCurr As Double ' Read from Serial

Public lpBmpFileName As String
Public nRetestCount As Integer

Public bDisp As Boolean
Public nDispCounter As Integer

Public nLeakDummyPos As Integer
Public bLeakDummyUse As Boolean
Public bReleaseLeak As Boolean

Public lpScannerData As String
Public nScannerResult As Integer
Public nScannerCount As Integer

Public lpVisionCom As String
Public lpNvhCom As String

Public lpSteppingData(1) As String
Public bActStallBit(3) As Boolean
Public bIOActRotation(3) As Boolean

Public nSteppingMode As Integer

Public lpReadStepRotation(3) As String
Public lpReadStepData(7) As String

Public bSteppingWrite(3) As Boolean
Public nSteppingMoveBit(1) As Integer

Public lpWriteStepCmd(1) As String ' PORT CMD
Public lpWriteStepBit(1) As String ' PORT STALL, ACT ID
Public lpWriteStepRotation(3) As String ' ACT CW, CCW
Public lpWriteStepData(7) As String

Public dOldStepData(3) As Double

Public bMSTestStart As Boolean
Public nTime(1) As Integer

Public bAutoLinRead As Boolean
Public nLinAutoAddressFlag As Integer
Public nLinActNo As Integer
Public nLinReadSeq(9) As Integer
Public nLinCount(9, 9) As Integer
Public nLinTimeNo As Integer
Public bLinCommError As Boolean
Public bLinErrorResult(9) As Integer

Public bLinAct01Detect As Boolean

Public bScreenLoad As Boolean

Public nActFirstMovePos(9) As Integer

Public lStepActData(3) As Long
