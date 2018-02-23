Attribute VB_Name = "mdlHardCtrl"
Option Explicit


' =============================================================================
' P1202 API & CONST
' -----------------------------------------------------------------------------
Private Declare Function P1202_DriverInit Lib "p1202.dll" (wTotalBoards As Integer) As Integer
Private Declare Sub P1202_DriverClose Lib "p1202.dll" ()

Private Declare Function P1202_GetDriverVersion Lib "p1202.dll" (wVxdVersion As Integer) As Integer
Private Declare Function P1202_GetConfigAddressSpace Lib "p1202.dll" (ByVal wBoardNo As Integer, wAddrTimer As Integer, wAddrCtrl As Integer, wAddrDio As Integer, wAddrAdda As Integer) As Integer

Private Declare Function P1202_ActiveBoard Lib "p1202.dll" (ByVal wBoardNo As Integer) As Integer
Private Declare Function P1202_WhichBoardActive Lib "p1202.dll" () As Integer

Private Declare Function P1202_M_FUN_1 Lib "p1202.dll" (ByVal wDaFrequency As Integer, ByVal wDaWave As Integer, ByVal fDaAmplitude As Single, ByVal wAdClock As Integer, ByVal wAdNumber As Integer, ByVal wAdConfig As Integer, fAdBuf As Single, ByVal fLowAlarm As Single, ByVal fHighAlarm As Single) As Integer
Private Declare Function P1202_M_FUN_2 Lib "p1202.dll" (ByVal wDaNumber As Integer, ByVal wDaWave As Integer, wDaBuf As Integer, ByVal wAdClock As Integer, ByVal wAdNumber As Integer, ByVal wAdConfig As Integer, wAdBuf As Integer) As Integer
Private Declare Function P1202_M_FUN_3 Lib "p1202.dll" (ByVal wDaFrequency As Integer, ByVal wDaWave As Integer, ByVal fDaAmplitude As Single, ByVal wAdClock As Integer, ByVal wAdNumber As Integer, wChannelStatus As Integer, wAdConfig As Integer, fAdBuf As Single, ByVal fLowAlarm As Single, ByVal fHighAlarm As Single) As Integer
Private Declare Function P1202_M_FUN_4 Lib "p1202.dll" (ByVal wType As Integer, ByVal wDaFrequency As Integer, ByVal wDaWave As Integer, ByVal fDaAmplitude As Single, ByVal wAdClock As Integer, ByVal wAdNumber As Integer, wChannelStatus As Integer, wAdConfig As Integer, fAdBuf As Single, ByVal fLowAlarm As Single, ByVal fHighAlarm As Single) As Integer

Private Declare Function P1202_Di Lib "p1202.dll" (wDi As Integer) As Integer
Private Declare Function P1202_Do Lib "p1202.dll" (ByVal wDO As Integer) As Integer

Private Declare Function P1202_Da Lib "p1202.dll" (ByVal wDaChannel As Integer, ByVal wDaVal As Integer) As Integer
Private Declare Function P1202_SetChannelConfig Lib "p1202.dll" (ByVal wAdChannel As Integer, ByVal wConfig As Integer) As Integer

Private Declare Function P1202_AdPolling Lib "p1202.dll" (fAdVal As Single) As Integer
Private Declare Function P1202_AdsPolling Lib "p1202.dll" (fAdVal As Single, ByVal wNum As Integer) As Integer
Private Declare Function P1202_AdsPacer Lib "p1202.dll" (fAdVal As Single, ByVal dwNum As Long, ByVal wSample As Integer) As Integer

Private Declare Function P1202_ClearScan Lib "p1202.dll" () As Integer
Private Declare Function P1202_StartScan Lib "p1202.dll" (ByVal wSampleRate As Integer, ByVal dwNum As Long, ByVal nPriority As Integer) As Integer
Private Declare Sub P1202_ReadScanStatus Lib "p1202.dll" (wStatus As Integer, dwLowAlarm As Long, dwHighAlarm As Long)
Private Declare Function P1202_AddToScan Lib "p1202.dll" (ByVal wAdChannel As Integer, ByVal wConfig As Integer, ByVal wAverage As Integer, ByVal wLowAlarm As Integer, ByVal wHighAlarm As Integer, ByVal wAlarmType As Integer) As Integer
Private Declare Function P1202_SaveScan Lib "p1202.dll" (ByVal wOridinalOrder As Integer, wBuf As Integer) As Integer
Private Declare Sub P1202_WaitMagicScanFinish Lib "p1202.dll" (wStatus As Integer, wLowAlarm As Integer, wHighAlarm As Integer)
Private Declare Function P1202_StopMagicScan Lib "p1202.dll" () As Integer

Private Declare Function P1202_StartScanPostTrg Lib "p1202.dll" (ByVal wSampleRateDiv As Integer, ByVal dwNum As Long, ByVal nPriority As Integer) As Integer
Private Declare Function P1202_StartScanPreTrg Lib "p1202.dll" (ByVal wSampleRateDiv As Integer, ByVal dwNum As Long, ByVal nPriority As Integer) As Integer
Private Declare Function P1202_StartScanMiddleTrg Lib "p1202.dll" (ByVal wSampleRateDiv As Integer, ByVal dwN1 As Long, ByVal dwN2 As Long, ByVal nPriority As Integer) As Integer

Private Declare Function P1202_DelayUs Lib "p1202.dll" (ByVal wDelayUs As Integer) As Integer

Private Declare Function P1202_FunB_Start Lib "p1202.dll" (ByVal wClockDiv As Integer, wChannel As Integer, wConfig As Integer, Buffer As Integer, ByVal dwMaxCount As Long, ByVal nPriority As Integer) As Integer
Private Declare Function P1202_FunB_ReadStatus Lib "p1202.dll" () As Integer
Private Declare Function P1202_FunB_Stop Lib "p1202.dll" () As Integer
Private Declare Function P1202_FunB_Get Lib "p1202.dll" (P0 As Long) As Integer

Private Declare Function P1202_Card0_StartScan Lib "p1202.dll" (ByVal wSampleRate As Integer, wChannelStatus As Integer, wChannelConfig As Integer, ByVal wCount As Integer) As Integer
Private Declare Function P1202_Card0_ReadStatus Lib "p1202.dll" (wBuf As Integer, wBuf2 As Integer, dwP1 As Long, dwP2 As Long, wStatus As Integer) As Integer
Private Declare Sub P1202_Card0_Stop Lib "p1202.dll" ()

Private Declare Function P1202_Card1_StartScan Lib "p1202.dll" (ByVal wSampleRate As Integer, wChannelStatus As Integer, wChannelConfig As Integer, ByVal wCount As Integer) As Integer
Private Declare Function P1202_Card1_ReadStatus Lib "p1202.dll" (wBuf As Integer, wBuf2 As Integer, dwP1 As Long, dwP2 As Long, wStatus As Integer) As Integer
Private Declare Sub P1202_Card1_Stop Lib "p1202.dll" ()

Private Const NoError               As Integer = 0
Private Const DriverHandleError     As Integer = 1
Private Const DriverCallError       As Integer = 2
Private Const AdControllerError     As Integer = 3
Private Const M_FunExecError        As Integer = 4
Private Const ConfigCodeError       As Integer = 5
Private Const FrequencyComputeError As Integer = 6
Private Const HighAlarm             As Integer = 7
Private Const LowAlarm              As Integer = 8
Private Const AdPollingTimeOut      As Integer = 9
Private Const AlarmTypeError        As Integer = 10
Private Const FindBoardError        As Integer = 11
Private Const AdChannelError        As Integer = 12
Private Const DaChannelError        As Integer = 13
Private Const InvalidateDelay       As Integer = 14
Private Const DelayTimeOut          As Integer = 15
Private Const InvalidateData        As Integer = 16
Private Const FifoOverflow          As Integer = 17
Private Const Timeout               As Integer = 18
Private Const ExceedBoardNumber     As Integer = 19
Private Const NotFoundBoard         As Integer = 20
Private Const OpenError             As Integer = 21
Private Const FindTwoBoardError     As Integer = 22
Private Const ThreadCreateError     As Integer = 23
Private Const StopError             As Integer = 24
Private Const AllocateMemoryError   As Integer = 25


' =============================================================================
' PIO-DA API
' -----------------------------------------------------------------------------
' The Test functions
Declare Function PIODA_ShortSub Lib "PIODA.dll" (ByVal A As Integer, ByVal B As Integer) As Integer
Declare Function PIODA_FloatSub Lib "PIODA.dll" (ByVal A As Single, ByVal B As Single) As Single
Declare Function PIODA_GetDllVersion Lib "PIODA.dll" () As Integer

' The Driver functions
Declare Function PIODA_DriverInit Lib "PIODA.dll" () As Integer
Declare Sub PIODA_DriverClose Lib "PIODA.dll" ()
Declare Function PIODA_SearchCard Lib "PIODA.dll" (wBoards As Integer, ByVal dwPIOPISOCardID As Long) As Integer
Declare Function PIODA_GetDriverVersion Lib "PIODA.dll" (wDriverVersion As Integer) As Integer
Declare Function PIODA_GetConfigAddressSpace Lib "PIODA.dll" (ByVal wBoardNo As Integer, wAddrBase As Long, wIrqNo As Integer, wSubVendor As Integer, wSubDevice As Integer, wSubAux As Integer, wSlotBus As Integer, wSlotDevice As Integer) As Integer

Declare Function PIODA_ActiveBoard Lib "PIODA.dll" (ByVal wBoardNo As Integer) As Integer
Declare Function PIODA_WhichBoardActive Lib "PIODA.dll" () As Integer
Declare Function PIODA_SetCounter Lib "PIODA.dll" (ByVal wBoardNo As Integer, ByVal wWhichCounter As Integer, ByVal bConfig As Integer, ByVal wValue As Long) As Long
Declare Function PIODA_GetBaseAddress Lib "PIODA.dll" (ByVal wBoardNo As Integer) As Long

' EEPROM functions
Declare Function PIODA_EEP_READ Lib "PIODA.dll" (ByVal wBoardNo As Integer, ByVal wOffset As Integer, bHi As Integer, bLo As Integer) As Integer
Declare Function PIODA_EEP_WR_EN Lib "PIODA.dll" (ByVal wBoardNo As Integer) As Integer
Declare Function PIODA_EEP_WR_DIS Lib "PIODA.dll" (ByVal wBoardNo As Integer) As Integer
Declare Function PIODA_EEP_WRITE Lib "PIODA.dll" (ByVal wBoardNo As Integer, ByVal wOffset As Integer, ByVal HI As Integer, ByVal LO As Integer) As Integer

' DA functions
Declare Function PIODA_Voltage Lib "PIODA.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal fValue As Single) As Integer
Declare Function PIODA_Current Lib "PIODA.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal fValue As Single) As Integer
Declare Function PIODA_CalVoltage Lib "PIODA.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal fValue As Single) As Integer
Declare Function PIODA_CalCurrent Lib "PIODA.dll" (ByVal wBoardNo As Integer, ByVal wChannel As Integer, ByVal fValue As Single) As Integer

' DIO functions
Declare Sub PIODA_OutputByte Lib "PIODA.dll" (ByVal wBaseAddress As Long, ByVal dataout As Integer)
Declare Sub PIODA_OutputWord Lib "PIODA.dll" (ByVal wBaseAddress As Long, ByVal dataout As Long)
Declare Function PIODA_InputByte Lib "PIODA.dll" (ByVal wBaseAddress As Long) As Integer
Declare Function PIODA_InputWord Lib "PIODA.dll" (ByVal wBaseAddress As Long) As Long
Declare Function PIODA_DI Lib "PIODA.dll" (ByVal wBoardNo As Integer, wVal As Long) As Integer
Declare Function PIODA_DO Lib "PIODA.dll" (ByVal wBoardNo As Integer, ByVal wDO As Long) As Integer

' Interrupt functions
Declare Function PIODA_IntInstall Lib "PIODA.dll" (ByVal wBoard As Integer, hEvent As Long, ByVal wInterruptSource As Integer, ByVal wActiveMode As Integer) As Integer
Declare Function PIODA_IntRemove Lib "PIODA.dll" () As Integer
Declare Function PIODA_IntResetCount Lib "PIODA.dll" () As Integer
Declare Function PIODA_IntGetCount Lib "PIODA.dll" (dwIntCount As Long) As Integer

Private Const PIODA_NoError                 As Integer = 0
Private Const PIODA_DriverOpenError         As Integer = 1
Private Const PIODA_DriverNoOpen            As Integer = 2
Private Const PIODA_GetDriverVersionError   As Integer = 3
Private Const PIODA_InstallIrqError         As Integer = 4
Private Const PIODA_ClearIntCountError      As Integer = 5
Private Const PIODA_GetIntCountError        As Integer = 6
Private Const PIODA_RegisterApcError        As Integer = 7
Private Const PIODA_RemoveIrqError          As Integer = 8
Private Const PIODA_FindBoardError          As Integer = 9
Private Const PIODA_ExceedBoardNumber       As Integer = 10
Private Const PIODA_ResetError              As Integer = 11

Private Const PIODA_EEPROMDataError         As Integer = 12
Private Const PIODA_EEPROMWriteError        As Integer = 13

' to trigger a interrupt when high -> low
Private Const PIODA_ActiveLow               As Integer = 0
' to trigger a interrupt when low -> high
Private Const PIODA_ActiveHigh              As Integer = 1

' ID
Private Const PIO_DA                        As Long = &H800400        ' PIO-DA16/DA8/DA4

Private DaErrMsg As String


' =============================================================================
' DLPORTIO API
' -----------------------------------------------------------------------------

Public Declare Function DlPortReadPortUchar Lib "dlportio.dll" (ByVal Port As Long) As Byte
Public Declare Function DlPortReadPortUshort Lib "dlportio.dll" (ByVal Port As Long) As Integer
Public Declare Function DlPortReadPortUlong Lib "dlportio.dll" (ByVal Port As Long) As Long

Public Declare Sub DlPortReadPortBufferUchar Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortReadPortBufferUshort Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortReadPortBufferUlong Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)

Public Declare Sub DlPortWritePortUchar Lib "dlportio.dll" (ByVal Port As Long, ByVal Value As Byte)
Public Declare Sub DlPortWritePortUshort Lib "dlportio.dll" (ByVal Port As Long, ByVal Value As Integer)
Public Declare Sub DlPortWritePortUlong Lib "dlportio.dll" (ByVal Port As Long, ByVal Value As Long)

Public Declare Sub DlPortWritePortBufferUchar Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortWritePortBufferUshort Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortWritePortBufferUlong Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)



' =============================================================================
' PISODIO API
' -----------------------------------------------------------------------------
Private Const PISODIO_NoError               As Integer = 0
Private Const PISODIO_DriverOpenError       As Integer = 1
Private Const PISODIO_DriverNoOpen          As Integer = 2
Private Const PISODIO_GetDriverVersionError As Integer = 3
Private Const PISODIO_InstallIrqError       As Integer = 4
Private Const PISODIO_ClearIntCountError    As Integer = 5
Private Const PISODIO_GetIntCountError      As Integer = 6
Private Const PISODIO_RegisterApcError      As Integer = 7
Private Const PISODIO_RemoveIrqError        As Integer = 8
Private Const PISODIO_FindBoardError        As Integer = 9
Private Const PISODIO_ExceedBoardNumber     As Integer = 10
Private Const PISODIO_ResetError            As Integer = 11

' to trigger a interrupt when high -> low
Private Const PISODIO_ActiveLow             As Integer = 0
' to trigger a interrupt when low -> high
Private Const PISODIO_ActiveHigh            As Integer = 1

' ID
Private Const PISO_C64      As Long = &H800800                      ' for PISO-C64
Private Const PISO_P64      As Long = &H800810                      ' for PISO-P64
Private Const PISO_A64      As Long = &H800850                      ' for PISO-A64
Private Const PISO_P32C32   As Long = &H800820                      ' for PISO-P32C32
Private Const PISO_P32A32   As Long = &H800870                      ' for PISO-P32A32
Private Const PISO_P8R8     As Long = &H800830                      ' for PISO-P8R8
Private Const PISO_P8SSR8AC As Long = &H800830                      ' for PISO-P8SSR8AC
Private Const PISO_P8SSR8DC As Long = &H800830                      ' for PISO-P8SSR8DC
Private Const PISO_730      As Long = &H800840                      ' for PISO-730
Private Const PISO_730A     As Long = &H800880                      ' for PISO-730A

' The Test functions
Private Declare Function PISODIO_ShortSub Lib "PISODIO.dll" (ByVal A As Integer, ByVal B As Integer) As Integer
Private Declare Function PISODIO_FloatSub Lib "PISODIO.dll" (ByVal A As Single, ByVal B As Single) As Single
Private Declare Function PISODIO_GetDllVersion Lib "PISODIO.dll" () As Integer

' The Driver functions
Private Declare Function PISODIO_DriverInit Lib "PISODIO.dll" () As Integer
Private Declare Sub PISODIO_DriverClose Lib "PISODIO.dll" ()
Private Declare Function PISODIO_SearchCard Lib "PISODIO.dll" (wBoards As Integer, ByVal dwPIOPISOCardID As Long) As Integer
Private Declare Function PISODIO_GetDriverVersion Lib "PISODIO.dll" (wDriverVersion As Integer) As Integer
Private Declare Function PISODIO_GetConfigAddressSpace Lib "PISODIO.dll" (ByVal wBoardNo As Integer, wAddrBase As Long, wIrqNo As Integer, wSubVendor As Integer, wSubDevice As Integer, wSubAux As Integer, wSlotBus As Integer, wSlotDevice As Integer) As Integer
Private Declare Function PISODIO_ActiveBoard Lib "PISODIO.dll" (ByVal wBoardNo As Integer) As Integer
Private Declare Function PISODIO_WhichBoardActive Lib "PISODIO.dll" () As Integer

' DIO functions
Private Declare Sub PISODIO_OutputByte Lib "PISODIO.dll" (ByVal address As Long, ByVal dataout As Integer)
Private Declare Sub PISODIO_OutputWord Lib "PISODIO.dll" (ByVal address As Long, ByVal dataout As Long)
Private Declare Function PISODIO_InputByte Lib "PISODIO.dll" (ByVal address As Long) As Integer
Private Declare Function PISODIO_InputWord Lib "PISODIO.dll" (ByVal address As Long) As Long

' Interrupt functions
Private Declare Function PISODIO_IntInstall Lib "PISODIO.dll" (ByVal wBoard As Integer, hEvent As Long, ByVal wInterruptSource As Integer, ByVal wActiveMode As Integer) As Integer
Private Declare Function PISODIO_IntRemove Lib "PISODIO.dll" () As Integer
Private Declare Function PISODIO_IntGetCount Lib "PISODIO.dll" (dwIntCount As Long) As Integer
Private Declare Function PISODIO_IntResetCount Lib "PISODIO.dll" () As Integer


' =============================================================================
' API
' -----------------------------------------------------------------------------

Declare Function GetTickCount Lib "kernel32" () As Long


' =============================================================================
' PRINTER API
' -----------------------------------------------------------------------------

Private Type DOCINFO
    lpDocName       As String
    lpOutputFile    As String
    lpDatatype      As String
End Type

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long


' =============================================================================
' START
' -----------------------------------------------------------------------------

Private Const MAX_DIO       As Integer = 11

' AD (P1202)
Private nADTotalBoards      As Integer
Private nADRes              As Integer

Private Channel(32)         As Integer
Private ConfigCode(32)      As Integer
Private AdNumber            As Integer
Private AdErrMsg            As String

' DIO (PISODIO)
Private wBaseAddr(1)        As Long

Private wIrq                As Integer
Private wSubVendor          As Integer
Private wSubDevice          As Integer
Private wSubAux             As Integer
Private wSlotBus            As Integer
Private wSlotDevice         As Integer
Private wTotalBoards
Private wInitialCode        As Integer

Private nStatusO(MAX_DIO)   As Integer
Private lAddrO(MAX_DIO)     As Long
Private lAddrI(MAX_DIO)     As Long
Private bBoardEnable(9) As Boolean

Public Type TagPowerSerialVariables
    iCOM                As String
    bVolt               As Boolean
    bCurr               As Boolean
    nCount              As Integer
End Type

Public PowerVar        As TagPowerSerialVariables

Public Function InitAD() As Boolean
    Dim bRes As Boolean
    
    bRes = True
    
    If IS_CARD Then
        nADRes = P1202_DriverInit(nADTotalBoards)
        If nADRes <> NoError Then
            Call MsgBox("The Return Error Code = " + str$(nADRes) + vbCrLf + "The 1202 Card Not Found !" + vbCrLf + "Function : InitAD(). #1", vbOKOnly, "P1202 Return Error Code !")
            bRes = False
        End If
        
        nADRes = P1202_ActiveBoard(0)
        If nADRes <> NoError Then
            Call MsgBox("The Return Error Code = " + str$(nADRes) + vbCrLf + "The 180X Card Not Found !", 0, "P1202 Return Error Code !")
            bRes = False
        End If
        
        Call P1202_DelayUs(23)                  ' delay 23 us settling time
    End If
    
    InitAD = bRes
End Function

Private Static Function GetAdPolling(ByVal nCh As Integer) As Double
    Static fData As Single
    
    If IS_CARD Then
        Call P1202_SetChannelConfig(nCh, 0)      ' +/- 5V range
        Call P1202_AdPolling(fData)
    End If
    
    GetAdPolling = CDbl(fData)
End Function

Public Static Function ADRead(ByVal nCh As Integer) As Double
    Dim dValue As Double
    Dim dAvg As Double
    Dim i As Integer
    Dim nFiltering As Integer
    Dim bZero As Boolean
    Dim dNaive As Double
    Dim bMinus As Boolean
    Dim dAdd As Double
    Dim dMulti As Double
    Dim bSetupAdjustUse As Boolean
    Dim dSetupAdd As Double
    Dim dSetupMulti As Double
    Dim dZeroCurr As Double
    Dim bPercent As Boolean
    
    If nCh = 999 Then
        ADRead = 99999
        
        Exit Function
    End If
    
    nFiltering = SysVar.nFiltering
    
    If nFiltering = 0 Then nFiltering = 1
    
    bZero = SysVar.bZero(nCh)
    dNaive = SysVar.dNaive(nCh)
    bMinus = SysVar.bMinus(nCh)
    dAdd = SysVar.dAdd(nCh)
    dMulti = SysVar.dMulti(nCh)
    bSetupAdjustUse = SetupVar.bAdjustUse(nCh)
    dSetupAdd = SetupVar.dAdd(nCh)
    dSetupMulti = SetupVar.dMulti(nCh)
    dZeroCurr = ZeroCurr(nCh)
    bPercent = SysVar.bPercent(nCh)
    
    dAvg = 0
    
    If IS_CARD Then
        For i = 0 To nFiltering - 1
            dAvg = dAvg + GetAdPolling(nCh)
        Next
        
        dAvg = dAvg / CDbl(nFiltering)
    Else
        If LOCALTEST Then
            If dNaive > 100 Then
                If dValue < 5 Then
                    dValue = dValue + 0.00001
                End If
            Else
                dValue = dValue + 0.0001
            End If
        Else
            dValue = CDbl((10 * Rnd) + 0)
        End If
        
        dAvg = dAvg + dValue
    End If
    
    If bZero Then
        dAvg = dAvg - dZeroCurr
    End If
    
    dAvg = (((dAvg * dNaive) + dAdd) * dMulti)
    
    If Not bMinus Then
        dAvg = Abs(dAvg)
    End If
    
    If bSetupAdjustUse Then
        dAvg = (dAvg + dSetupAdd) * dSetupMulti
    End If
    
    If bPercent Then
        dAvg = dAvg / 5 * 100
    End If
    
    ADRead = dAvg
End Function

Public Sub ClearAD()
    Erase ZeroCurr
End Sub

Public Sub ZeroAD()
    Dim dAvg    As Double
    Dim i       As Integer
    Dim k       As Integer

    ' DO 등의 컨트롤 사용시 이 시간이 존재하지 않으면 잘못된 Zero AD 가 될 수 있다.
    Call Sleep(100)
    
    Call ClearAD

    If IS_CARD Then
        For i = 0 To MAX_AD_CHANNEL
            dAvg = 0
            'For k = 0 To SysVar.nFiltering - 1
            For k = 0 To 49
                dAvg = dAvg + GetAdPolling(i)
            Next
            dAvg = dAvg / 50
            
            ZeroCurr(i) = dAvg
        Next
    End If
End Sub

Public Function OutDa(ByVal nCh As Integer, ByVal fValue As Single, Optional ByVal bPercent As Boolean = False)
    Dim dOutput As Double
    
    If IS_CARD = False Then Exit Function
    
    Select Case nCh
        Case 999:
            Exit Function
        
        Case DA_SUPPLY:
            Select Case nCh
                Case DA_SUPPLY: dOutput = (fValue / 10) / 0.0048828
            End Select
            
            dOutput = Val("&h" + Hex(dOutput + &H800))
            If dOutput < &H800 Then dOutput = &H800
            If dOutput > &HFFF Then dOutput = &HFFF
            Call P1202_Da(nCh, CInt(dOutput))
        
        Case Else:
            If bPercent Then
                fValue = fValue * 5 / 100
            End If
            
            Call PIODA_CalVoltage(0, nCh - 2, fValue)
    
    End Select
End Function

Public Function SetVolt(ByVal dVolt As Double)
    Call Sleep(100)
    
    Select Case POWERTYPE
        Case 0:
            Call OutDa(DA_SUPPLY, dVolt)
        
        Case 1:
            If frmMain.Power.PortOpen Then
                frmMain.Power.Output = "PV " + Format(dVolt, "#0.00") + vbCrLf
                frmMain.Power.Output = "RMT REM" + vbCrLf
            End If
        
        Case 2:
            If frmMain.Power.PortOpen Then
                frmMain.Power.Output = ":VOL" & Format(dVolt, "#00.000") & ";"
            End If
    
    End Select
    
    Call Sleep(100)
End Function

Public Function SetIo()
    Dim i As Integer
    Dim rtn
    Dim wRetVal As Integer
    
    '********************************************************************
    '* NOTICE: call PISODIO_DriverInit() to initialize the driver.        *
    '* Initial the device driver, and return the board number in the PC *
    '********************************************************************
    wInitialCode = PISODIO_DriverInit()
    
    If wInitialCode <> PISODIO_NoError Then
        rtn = MsgBox("Driver initialize error!!!", , "PISODIO Card Error")
        Exit Function
    End If
    
    If PISODIO_SearchCard(wTotalBoards, PISO_P32C32) <> PISODIO_NoError Then
        rtn = MsgBox("Search Card Error!!", , "PISODIO Card Error")
    End If
    
    If wTotalBoards < MAX_DIO_CARD - 1 Then
        rtn = MsgBox("I/O CARD NOT DETECT ERROR!!!", , "PISODIO Card Error")
        Exit Function
    End If
    
    Erase bBoardEnable
    
    For i = 0 To MAX_DIO_CARD - 1
        'Get board's Configuration Space
        wRetVal = PISODIO_GetConfigAddressSpace(i, wBaseAddr(i), wIrq, wSubVendor, wSubDevice, wSubAux, wSlotBus, wSlotDevice)
        
        ' enable DI/DO
        PISODIO_OutputByte wBaseAddr(i), 1
        
        bBoardEnable(i) = True
    Next
End Function

Public Sub InitMapping()
    Dim IAddr   As Variant
    Dim OAddr   As Variant
    Dim nLoop   As Integer
    Dim i       As Integer

    Erase lAddrI
    Erase lAddrO
    Erase nStatusO
    
    If CARD_TYPE Then
        IAddr = 0
        OAddr = 0
                     ' Board1                    ' Board2                    ' Board3
        IAddr = Array(&H180, &H181, &H182, &H183, &H188, &H189, &H18A, &H18B, &H190, &H191, &H192, &H193)
        OAddr = Array(&H184, &H185, &H186, &H187, &H18C, &H18D, &H18E, &H18F, &H194, &H195, &H196, &H197)
    
        If IS_CARD Then
            For nLoop = 0 To UBound(IAddr)
                lAddrI(nLoop) = IAddr(nLoop)
                lAddrO(nLoop) = OAddr(nLoop)
            Next
    
    '        For i = 0 To IO_MAX
    '            Call DlPortWritePortUchar(O_Addr(i), 0)
    '        Next
        End If
    Else
        If IS_CARD Then
            Call SetIo
        End If
        
        For i = 0 To UBound(bBoardEnable)
            If bBoardEnable(i) Then
                lAddrO(0 + (i * 4)) = wBaseAddr(i) + &HC0
                lAddrO(1 + (i * 4)) = wBaseAddr(i) + &HC4
                lAddrO(2 + (i * 4)) = wBaseAddr(i) + &HC8
                lAddrO(3 + (i * 4)) = wBaseAddr(i) + &HCC
                
                lAddrI(0 + (i * 4)) = wBaseAddr(i) + &HC0
                lAddrI(1 + (i * 4)) = wBaseAddr(i) + &HC4
                lAddrI(2 + (i * 4)) = wBaseAddr(i) + &HC8
                lAddrI(3 + (i * 4)) = wBaseAddr(i) + &HCC
            End If
        Next
        
        If IS_CARD Then
            For i = 0 To ((MAX_DIO_CARD * 2) * 2) - 1
                Call PISODIO_OutputByte(lAddrO(i), 0)
            Next
        End If
    End If
End Sub

Public Sub DO_Control(ByVal nCh As Integer, ByVal bOutput As Boolean)
    Dim nNo     As Integer
    Dim nStatus As Integer
    
    If nCh = 999 Then Exit Sub
    
    If CARD_TYPE Then
        nNo = Int(nCh / 10)
        nStatus = 2 ^ (nCh Mod 10)
        
        If bOutput = False Then
            nStatus = Not nStatus
        End If
        
        If bOutput = False Then ' OFF
            nStatusO(nNo) = nStatusO(nNo) And nStatus
        Else                    ' ON
            nStatusO(nNo) = nStatusO(nNo) Or nStatus
        End If
        
        If IS_CARD Then
            Call DlPortWritePortUchar(lAddrO(nNo), nStatusO(nNo))
        End If
    Else
        nNo = Int(nCh / 10)
        nStatus = 2 ^ (nCh Mod 10)
        
        If bOutput = False Then
            nStatus = (Not nStatus) And &HFF
        End If
        
        If bOutput = False Then
            nStatusO(nNo) = nStatusO(nNo) And nStatus
        Else
            nStatusO(nNo) = nStatusO(nNo) Or nStatus
        End If
        
        If IS_CARD Then
            Call PISODIO_OutputByte(lAddrO(nNo), nStatusO(nNo))
        End If
    End If
End Sub

Public Function DOS(ByVal nCh As Integer) As Boolean
    Dim nNo     As Integer
    Dim nStatus As Integer
    Dim bRes    As Boolean
    
    If nCh = 999 Then Exit Function
    
    ' PISODIO 와 DLPORTIO 가 같다.
    
    nNo = Int(nCh / 10)
    nStatus = 2 ^ (nCh Mod 10)
    
    If (nStatusO(nNo) And nStatus) = 0 Then
        bRes = False
    Else
        bRes = True
    End If
    
    DOS = bRes
End Function

Public Function DIS(ByVal nCh As Integer) As Boolean
    Dim nNo     As Integer
    Dim nStatus As Integer
    Dim byGet   As Byte
    Dim bRes    As Boolean
    
    If nCh = 999 Then Exit Function
    
    nNo = Int(nCh / 10)
    nStatus = 2 ^ (nCh Mod 10)
    
    If LOCALTEST Then
        Select Case nCh
            Case I_AUTO_SW: bRes = True
        End Select
    Else
        If CARD_TYPE Then
            If IS_CARD Then
                byGet = DlPortReadPortUchar(lAddrI(nNo))
            End If
            
            If (byGet And nStatus) = 0 Then
                bRes = False
            Else
                bRes = True
            End If
        Else
            If IS_CARD Then
                byGet = PISODIO_InputByte(lAddrI(nNo))
            End If
            
            If (byGet And nStatus) = 0 Then
                bRes = True
            Else
                bRes = False
            End If
        End If
        
        If IS_CARD = False Then
            If nCh = 2 Then
                bRes = Not bRes
            End If
        End If
    End If
    
    DIS = bRes
End Function

Public Sub AD_Close()
    If IS_CARD Then
        Call P1202_DriverClose
    End If
End Sub

Public Sub DO_Clear()
    ' All Output False
    If DOS(O_START_LAMP) Then Call DO_Control(O_START_LAMP, False)
    If DOS(O_STOP_LAMP) Then Call DO_Control(O_STOP_LAMP, False)
    If DOS(O_OK_LAMP) Then Call DO_Control(O_OK_LAMP, False)
    If DOS(O_NG_LAMP) Then Call DO_Control(O_NG_LAMP, False)
    If DOS(O_RUN_LAMP) Then Call DO_Control(O_RUN_LAMP, False)
    If DOS(O_BUZZER) Then Call DO_Control(O_BUZZER, False)
    
    If DOS(O_MSOK) Then Call DO_Control(O_MSOK, False)
    If DOS(O_MSNG) Then Call DO_Control(O_MSNG, False)
    
    If DOS(O_BLOWER_POWER) Then Call DO_Control(O_BLOWER_POWER, False)
    If DOS(O_BLOWER_01) Then Call DO_Control(O_BLOWER_01, False)
    If DOS(O_BLOWER_02) Then Call DO_Control(O_BLOWER_02, False)
    If DOS(O_BLOWER_03) Then Call DO_Control(O_BLOWER_03, False)
    If DOS(O_BLOWER_04) Then Call DO_Control(O_BLOWER_04, False)
    If DOS(O_BLOWER_05) Then Call DO_Control(O_BLOWER_05, False)
    If DOS(O_BLOWER_06) Then Call DO_Control(O_BLOWER_06, False)
    If DOS(O_BLOWER_07) Then Call DO_Control(O_BLOWER_07, False)
    If DOS(O_BLOWER_08) Then Call DO_Control(O_BLOWER_08, False)
    If DOS(O_BLOWER_PWM) Then Call DO_Control(O_BLOWER_PWM, False)
    If DOS(O_BLOWER_DIRECTION) Then Call DO_Control(O_BLOWER_DIRECTION, False)
    
    If DOS(O_ACT01_POWER) Then Call DO_Control(O_ACT01_POWER, False)
    If DOS(O_ACT02_POWER) Then Call DO_Control(O_ACT02_POWER, False)
    If DOS(O_ACT03_POWER) Then Call DO_Control(O_ACT03_POWER, False)
    If DOS(O_ACT04_POWER) Then Call DO_Control(O_ACT04_POWER, False)
    If DOS(O_ACT05_POWER) Then Call DO_Control(O_ACT05_POWER, False)
    If DOS(O_ACT06_POWER) Then Call DO_Control(O_ACT06_POWER, False)
    If DOS(O_ACT07_POWER) Then Call DO_Control(O_ACT07_POWER, False)
    If DOS(O_ACT08_POWER) Then Call DO_Control(O_ACT08_POWER, False)
    If DOS(O_SENSOR_POWER) Then Call DO_Control(O_SENSOR_POWER, False)
    
    If SetupVar.nBlowerType = 2 Then
        Call LinPortClose(1)
    End If

'    If DOS(O_LEAK01_START) Then Call DO_Control(O_LEAK01_START, False)
'    If DOS(O_LEAK01_STOP) Then Call DO_Control(O_LEAK01_STOP, False)
'    If DOS(O_LEAK01_CAL) Then Call DO_Control(O_LEAK01_CAL, False)
    
'    If DOS(O_LEAK02_START) Then Call DO_Control(O_LEAK02_START, False)
'    If DOS(O_LEAK02_STOP) Then Call DO_Control(O_LEAK02_STOP, False)
'    If DOS(O_LEAK02_CAL) Then Call DO_Control(O_LEAK02_CAL, False)
    
'    If DOS(O_LEAK_MODE_ID1) Then Call DO_Control(O_LEAK_MODE_ID1, False)
'    If DOS(O_LEAK_MODE_ID2) Then Call DO_Control(O_LEAK_MODE_ID2, False)
'    If DOS(O_LEAK_MODE_ID4) Then Call DO_Control(O_LEAK_MODE_ID4, False)
'    If DOS(O_LEAK_MODE_ID8) Then Call DO_Control(O_LEAK_MODE_ID8, False)
    
'    If DOS(O_LEAK_CLAMP_SOL) Then Call DO_Control(O_LEAK_CLAMP_SOL, False)
End Sub

Public Function InitDA() As Boolean
    Dim wTotalBoards    As Integer
    Dim wInitialCode    As Integer
    Dim wRtn            As Integer
    Dim bRes            As Boolean
    
    bRes = True
    
    '********************************************************************
    '* NOTICE: call PIODA_DriverInit() to initialize the driver.        *
    '* Initial the device driver, and return the board number in the PC *
    '********************************************************************
    DaErrMsg = ""
    wTotalBoards = 1
    
    If IS_CARD Then
        wInitialCode = PIODA_DriverInit()
        If wInitialCode <> PIODA_NoError Then
            DaErrMsg = "DA-CARD Driver Open Error !!!"
            bRes = False
        End If
        
        If bRes Then
            wRtn = PIODA_SearchCard(wTotalBoards, PIO_DA)
            If (wRtn <> PIODA_NoError) Then
                DaErrMsg = "DA-CARD Search Card Error!!" + vbCrLf + "Error Code:" + str(wRtn)
                bRes = False
            End If
        End If
    End If
    
    InitDA = bRes
End Function

Public Function MarkingInterlock(ByVal Index As Integer) As Boolean
    MarkingInterlock = True
    
    Select Case Index
        Case O_MARKING1:
            If SetupVar.bMarkingUse(0) = False Then
                Call OnLog("NOT USED MARKING #1")
                
                MarkingInterlock = False
            End If
            
            If DIS(I_WORK_ON) = False And DIS(I_WORK_OFF) = False Then
                Call OnLog("WORK MOVE NOT COMPLATE...")
                
                MarkingInterlock = False
            End If
            
            If DOS(O_WORK_ON) And DIS(I_WORK_ON) = False And DIS(I_WORK_OFF) Then
                Call OnLog("MARKING INTERLOCK !!!")
                
                MarkingInterlock = False
            End If
            
            If DOS(O_WORK_OFF) And DIS(I_WORK_OFF) = False And DIS(I_WORK_ON) Then
                Call OnLog("MARKING INTERLOCK !!!")
                
                MarkingInterlock = False
            End If
        
        Case O_WORK_ON, O_WORK_OFF:
            If SetupVar.bMarkingUse(0) Then
                If DIS(I_MARKING1_OFF) = False Then
                    Call OnLog("MARKING INTERLOCK !!!")
                    
                    MarkingInterlock = False
                End If
            End If
        
        Case Else:
            MarkingInterlock = True
    
    End Select
End Function

