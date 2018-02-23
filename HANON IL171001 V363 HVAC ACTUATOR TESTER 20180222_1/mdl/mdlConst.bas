Attribute VB_Name = "mdlConst"
Option Explicit


Public Declare Function QueryPerformanceFrequencyAny Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Any) As Long
Public Declare Function QueryPerformanceCounterAny Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const LOCALTEST As Boolean = False
Public Const DEBUGMODE As Boolean = True
Public Const IS_CARD As Boolean = True
Public Const SETUPSELECT As Boolean = False ' setup select true : enable, false : disable
Public Const POWERTYPE As Integer = 1 ' 0 : ilho, 1 : big lamda, 2 : small lamda
Public Const POWERGET As Boolean = True
Public Const CARD_TYPE As Boolean = False   ' True : DLPORTIO, False : PISODIO
Public Const PLCUSE As Boolean = False
Public Const SCANNERUSE As Boolean = False
Public Const LEAKUSE As Boolean = False
Public Const TABLETYPE As Boolean = False
Public Const LINUSE As Boolean = False
Public Const LINDEBUG As Boolean = False
Public Const NVHUSE As Boolean = False
Public Const STEPUSE As Boolean = False
Public Const STEPLOGUSE As Boolean = False

Public Const NVHACT As String = "MODE: ACT"
Public Const NVHLOW As String = "MODE: LOW"
Public Const NVHHD As String = "MODE: H-D"
Public Const NVHHV As String = "MODE: H-V"
Public Const NVHSTOP As String = "MeasureA: 0"
Public Const NVHREMOVE As String = "Remove: "

Public Const STEP_MODE_READ As Integer = 1
Public Const STEP_MODE_MOVE As Integer = 2

Public Const STEP_START As String = "FF"
Public Const STEP_END As String = "FE"
Public Const STEP_MOVE As String = "4D"
Public Const STEP_READ As String = "52"
Public Const STEP_STOP As String = "50"
Public Const STEP_INFO As String = "45"

Public Const STEP_POS1 As String = "EE"
Public Const STEP_POS2 As String = "EE"

Public Const STEP_LITTLE1 As String = "01"
Public Const STEP_LITTLE2 As String = "2C"

Public Const STEP_ROTA_P1 As String = "00"
Public Const STEP_ROTA_P2 As String = "01"

Public Const STEP_NULL As String = "00"

Public Const CM_HI As Integer = 1
Public Const CM_LO As Integer = 0

' 0 start = total + 1
Public Const RANK_OR_GROUP As Integer = 0 ' 0 = rank, 1 = group ' ref = plc
Public Const CARSAVECOUNT As Integer = 10 ' total car name
Public Const SUBSAVECOUNT As Integer = 10
Public Const TOTALSAVECOUNT As Integer = CARSAVECOUNT * SUBSAVECOUNT

Public Const MASTERSAMPLELANG As Integer = 0 ' 0 - mastersample    1 - calibration

Public Const DISP_TIME As Integer = 10
Public Const MANUAL_INIT As Integer = -9
Public Const EMPTY_STACK_ADDR As Integer = -1
Public Const MAX_INT As Integer = 32767
Public Const LIN_1STEP_ANGLE As Double = 0.05625

Public Const CO_NONE As Long = &HE0E0E0
Public Const CO_BNONE As Long = &H808080
Public Const CO_CNONE As Long = &HF0F0F0
Public Const CO_GRASS As Long = &HC0C000
Public Const CO_ORANGE As Long = &H8080FF

' Step Positon
' 참고 : 0~99번대는 Run 시 사용하므로 쓰지 않는 것이 건강에 좋다.
Public Const POS_INIT As Integer = 1000
Public Const POS_INIT2 As Integer = 1010
Public Const POS_START_INIT As Integer = 1100
Public Const POS_START_RUN As Integer = 1200
Public Const POS_START_RUN2 As Integer = 1210
Public Const POS_RUN_INIT As Integer = 1300
Public Const POS_RUN As Integer = 1400
'Public Const POS_RUN_VISION As Integer = 1500
Public Const POS_END_INIT As Integer = 1600
Public Const POS_END_RUN As Integer = 1700
Public Const POS_AFTER_VIB_WAIT As Integer = 2000
Public Const POS_AFTER_VIB_INIT As Integer = 2100
Public Const POS_AFTER_VIB_RUN As Integer = 2200
Public Const POS_END As Integer = 3000
Public Const POS_BLOWER_HI As Integer = 8

' Test Positions
Public Const TP_TEST As Integer = 1
Public Const TP_BLOWER As Integer = 2
Public Const TP_ACT01 As Integer = 3
Public Const TP_ACT02 As Integer = 4
Public Const TP_ACT03 As Integer = 5
Public Const TP_ACT04 As Integer = 6
Public Const TP_SENSOR As Integer = 7
Public Const TP_ION As Integer = 9
Public Const TP_LINACT01 As Integer = 12
Public Const TP_LINACT02 As Integer = 13
Public Const TP_LINACT03 As Integer = 14
Public Const TP_LINACT04 As Integer = 15
Public Const TP_LINPTC As Integer = 16

Public Const MSG_STOP As Integer = 0
Public Const MSG_RUN As Integer = 1
Public Const MSG_READY As Integer = 2
Public Const MSG_OK As Integer = 3
Public Const MSG_NG As Integer = 4
Public Const MSG_CAL As Integer = 5
Public Const MSG_ERR As Integer = 6
Public Const MSG_OKPASS As Integer = 10
Public Const MSG_OKFAIL As Integer = 11
Public Const MSG_NGPASS As Integer = 12
Public Const MSG_NGFAIL As Integer = 13
Public Const MSG_SCANNER As Integer = 14
Public Const MSG_NGTABLE As Integer = 15

Public Const FM_RUN As Integer = 0
Public Const FM_SETUP As Integer = 1
Public Const FM_IO As Integer = 2
Public Const FM_SYSTEM As Integer = 3
Public Const FM_DATABASE As Integer = 4
Public Const FM_EXIT As Integer = 5
Public Const FM_SPLASH As Integer = 6
Public Const FM_MAIN As Integer = 7
Public Const FM_LOGIN As Integer = 8
Public Const FM_GRAPH As Integer = 9
Public Const FM_SETUPSELECT As Integer = 10
Public Const FM_PROGRAM As Integer = 11

Public Const TM_AUTO As Integer = 0
Public Const TM_WAIT As Integer = 1
Public Const TM_CAL As Integer = 2
Public Const TM_RUN As Integer = 3
Public Const TM_MS As Integer = 4
Public Const TM_TEMPDELAY As Integer = 5
Public Const TM_SUPPLY_POWER As Integer = 6
Public Const TM_SCANNER As Integer = 7
Public Const TM_SERIAL As Integer = 8

Public Const TM_LIN As Integer = 10
Public Const TM_LINACT01 As Integer = 11
Public Const TM_LINACT02 As Integer = 12
Public Const TM_LINACT03 As Integer = 13
Public Const TM_LINACT04 As Integer = 14
Public Const TM_LINBLOWERCHECK As Integer = 15
Public Const TM_AUTOADDR As Integer = 16
Public Const TM_LINMANUAL As Integer = 17
Public Const TM_TOTAL As Integer = 18

Public Const TM_BLOWER As Integer = 20
Public Const TM_ACT01 As Integer = 21
Public Const TM_ACT02 As Integer = 22
Public Const TM_ACT03 As Integer = 23
Public Const TM_ACT04 As Integer = 24
Public Const TM_ACT05 As Integer = 25
Public Const TM_ACT06 As Integer = 26
Public Const TM_ACT07 As Integer = 27
Public Const TM_ACT08 As Integer = 28

Public Const TM_SENSOR As Integer = 30
Public Const TM_VISION As Integer = 31
Public Const TM_VISION1 As Integer = 32
Public Const TM_VISION2 As Integer = 33
Public Const TM_VISION3 As Integer = 34
Public Const TM_VISION4 As Integer = 35
Public Const TM_ION As Integer = 36

Public Const TM_AD As Integer = 40
Public Const TM_BUZZER As Integer = 41
Public Const TM_PTC As Integer = 42
Public Const TM_SPLASH As Integer = 43
Public Const TM_WORKON As Integer = 44
Public Const TM_SOL As Integer = 45
Public Const TM_LEAKTEST As Integer = 46
Public Const TM_LEAKSOL As Integer = 47
Public Const TM_MARKING As Integer = 48
Public Const TM_PRODUCT As Integer = 49

Public Const TM_PLCPROC As Integer = 50
Public Const TM_LEAK01 As Integer = 51
Public Const TM_LEAK02 As Integer = 52

Public Const TM_STEP01 As Integer = 55
Public Const TM_STEP02 As Integer = 56
Public Const TM_STEP03 As Integer = 57
Public Const TM_STEP04 As Integer = 58

Public Const MAX_AD_CHANNEL As Integer = 14 ' 0 : SUPPLY
Public Const MAX_DIO_CARD As Integer = 1
Public Const MAX_DIO_CHANNEL As Integer = (40 * MAX_DIO_CARD) - 1 ' 카드 수량에 따라 조절.
Public Const TEMP_DELAY_TIME As Integer = 3

Public Const MASTER_PASSWORD As String = "ILHOENG1525"
Public Const INISETUPFILE As String = "\SetupFile\setup.ini"
Public Const INISYSTEMFILE As String = "\SetupFile\system.ini"
Public Const INILANGFILE As String = "\SetupFile\language.ini"
Public Const INIPLCFILE As String = "\SetupFile\plcaddress.ini"
Public Const INIMODELFILE As String = "\SetupFile\modelname.ini"

' 999 not use ( AD channel is order array )
Public Const AD_SUPPLY_VOLT As Integer = 0
Public Const AD_ACT01_CURR As Integer = 1
Public Const AD_ACT01_VOLT As Integer = 2
Public Const AD_ACT02_CURR As Integer = 3
Public Const AD_ACT02_VOLT As Integer = 4
Public Const AD_ACT03_CURR As Integer = 5
Public Const AD_ACT03_VOLT As Integer = 6
Public Const AD_ACT04_CURR As Integer = 7
Public Const AD_ACT04_VOLT As Integer = 8
Public Const AD_SENSOR1 As Integer = 9
Public Const AD_SENSOR2 As Integer = 10
Public Const AD_SENSOR3 As Integer = 11
Public Const AD_SENSOR4 As Integer = 12
Public Const AD_SENSOR5 As Integer = 13
Public Const AD_SENSOR6 As Integer = 14

Public Const AD_BLOWER_CURR As Integer = 999
Public Const AD_BLOWER_RPM As Integer = 999
Public Const AD_LIN_ACT_CURR As Integer = 999
Public Const AD_VIB As Integer = 999
Public Const AD_ION As Integer = 999

Public Const DA_ACT01 As Integer = 2
Public Const DA_ACT02 As Integer = 3
Public Const DA_ACT03 As Integer = 4
Public Const DA_ACT04 As Integer = 5

Public Const DA_SUPPLY As Integer = 999
Public Const DA_ACT05 As Integer = 999
Public Const DA_ACT06 As Integer = 999
Public Const DA_ACT07 As Integer = 999
Public Const DA_ACT08 As Integer = 999

Public Const I_START_SW As Integer = 0
Public Const I_STOP_SW As Integer = 1
Public Const I_AUTO_SW As Integer = 2

Public Const I_WORK_ON As Integer = 999
Public Const I_WORK_OFF As Integer = 999
Public Const I_VIB_ON As Integer = 999
Public Const I_VIB_OFF As Integer = 999
Public Const I_MARKING1_ON As Integer = 999
Public Const I_MARKING1_OFF As Integer = 999

Public Const I_LIN_ACT01 As Integer = 999

Public Const O_START_LAMP As Integer = 0
Public Const O_STOP_LAMP As Integer = 1
Public Const O_OK_LAMP As Integer = 2
Public Const O_NG_LAMP As Integer = 3
Public Const O_RUN_LAMP As Integer = 4
Public Const O_BUZZER As Integer = 5
Public Const O_MSOK As Integer = 6

Public Const O_ACT01_POWER As Integer = 10
Public Const O_ACT02_POWER As Integer = 11
Public Const O_ACT03_POWER As Integer = 12
Public Const O_ACT04_POWER As Integer = 13
Public Const O_SENSOR_POWER As Integer = 14

Public Const O_STEPPING1_RESET As Integer = 16
Public Const O_STEPPING2_RESET As Integer = 17

Public Const O_BLOWER_POWER As Integer = 999
Public Const O_BLOWER_PWM As Integer = 999
Public Const O_BLOWER_01 As Integer = 999
Public Const O_BLOWER_02 As Integer = 999
Public Const O_BLOWER_03 As Integer = 999
Public Const O_BLOWER_04 As Integer = 999
Public Const O_BLOWER_05 As Integer = 999
Public Const O_BLOWER_06 As Integer = 999
Public Const O_BLOWER_07 As Integer = 999
Public Const O_BLOWER_08 As Integer = 999

Public Const O_MSNG As Integer = 999
Public Const O_ACT05_POWER As Integer = 999
Public Const O_ACT06_POWER As Integer = 999
Public Const O_ACT07_POWER As Integer = 999
Public Const O_ACT08_POWER As Integer = 999
Public Const O_BLOWER_DIRECTION As Integer = 999
Public Const O_ION_POWER As Integer = 999
Public Const O_ION_DIAG As Integer = 999

Public Const O_LIN_ACT_POWER As Integer = 999
Public Const O_LIN_POWER As Integer = 999

Public Const O_WORK_ON As Integer = 999
Public Const O_WORK_OFF As Integer = 999
Public Const O_VIB As Integer = 999
Public Const O_MARKING1 As Integer = 999

Public Const RT_ERROR As Integer = -1

' PLC NG LIST
Public Const PLC_ACT01_CURR As Long = &H1
Public Const PLC_ACT01_VOLT As Long = &H2
Public Const PLC_ACT02_CURR As Long = &H4
Public Const PLC_ACT02_VOLT As Long = &H8
Public Const PLC_ACT03_CURR As Long = &H10
Public Const PLC_ACT03_VOLT As Long = &H20
Public Const PLC_ACT04_CURR As Long = &H40
Public Const PLC_ACT04_VOLT As Long = &H80
Public Const PLC_SENSOR1 As Long = &H100
Public Const PLC_SENSOR2 As Long = &H200
Public Const PLC_SENSOR3 As Long = &H400
Public Const PLC_SENSOR4 As Long = &H800
Public Const PLC_SENSOR5 As Long = &H1000
Public Const PLC_SENSOR6 As Long = &H2000

Public Const PLC_BLOWER_CURR As Long = &H0
Public Const PLC_BLOWER_RPM As Long = &H0
Public Const PLC_BLOWER_VIB As Long = &H0
Public Const PLC_ACT01_ANGLE As Long = &H0
Public Const PLC_ACT02_ANGLE As Long = &H0
Public Const PLC_ACT03_ANGLE As Long = &H0
Public Const PLC_ACT04_ANGLE As Long = &H0
Public Const PLC_PTC As Long = &H0
Public Const PLC_ALL As Long = &HFFFF

' LIN
Public Const BYTE_DATASIZE As Integer = 8

Public Const BYTE_UNKNOWN01 As Long = &H22
Public Const BYTE_UNKNOWN02 As Long = &H3C

Public Const BYTE1_ACT01 As Long = &H5
Public Const BYTE1_ACT02 As Long = &H7
Public Const BYTE1_ACT03 As Long = &H6
Public Const BYTE1_ACT04 As Long = &H4
Public Const BYTE1_PTC As Long = &H26
Public Const BYTE1_BLOWER As Long = &H32

Public Const BYTE_MIN As Long = &H0
Public Const BYTE_MAX As Long = &HFF

Public Const BYTE1_BROADCASTING As Long = &H7F ' all move?
Public Const BYTE1_AUTOADDRESS As Long = &H3C

Public Const BYTE2_STORE As Long = &H6
Public Const BYTE2_CLEAR As Long = &HF0
