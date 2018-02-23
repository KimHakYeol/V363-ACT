Attribute VB_Name = "mdlCaptureUpload"
Option Explicit
 
 
' ScreenCapture
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Const SRCCOPY As Long = &HCC0020

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

' FtpUpload
Public Const BUFFERSIZE = 255
Public Const INTERNET_FLAG_PASSIVE = &H8000000
Public Const INTERNET_FLAG_ACTIVE = &O0
Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0

Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sBuff As String, ByVal Access As Long, ByVal Flags As Long, ByVal Context As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToRead As Long, dwNumberOfBytesRead As Long) As Integer
Public Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToWrite As Long, dwNumberOfBytesWritten As Long) As Integer

Private hOpen As Long
Private hConnection As Long
Private hFile As Long
Private dwType As Long

Public Function FtpOpen(strIP As String, strPort As String, strUser As String, strPassword As String) As Boolean
    '기존에 이미 접속되어 있으면 기존 접속 종료
    If hConnection <> 0 Then Call InternetCloseHandle(hConnection)
   
    '접속에대한 핸들링값 얻음
    hOpen = InternetOpen("FTP Module", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
   
    '접속시도
    hConnection = InternetConnect(hOpen, strIP, strPort, strUser, strPassword, 1, INTERNET_FLAG_ACTIVE, 0)
   
    '==== 접속이 되었다면 기본적인 값을 설정해준다.
    If hConnection <> 0 Then
        FtpOpen = True
        dwType = FTP_TRANSFER_TYPE_BINARY
    End If
End Function

'파일 업로드
Public Function FTPUpload(strLocal As String, strRemote As String) As Boolean
    Dim Data(BUFFERSIZE - 1) As Byte
    Dim Written As Long
    Dim lonSize As Long
    Dim Sum As Long
    Dim lBlock As Long
    
    Sum = 0
    lBlock = 0
    FTPUpload = False
    
    '로컬파일이나 원격파일 파라미터가 공백인지 체크
    If strLocal <> "" And strRemote <> "" Then
        hFile = FtpOpenFile(hConnection, strRemote, GENERIC_WRITE, dwType, 0)
        
        If hFile = 0 Then Exit Function
        
        Open strLocal For Binary Access Read As #1
            lonSize = FileLen(strLocal)
            
            For lBlock = 1 To lonSize \ BUFFERSIZE
                Get #1, , Data
                
                If (InternetWriteFile(hFile, Data(0), BUFFERSIZE, Written) = 0) Then Exit Function
                
                DoEvents
                
                Sum = Sum + BUFFERSIZE
                
                DoEvents
            
            Next
            
            Get #1, , Data
            
            If (InternetWriteFile(hFile, Data(0), lonSize Mod BUFFERSIZE, Written) = 0) Then Exit Function
            
            Sum = Sum + (lonSize Mod BUFFERSIZE)
            lonSize = Sum
        Close #1
        
        Call InternetCloseHandle(hFile)
        FTPUpload = True
    End If
End Function

'FTP접속종료
Public Sub FtpClose()
    If hConnection <> 0 Then InternetCloseHandle hConnection
    
    hConnection = 0
End Sub

Public Sub ImageFileSend()
    Dim nCount As Integer
    Dim nRes As Integer
    Dim lpPathName As String
    Dim lpSharingName As String
    Dim lpModelNameArr() As String
    
    On Error Resume Next
    
RESTART:
    
    nRes = 3 ' 1 = MIRAE, 2 = EDS, 3 = GONG
    
    lpPathName = App.Path & "\" & lpBmpFileName & ".bmp"
    
    Select Case nRes
        Case 1:
            ' Ftp Upload
            If FtpOpen(SysVar.lpFtpIp, SysVar.lpFtpPort, SysVar.lpFtpId, SysVar.lpFtpPw) Then
                Call OnLog("FTP SERVER CONNECT SUCCESS...")
            Else
                Call OnLog("FTP SERVER CONNECT FAIL...")
                
                Exit Sub
            End If
            
            If FTPUpload(lpPathName, lpBmpFileName & ".bmp") Then
                Call OnLog("FTP SERVER SEND SUCCESS...")
            Else
                Call OnLog("FTP SERVER SEND FAIL...")
            End If
            
            Call FtpClose
        
        Case 2:
            ' IP 가 아니라 공유 PC 이름을 사용해야함
            If SysVar.lpFtpIp <> "" Then
                ' Sharing Folder Copy
                lpSharingName = "\\" & SysVar.lpFtpIp & "\\" & "NGIMAGE"
                lpModelNameArr = Split(lpNowModel, "_")
                
                If DataVar.lpPallet = "" Then
                    DataVar.lpPallet = "0"
                End If
                
                Call SearchDir(lpSharingName & "\\" & lpModelNameArr(1))
                Call SearchDir(lpSharingName & "\\" & lpModelNameArr(1) & "\\" & lpModelNameArr(2))
                Call SearchDir(lpSharingName & "\\" & lpModelNameArr(1) & "\\" & lpModelNameArr(2) & "\\" & DataVar.lpPallet)
                Call FileCopy(App.Path & "\" & lpBmpFileName & ".bmp", lpSharingName & "\\" & lpModelNameArr(1) & "\\" & lpModelNameArr(2) & "\\" & DataVar.lpPallet & "\\" & lpBmpFileName & ".bmp")
                
                If SearchFile(lpSharingName & "\\" & lpModelNameArr(1) & "\\" & lpModelNameArr(2) & "\\" & DataVar.lpPallet & "\\" & lpBmpFileName & ".bmp") Then
                    Call OnLog("SHARING FOLDER COPY SUCCESS...")
                Else
                    nCount = nCount + 1
                    
                    Call OnLog("SHARING FOLDER COPY FAIL... " & nCount)
                    
                    If nCount > 3 Then
                        Call OnLog("SHARING FOLDER ERROR...")
                    Else
                        GoTo RESTART
                    End If
                End If
            Else
                Call OnLog("IP ADDRESS NOTHING...")
            End If
        
        Case 3:
            ' IP 가 아니라 공유 PC 이름을 사용해야함
            If SysVar.lpFtpIp <> "" Then
                ' Sharing Folder Copy
                lpSharingName = "\\" & SysVar.lpFtpIp & "\\" & "REWORK"
                
                If DataVar.lpPallet = "" Then
                    DataVar.lpPallet = "0"
                End If
                
                Call SearchDir(lpSharingName & "\\" & DataVar.lpPallet)
                Call FileCopy(App.Path & "\" & lpBmpFileName & ".bmp", lpSharingName & "\\" & DataVar.lpPallet & "\\" & lpBmpFileName & ".bmp")
                
                If SearchFile(lpSharingName & "\\" & DataVar.lpPallet & "\\" & lpBmpFileName & ".bmp") Then
                    Call OnLog("SHARING FOLDER COPY SUCCESS...")
                Else
                    nCount = nCount + 1
                    
                    Call OnLog("SHARING FOLDER COPY FAIL... " & nCount)
                    
                    If nCount > 3 Then
                        Call OnLog("SHARING FOLDER ERROR...")
                    Else
                        GoTo RESTART
                    End If
                End If
            Else
                Call OnLog("IP ADDRESS NOTHING...")
            End If
    
    End Select
    
    Call Kill(lpPathName)
End Sub

Public Sub SaveCapture()
    Dim llngDeskTopHWnd         As Long
    Dim llngDeskTopDC           As Long
    Dim ludtDeskTopRect         As RECT
    Dim llngDeskTopHeight       As Long
    Dim llngDeskTopWidth        As Long
    
    On Error GoTo 0
    
    llngDeskTopHWnd = GetDesktopWindow
    
    Call GetWindowRect(llngDeskTopHWnd, ludtDeskTopRect)
    
    llngDeskTopHeight = ludtDeskTopRect.Bottom
    llngDeskTopWidth = ludtDeskTopRect.Right
    
    frmRun.picCapture.Height = llngDeskTopHeight * Screen.TwipsPerPixelY
    frmRun.picCapture.Width = llngDeskTopWidth * Screen.TwipsPerPixelX
    
    llngDeskTopDC = GetDC(llngDeskTopHWnd)
    
    Call BitBlt(frmRun.picCapture.hdc, 0, 0, llngDeskTopWidth, llngDeskTopHeight, llngDeskTopDC, 0, 0, SRCCOPY)
    Call ReleaseDC(llngDeskTopHWnd, llngDeskTopDC)
    
    frmRun.imgCapture.Picture = frmRun.picCapture.Image
    frmRun.Refresh
    
    lpBmpFileName = Trim$(frmRun.pnlCarType(3).Caption) & "_ACT" ' _BLW _HTR
    
    Call SavePicture(frmRun.picCapture.Image, App.Path & "\" & lpBmpFileName & ".bmp")
End Sub

