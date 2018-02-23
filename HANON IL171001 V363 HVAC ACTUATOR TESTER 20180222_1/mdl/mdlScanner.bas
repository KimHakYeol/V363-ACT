Attribute VB_Name = "mdlScanner"
Option Explicit

Public Function ScannerOpen() As Boolean
    If SCANNERUSE = False Then
        ScannerOpen = True
        
        Exit Function
    End If
    
    If SysVar.nScannerPort = 0 Then
        Call MsgBox("SCANNER PORT CHECK...")
        
        ScannerOpen = False
        
        Exit Function
    End If
    
    frmMain.comScanner.CommPort = SysVar.nScannerPort
    
    If frmMain.comScanner.PortOpen = False Then
        frmMain.comScanner.PortOpen = True
        
        Call Delay(20)
        
        If frmMain.comScanner.PortOpen Then
            ScannerOpen = True
        Else
            ScannerOpen = False
        End If
    End If
End Function

Public Sub ScannerClose()
    If SCANNERUSE = False Then Exit Sub
    
    If frmMain.comScanner.PortOpen = True Then
        frmMain.comScanner.PortOpen = False
    End If
    
    Call Delay(20)
End Sub

Public Sub ScannerReceived()
    Dim i As Integer
    Dim a As String
    Dim B As Integer
    Dim lpActData(1) As String
    Dim lpBarcodeData As String
    
    If SCANNERUSE = False Then Exit Sub
    
    a = frmMain.comScanner.Input
    
    If a <> "" Then
        B = Asc(a)
        
        ' 10 = CR ' carrage return
        ' 13 = LF ' line feed
        
        If B = 10 Then
            lpScannerData = ""
        ElseIf B = 13 Then
            
            Call OnLog("[SCANNER] DATA : " & lpScannerData)
'            Call OnLog("[SCANNER] COUNT : " & nScannerCount)
'            Call OnLog("[SCANNER] SAVE : " & SetupVar.lpBarCode(0))
            
            frmRun.pnlScannerResult.Caption = lpScannerData
            
            '바코드 인쇄값과 스캔값 비교
            For i = 0 To 1
                If frmRun.pnlAct04Volt(i).Caption = "" Then
                    lpActData(i) = "0.00"
                Else
                    lpActData(i) = frmRun.pnlAct04Volt(i).Caption
                End If
            Next
            
            lpBarcodeData = Trim$(SetupVar.lpBarcode(4)) & " " & Trim$(SetupVar.lpBarcode(2)) & " " & Trim$(SetupVar.lpBarcode(3)) & " " & Format(Now, "YYYYMMDD") & "-" & Format(SysVar.lOkCounter + 1, "0000") & " " & Format(lpActData(0), "0.00") & " " & Format(lpActData(1), "0.00")
            
            If lpBarcodeData = lpScannerData Then
                frmRun.pnlScannerResult.BackColor = vbGreen
                nScannerResult = 1 ' same
            Else
                frmRun.pnlScannerResult.BackColor = vbRed
                nScannerResult = 2 ' diff
                nScannerCount = nScannerCount - 1
                
'                Call OnLog("[SCANNER] RESULT DIFF")
            End If
        Else
            lpScannerData = lpScannerData & a
        End If
    End If
End Sub

