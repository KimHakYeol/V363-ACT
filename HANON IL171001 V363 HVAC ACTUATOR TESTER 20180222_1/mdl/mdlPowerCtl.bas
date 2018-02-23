Attribute VB_Name = "mdlPowerCtl"
Option Explicit

Public Function PowerOpen() As Boolean
    Dim nRes As Integer
    Dim lpChar As String
    
    If POWERTYPE > 0 Then
        PowerOpen = True
        
        If SysVar.nPowerPort = 0 Then
            PowerOpen = False
            
            Exit Function
        End If
        
        frmMain.Power.CommPort = SysVar.nPowerPort
        
        If frmMain.Power.PortOpen = False Then
            frmMain.Power.InputLen = Choose(POWERTYPE, 6, 8)
            frmMain.Power.PortOpen = True
            Call Delay(50)
        End If
        
        nRes = 0

ADRAGAIN:
        
        lpChar = frmMain.Power.Input
        lpChar = ""
        
        Select Case POWERTYPE
            Case 1: frmMain.Power.Output = "ADR 6" & vbCrLf
            Case 2:
                frmMain.Power.Output = ":ADR01;"
                
                Call Sleep(100)
                
                frmMain.Power.Output = ":DCL;"
                
                Call Sleep(100)
                
                frmMain.Power.Output = ":RMT1;"
                
                Call Sleep(100)
                
        End Select
        
        If POWERTYPE = 1 Then
            Call SetTime(TM_SUPPLY_POWER)
            Do
                DoEvents
                
                If frmMain.Power.InBufferCount <> 0 Then lpChar = lpChar & frmMain.Power.Input
                If ElapseTime(TM_SUPPLY_POWER) > 1 Then Exit Do
            Loop
            
            If Left$(lpChar, 2) <> "OK" Then
                nRes = nRes + 1
                
                If nRes < 5 Then
                    GoTo ADRAGAIN
                Else
                    Call OnLog("[ERROR] POWER OPEN 1 FAIL...")
                    
                    PowerOpen = False
                    
                    Exit Function
                End If
            End If
        Else
            PowerOpen = True
        End If
        
        nRes = 0

OUTAGAIN:
        
        lpChar = frmMain.Power.Input
        lpChar = ""
        
        Select Case POWERTYPE
            Case 1: frmMain.Power.Output = "OUT ON" & vbCrLf
            Case 2: frmMain.Power.Output = ":OUT1;"
        End Select
        
        If POWERTYPE = 1 Then
            Call SetTime(TM_SUPPLY_POWER)
            Do
                DoEvents
                
                If frmMain.Power.InBufferCount <> 0 Then lpChar = lpChar + frmMain.Power.Input
                If ElapseTime(TM_SUPPLY_POWER) > 2 Then Exit Do
            Loop
            
            If Left$(lpChar, 2) <> "OK" Then
                nRes = nRes + 1
                
                If nRes < 5 Then
                    GoTo OUTAGAIN
                Else
                    Call OnLog("[ERROR] POWER OPEN 2 FAIL...")
                    
                    PowerOpen = False
                    
                    Exit Function
                End If
            End If
        Else
            PowerOpen = True
        End If
    End If
End Function

Public Function PowerClose()
    If frmMain.Power.PortOpen = True Then
        frmMain.Power.PortOpen = False
        Call Sleep(25)
    End If
End Function

Public Sub SetPowerProcess()
    If POWERTYPE = 0 Then Exit Sub
    
    If PowerVar.nCount = 0 Then
        If PowerVar.bVolt = False Then
            PowerVar.bVolt = True
            
            Select Case POWERTYPE
                Case 1: frmMain.Power.Output = "MV?" & vbCrLf
                Case 2: frmMain.Power.Output = ":VOL?;"
            End Select
        End If
        
        Call SetTime(TM_SUPPLY_POWER)
        PowerVar.nCount = 1
    End If
    
    If PowerVar.nCount = 1 Then
        If ElapseTime(TM_SUPPLY_POWER) > 0.2 Then
            PowerVar.bVolt = False
            PowerVar.nCount = 1000
        End If
    End If
    
    If PowerVar.nCount = 1000 Then
        If PowerVar.bCurr = False Then
            PowerVar.bCurr = True
            
            Select Case POWERTYPE
                Case 1: frmMain.Power.Output = "MC?" & vbCrLf
                Case 2: frmMain.Power.Output = ":CUR?;"
            End Select
        End If
        
        Call SetTime(TM_SUPPLY_POWER)
        PowerVar.nCount = 1001
    End If
    
    If PowerVar.nCount = 1001 Then
        If ElapseTime(TM_SUPPLY_POWER) > 0.2 Then
            PowerVar.bCurr = False
            PowerVar.nCount = 20000
        End If
    End If
    
    If PowerVar.nCount = 20000 Then
        If PowerVar.bCurr = False Then
            PowerVar.nCount = 0
        End If
    End If
End Sub

Public Sub GetPowerProcess()
    Dim A As String
    Dim B As Integer
    
    If POWERTYPE = 0 Then Exit Sub
    If frmMain.Power.CommEvent <> comEvReceive Then Exit Sub
    
    A = frmMain.Power.Input
    
    If A <> "" Then
        B = Asc(A)
        
        If B = 13 Then
            If PowerVar.bVolt Then
                Select Case POWERTYPE
                    Case 1:
                        dSupplyVolt = Val(PowerVar.iCOM)
                    
                    Case 2:
                        If Left$(PowerVar.iCOM, 2) = "AV" Then
                            dSupplyVolt = Val(Mid(PowerVar.iCOM, 3))
                        End If
                    
                End Select
                
                PowerVar.bVolt = False
            End If
            
            If PowerVar.bCurr Then
                Select Case POWERTYPE
                    Case 1:
                        dSupplyCurr = Val(PowerVar.iCOM)
                    
                    Case 2:
                        If Left$(PowerVar.iCOM, 2) = "AA" Then
                            dSupplyCurr = Val(Mid(PowerVar.iCOM, 3))
                        End If
                    
                End Select
                
                PowerVar.bCurr = False
            End If
            
            PowerVar.iCOM = ""
        Else
            PowerVar.iCOM = PowerVar.iCOM & A
        End If
    End If
End Sub

