Attribute VB_Name = "mdlNvh"
Option Explicit

Public Function NvhOpen() As Boolean
    If NVHUSE = False Then
        NvhOpen = True
        
        Exit Function
    End If
    
    If SysVar.nNvhPort = 0 Then
        Call MsgBox("Nvh PORT CHECK...")
        
        NvhOpen = False
        
        Exit Function
    End If
    
    frmMain.comNvh.CommPort = SysVar.nNvhPort
    
    If frmMain.comNvh.PortOpen = False Then
        frmMain.comNvh.PortOpen = True
        
        Call Delay(20)
        
        If frmMain.comNvh.PortOpen Then
            NvhOpen = True
        Else
            NvhOpen = False
        End If
    End If
End Function

Public Sub NvhClose()
    If NVHUSE = False Then Exit Sub
    
    If frmMain.comNvh.PortOpen = True Then
        frmMain.comNvh.PortOpen = False
    End If
    
    Call Delay(20)
End Sub

Public Sub NvhReceived()
    Dim A As String
    Dim B As Integer
    
    If NVHUSE = False Then Exit Sub
    
    A = frmMain.comNvh.Input
    
    If A <> "" Then
        B = Asc(A)
        
        ' 10 CR
        ' 13 LF
        
        If B = 10 Then
            lpNvhCom = ""
        ElseIf B = 13 Then
            If Left(lpNvhCom, 4) = "Hz: " Then
                lpNvhCom = Replace(lpNvhCom, Left(lpNvhCom, 4), "")
                frmRun.pnlNvhRpm.Caption = lpNvhCom
            End If
        Else
            lpNvhCom = lpNvhCom & A
        End If
    End If
End Sub

Public Function NvhSend(ByVal lpStr As String) As String
    If NVHUSE = False Then Exit Function
    
    lpNvhCom = ""
    
    frmMain.comNvh.OutBufferCount = 0
    
    frmMain.comNvh.Output = lpStr & vbCrLf
    
    NvhSend = lpStr
End Function

