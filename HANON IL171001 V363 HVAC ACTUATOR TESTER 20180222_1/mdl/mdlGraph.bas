Attribute VB_Name = "mdlGraph"
Option Explicit

Private Sub SWAP(ByRef A As Double, ByRef B As Double)
    Dim c As Double
    
    c = A
    A = B
    B = c
End Sub

Public Sub BubbleSort(ByRef dBuf() As Double, ByVal nCount As Integer)
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To nCount
        For j = i To nCount
            If dBuf(2, i) < dBuf(2, j) Then
                Call SWAP(dBuf(2, i), dBuf(2, j))
            End If
        Next
    Next
End Sub

Public Sub TestProc(ByRef GraphVar As TagGraphVariables, ByRef picBox As PictureBox, ByRef GraphBuf() As Double, ByVal BufCount As Integer)
    Static XV As Double
    Static YV As Double
    
    If BufCount = 0 Then Exit Sub
    
    XV = GraphBuf(0, BufCount - 1)
    YV = GraphBuf(1, BufCount - 1)
    
    If GraphVar.YE - GraphVar.YS > 0# And GraphVar.XE - GraphVar.XS > 0# Then
        picBox.Line (GraphVar.OVX, GraphVar.OVY)-(XV, YV), vbBlue
        GraphVar.OVX = XV
        GraphVar.OVY = YV
    End If
End Sub

Public Sub DrawGrid(ByRef picBox As PictureBox, ByRef GraphVar As TagGraphVariables, ByVal nCount As Integer, Optional ByVal bOptimized As Boolean = False)
    Dim Point As Double         ' FOR-NEXT문의 변수
    Dim SoftLineColor As Long   ' 그리드의 색
    Dim DeepLineColor As Long
    Dim GridStepX As Double     ' X축 그리드 간격
    Dim GridStepY As Double     ' Y축 그리드 간격
    Dim runTime As Double
    Dim scaleXY(3) As Single
    Dim lGridStep(2) As Long
    
    Select Case picBox.Name
        Case "picVib":
            runTime = IIf(SetupVar.nVibMethod = 0, SetupVar.dBlowerTime(POS_BLOWER_HI), SetupVar.dVibTime)
            scaleXY(0) = 0.07
            scaleXY(1) = 0.07
            scaleXY(2) = 0.03
            scaleXY(3) = 0.15
            lGridStep(0) = 8#
            lGridStep(1) = 10#
            lGridStep(2) = 2
    
        Case "picIntake":
            runTime = 5
            scaleXY(0) = 0.07
            scaleXY(1) = 0.07
            scaleXY(2) = 0.03
            scaleXY(3) = 0.15
            lGridStep(0) = 10#
            lGridStep(1) = 10#
            lGridStep(2) = 5
    
        Case Else:
            Debug.Print "nothing name..."
            Exit Sub
    
    End Select
    
    GraphVar.XS = 0
    
    If bOptimized Then
        GraphVar.XE = nCount
    Else
        GraphVar.XE = 1000 * runTime
    End If
    
    GraphVar.YS = 0
    GraphVar.YE = CInt(lGridStep(2))
    
    GridStepY = (GraphVar.YE - GraphVar.YS) / lGridStep(0)
    GridStepX = (GraphVar.XE - GraphVar.XS) / lGridStep(1)
    picBox.Font.Name = "Tahoma"
    picBox.Font.Size = 8
    picBox.ForeColor = QBColor(0)
    
'    SoftLineColor = &HE0E0E0    '흐린 그리드라인(전부다그림)
'    DeepLineColor = &H8000000C  '진한 그리드라인(건너뛰어그림)
    SoftLineColor = &H8000000C '흐린 그리드라인(전부다그림)
    DeepLineColor = &H8000000C '진한 그리드라인(건너뛰어그림)
    
    If GraphVar.XE - GraphVar.XS <= 0# Then Exit Sub
    If GraphVar.YE - GraphVar.YS <= 0# Then Exit Sub
    
    picBox.Cls
    picBox.DrawWidth = 1
    
    picBox.Scale (GraphVar.XS - ((GraphVar.XE - GraphVar.XS) * scaleXY(0)), GraphVar.YE + ((GraphVar.YE - GraphVar.YS) * scaleXY(1)))-(GraphVar.XE + ((GraphVar.XE - GraphVar.XS) * scaleXY(2)), GraphVar.YS - ((GraphVar.YE - GraphVar.YS) * scaleXY(3)))
    
    ' 연한 그리드 라인 X축으로 이동하며 Y축 그리기
'    For Point = GraphVar.XS To GraphVar.XE Step GridStepX
'        picBox.Line (Point, GraphVar.YS)-(Point, GraphVar.YE), SoftLineColor
'    Next Point
    
    ' 연한 그리드 라인 Y축으로 이동하며 X축 그리기
'    For Point = GraphVar.YS To GraphVar.YE Step GridStepY
'        picBox.Line (GraphVar.XS, Point)-(GraphVar.XE, Point), SoftLineColor
'    Next Point
    
    ' 진한 그리드 라인 X축으로 이동하며 Y축 그리고 X축 숫자 기입
    For Point = GraphVar.XS To GraphVar.XE + 0.0001 Step GridStepX * 2
        picBox.Line (Point, GraphVar.YS - ((GraphVar.YE - GraphVar.YS) * 0.015))-(Point, GraphVar.YE), DeepLineColor
        picBox.CurrentX = Point - (picBox.TextWidth(Format(Point, "#0") + " ") / 2)
        picBox.CurrentY = GraphVar.YS - ((GraphVar.YE - GraphVar.YS) * 0.04)
        picBox.ForeColor = vbBlack
        picBox.Print Format(Point, "#0")
    Next Point
    
    ' 진한 그리드 라인 Y축으로 이동하며 X축 그리고 Y축 숫자 기입
    For Point = GraphVar.YS To GraphVar.YE + 0.0001 Step GridStepY * 2
        picBox.Line (GraphVar.XS - ((GraphVar.XE - GraphVar.XS) * 0.01), Point)-(GraphVar.XE, Point), DeepLineColor
        picBox.CurrentX = GraphVar.XS - ((GraphVar.XE - GraphVar.XS) * 0.03) - (picBox.TextWidth(Format(Point, "####0")))
        picBox.CurrentY = Point - (picBox.TextHeight(Format(Point, "#0") + " ") / 2.5)
        picBox.ForeColor = vbBlack
        picBox.Print Format(Point, "#0.0")
    Next Point
    
    ' 그래프영역 외곽라인
    picBox.Line (GraphVar.XS, GraphVar.YS)-(GraphVar.XE, GraphVar.YE), DeepLineColor, B
End Sub

Public Sub ReDraw(ByVal dValue As Double, ByRef picBox As PictureBox, ByRef GraphVar As TagGraphVariables, ByRef dBuf() As Double, ByVal nCount As Integer)
    Static l    As Long
    Static XXX  As Double
    Static YYY  As Double
    Static ZZZ  As Double
    
    Dim bDraw   As Boolean
    
    bDraw = False
    
    Call DrawGrid(picBox, GraphVar, nCount, True)
    
    GraphVar.OVY = 0#
    GraphVar.OVX = 0#
    GraphVar.OVZ = dBuf(2, 0)
    
    For l = 0 To nCount - 1
        XXX = dBuf(0, l)
        YYY = dBuf(1, l)
        ZZZ = dBuf(2, l)
        
        If GraphVar.XS >= GraphVar.XE Then Exit For
        If GraphVar.YS >= GraphVar.YE Then Exit For
        
        picBox.Line (GraphVar.OVX, GraphVar.OVY)-(XXX, YYY), vbBlue
        
        If SetupVar.nVibResultType = 1 Then ' RMS
            picBox.Line (GraphVar.OVX, GraphVar.OVZ)-(XXX, ZZZ), vbRed
            
            If ZZZ < dValue And bDraw = False Then
                bDraw = True
            End If
        End If
        
        GraphVar.OVY = YYY
        GraphVar.OVX = XXX
        GraphVar.OVZ = ZZZ
    Next
    
    Select Case picBox.Name
        Case "picVib":
            Select Case SetupVar.nVibResultType
                Case 0: ' Peak
                    picBox.DrawWidth = 6
                    picBox.Line (GraphVar.nPeak, dBuf(1, GraphVar.nPeak))-(GraphVar.nPeak, dBuf(1, GraphVar.nPeak)), vbRed, BF
                
                Case 1: ' RMS
                    picBox.DrawStyle = vbSolid
                    picBox.Line (GraphVar.PS, 0)-(GraphVar.PS, 100), vbRed
                    picBox.Line (GraphVar.PE, 0)-(GraphVar.PE, 100), vbRed
            
            End Select
        
        Case "picIntake":
        
        Case Else:
            Debug.Print "picbox 2 name none..."
            
            Exit Sub
        
    End Select
    
End Sub

