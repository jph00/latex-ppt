Attribute VB_Name = "LatexModule"
Private Function GetMathZoneAtCursor() As TextRange2
    ' Get current cursor location
    Dim cursorPt As Long
    cursorPt = Application.ActiveWindow.Selection.TextRange2.Start
    
    ' while browse all equations...
    Dim mathStart, mathEnd As Long
    Dim mathZone As TextRange2
    
    With Application.ActiveWindow.Selection.ShapeRange.TextFrame2
        For Each mathZone In .TextRange.MathZones
            mathStart = mathZone.Start
            mathEnd = mathZone.Start + mathZone.Length
            If cursorPt >= mathStart And cursorPt <= mathEnd Then
                Set GetMathZoneAtCursor = mathZone
            End If
        Next
    End With
End Function

Sub NewLatex()
  ' Creat a new MathZone
  Dim aNewMathZone As TextRange2
  Application.CommandBars.ExecuteMso ("EquationInsertNew")
  Set aNewMathZone = GetMathZoneAtCursor()
  
  ' Create dialog box for user to input the equation source,
  ' then set the MathZone text
  aNewMathZone.Text = InputBox("Type new equation below", "LaTeX Editor")
  
  ' Convert format to Professional
  Application.CommandBars.ExecuteMso ("EquationProfessional")
End Sub

Sub EditLatex()
  ' Convert the MathZone under cursor to Linear format
  ' and select it
  Dim aMathZone As TextRange2
  Application.CommandBars.ExecuteMso ("EquationLinearFormat")
  Set aMathZone = GetMathZoneAtCursor()
  
  ' Create dialog box for user to input the desired changes,
  ' (the default dialog box text is not empty, but has the current equation to be edited)
  ' then set the MathZone text
  aMathZone.Text = InputBox("Edit the equation below", "LaTeX Editor", aMathZone.Text)
  
  ' Convert format back to Professional
  Application.CommandBars.ExecuteMso ("EquationProfessional")
End Sub

Sub SwitchLatex()
  ' Creat a new MathZone
  Dim aNewMathZone As TextRange2
  Application.CommandBars.ExecuteMso ("EquationInsertNew")
  Set aNewMathZone = GetMathZoneAtCursor()
  
  ' Type in that zone the special character sequence that switches the equation format
  aNewMathZone.Text = ChrW(&H24C9)
End Sub

Sub SwitchUnicode()
  ' Creat a new MathZone
  Dim aNewMathZone As TextRange2
  Application.CommandBars.ExecuteMso ("EquationInsertNew")
  Set aNewMathZone = GetMathZoneAtCursor()
  
  ' Type in that zone the special character sequence that switches the equation format
  aNewMathZone.Text = ChrW(&H24C1)
End Sub

