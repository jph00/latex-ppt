Sub PasteLatex()
  Application.CommandBars.ExecuteMso ("EquationInsertNew")
  With ActiveWindow.Selection.ShapeRange.TextFrame.TextRange
    .Characters(.Length - 1) = InputBox("Equation")
  End With
  Application.CommandBars.ExecuteMso ("EquationProfessional")
End Sub

Sub SwitchLatex()
  Application.CommandBars.ExecuteMso ("EquationInsertNew")
  With ActiveWindow.Selection.ShapeRange.TextFrame.TextRange
    .Characters(.Length - 1) = ChrW(&H24C9)
  End With
End Sub

