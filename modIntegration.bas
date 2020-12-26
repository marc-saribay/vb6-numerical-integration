Attribute VB_Name = "modIntegration"
Option Explicit
Public dblXsubI, dblA0, dblArea As Double, strTitle As String

Public Sub IntegrationMethod(intMethod As Integer)
  On Error GoTo ErrorHandler
  Dim strPrompt, strPrompt2, strPrompt3, strPrompt4, strPrompt5, Msg As String
  Dim dblDeltaX0, dblB0 As Double, intN0 As Integer
  strTitle = "Numerical Integration"
  strPrompt = "Please enter a numeric value for a, b, and n"
  strPrompt2 = "Please satisfy the following conditions (A > 0 or A = 0, B > A, N > 0)"
  strPrompt3 = "Please enter an even number for N (Simpson's Rule needs an even number N)"
  strPrompt4 = "Overflow Error: Value is too great!"
  strPrompt5 = "Please enter a function - f(x)"
  If Not IsNumeric(frmMain.txtValue(1).Text) Or Not IsNumeric(frmMain.txtValue(2).Text) Or Not IsNumeric(frmMain.txtValue(3).Text) Then
    Msg = MsgBox(strPrompt, vbOKOnly + vbInformation, strTitle)
    EmptyTextFields
  Else
    If Int(frmMain.txtValue(1).Text) >= 0 And Int(frmMain.txtValue(1).Text) < Int(frmMain.txtValue(2).Text) And Int(frmMain.txtValue(3).Text) > 0 Then
      If frmMain.txtFunction.Text = "" Then
        Msg = MsgBox(strPrompt5, vbOKOnly + vbInformation, strTitle)
      Else
        dblA0 = frmMain.txtValue(1).Text
        dblB0 = frmMain.txtValue(2).Text
        intN0 = frmMain.txtValue(3).Text
        dblDeltaX0 = (dblB0 - dblA0) / intN0
        dblXsubI = dblA0
        Select Case intMethod
          Case Is = 1
            Call Rectangular(dblA0, dblB0, dblDeltaX0, dblArea, intN0)
          Case Is = 2
            Call Midpoint(dblA0, dblB0, dblDeltaX0, dblArea, intN0)
          Case Is = 3
            Call Trapezoidal(dblA0, dblB0, dblDeltaX0, dblArea, intN0)
          Case Is = 4
            If intN0 Mod 2# = 0 Then
              Call Simpsons(dblA0, dblB0, dblDeltaX0, dblArea, intN0)
            Else
              Msg = MsgBox(strPrompt3, vbOKOnly + vbInformation, strTitle)
              frmMain.txtValue(3).Text = ""
              frmMain.txtValue(3).SetFocus
            End If
        End Select
      End If
    Else
      Msg = MsgBox(strPrompt2, vbOKOnly + vbInformation, strTitle)
    End If
    frmMain.lblArea.Caption = Str(dblArea)
  End If
ErrorHandler:
  If Err.Number = 6 Then
    Msg = MsgBox(strPrompt4, vbOKOnly + vbCritical, strTitle)
    EmptyTextFields
    Err.Clear
  End If
End Sub
  
Private Sub Rectangular(dblA, dblB, dblDeltaX, dblArea As Double, intN As Integer)
  Dim intI As Integer, dblMiddle As Double
  frmMain.fraOutput.Caption = "Rectangular Method"
  For intI = 1 To intN
    dblXsubI = dblXsubI + dblDeltaX
    dblMiddle = dblMiddle + (FofX(dblXsubI))
  Next
  dblArea = dblDeltaX * dblMiddle
End Sub

Private Sub Midpoint(dblA, dblB, dblDeltaX, dblArea As Double, intN As Integer)
  Dim intI As Integer, dblXsubIp1, dblMiddle As Double
  frmMain.fraOutput.Caption = "Midpoint Method"
  For intI = 1 To intN
    dblXsubIp1 = dblXsubI + dblDeltaX
    dblMiddle = dblMiddle + ((FofX(dblXsubI) + FofX(dblXsubIp1)) / 2)
    dblXsubI = dblXsubIp1
  Next
  dblArea = dblDeltaX * dblMiddle
End Sub

Private Sub Trapezoidal(dblA, dblB, dblDeltaX, dblArea As Double, intN As Integer)
  Dim intI As Integer, dblMiddle As Double
  frmMain.fraOutput.Caption = "Trapezoidal Rule"
  For intI = 1 To intN - 1
    dblXsubI = dblXsubI + dblDeltaX
    dblMiddle = dblMiddle + (2 * FofX(dblXsubI))
  Next
  dblArea = (dblDeltaX / 2) * (FofX(dblA) + dblMiddle + FofX(dblB))
End Sub

Private Sub Simpsons(dblA, dblB, dblDeltaX, dblArea As Double, intN As Integer)
  Dim intI As Integer, dblMiddle As Double
  frmMain.fraOutput.Caption = "Simpson's Rule"
  For intI = 1 To intN - 1
    dblXsubI = dblXsubI + dblDeltaX
    If intI Mod 2# = 0 Then
      dblMiddle = dblMiddle + (2 * FofX(dblXsubI))
    Else
      dblMiddle = dblMiddle + (4 * FofX(dblXsubI))
    End If
  Next
  dblArea = (dblDeltaX / 3) * (FofX(dblA) + dblMiddle + FofX(dblB))
End Sub
      
Private Function FofX(ByVal X As Double) As Double
  Dim strPrompt, Msg As String
  strPrompt = "There was an error in evaluating the function"
  On Error GoTo FncError
  frmMain.ScriptControl1.AddCode "X=" & X
  FofX = frmMain.ScriptControl1.Eval(frmMain.txtFunction.Text)
FncError:
  If Err.Number <> 0 Then
    Msg = MsgBox(strPrompt, vbOKOnly + vbInformation, strTitle)
    EmptyTextFields
    frmMain.txtFunction.Text = ""
    frmMain.txtFunction.SetFocus
    Err.Clear
  End If
End Function

Private Sub EmptyTextFields()
  frmMain.txtValue(1).Text = ""
  frmMain.txtValue(2).Text = ""
  frmMain.txtValue(3).Text = ""
  frmMain.lblArea.Caption = "0"
  frmMain.txtValue(1).SetFocus
End Sub

