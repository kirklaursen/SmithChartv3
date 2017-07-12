Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

'
'
'    Ladder Network Module
'
'

Rem function to combine impedances in parallel
Function Ladder_Name(code As Integer)

   Select Case code  ' Evaluate Code.
      Case 1
         MyString = "Series Inductor (nH)"
      Case 2
         MyString = "Shunt Inductor (nH)"
      Case 3
         MyString = "Series Capacitor (pF)"
      Case 4
         MyString = "Shunt Capacitor (pF)"
      Case 5
         MyString = "Series T-Line Line (Zo, E.L.)"
      Case 6
         MyString = "Open Ckt Stub (Zo, E.L.)"
      Case 7
         MyString = "Short Ckt Stub (Zo, E.L.)"
      Case 8
         MyString = "Series Resistor (ohms)"
      Case 9
         MyString = "Shunt Resistor (ohms)"
      Case Else   ' Other values.
         MyString = "Not a valid code"
   End Select
   Ladder_Name = MyString
End Function


Function ReLadder(code As Integer, FGHz As Double, ReZ As Double, ImZ As Double, Val1 As Double, Val2 As Double)
   Dim X As Double
   pi = 3.14159265358979
   Select Case code  ' Evaluate Code.
      Case 1
'         MyString = "Series Inductor (nH)"
          X = 2 * pi * FGHz * Val1
          ReNew = ReZ
          ImNew = ImZ + X
      Case 2
'         MyString = "Shunt Inductor (nH)"
          X = 2 * pi * FGHz * Val1
          ReNew = ReZIIZ(ReZ, ImZ, 0#, X)
          ImNew = ImZIIZ(ReZ, ImZ, 0#, X)
      Case 3
'         MyString = "Series Capacitor (pF)"
          X = -1 / (2 * pi * FGHz * Val1 * 0.001)
          ReNew = ReZ
          ImNew = ImZ + X
      Case 4
'         MyString = "Shunt Capacitor (pF)"
          X = -1 / (2 * pi * FGHz * Val1 * 0.001)
          ReNew = ReZIIZ(ReZ, ImZ, 0#, X)
          ImNew = ImZIIZ(ReZ, ImZ, 0#, X)
      Case 5
'         MyString = "Series Transmission Line (Zo, E.L.)"
          ReNew = ReTline(Val1, Val2, ReZ, ImZ)
          ImNew = ImTline(Val1, Val2, ReZ, ImZ)
      Case 6
'         MyString = "Open Ckt Stub (Zo, E.L.)"
          X = ImTline(Val1, Val2, 1E+20, 0)
          ReNew = ReZIIZ(ReZ, ImZ, 0#, X)
          ImNew = ImZIIZ(ReZ, ImZ, 0#, X)
      Case 7
'         MyString = "Short Ckt Stub (Zo, E.L.)"
          X = ImTline(Val1, Val2, 0, 0)
          ReNew = ReZIIZ(ReZ, ImZ, 0#, X)
          ImNew = ImZIIZ(ReZ, ImZ, 0#, X)
      Case 8
'         MyString = "Series Resistor (ohms)"
          ReNew = ReZ + Val1
          ImNew = ImZ
      Case 9
'         MyString = "Shunt Resistor (ohms)"
          ReNew = ReZIIZ(ReZ, ImZ, Val1, 0)
          ImNew = ImZIIZ(ReZ, ImZ, Val1, 0)
      Case Else   ' Other values.
'         MyString = "Not a valid code"
          ReNew = ReZ
          ImNew = ImZ
   End Select
   ReLadder = ReNew
End Function




Function ImLadder(code As Integer, FGHz As Double, ReZ As Double, ImZ As Double, Val1 As Double, Val2 As Double)
   Dim X As Double
   pi = 3.14159265358979
   Select Case code  ' Evaluate Code.
      Case 1
'         MyString = "Series Inductor (nH)"
          X = 2 * pi * FGHz * Val1
          ReNew = ReZ
          ImNew = ImZ + X
      Case 2
'         MyString = "Shunt Inductor (nH)"
          X = 2 * pi * FGHz * Val1
          ReNew = ReZIIZ(ReZ, ImZ, 0#, X)
          ImNew = ImZIIZ(ReZ, ImZ, 0#, X)
      Case 3
'         MyString = "Series Capacitor (pF)"
          X = -1 / (2 * pi * FGHz * Val1 * 0.001)
          ReNew = ReZ
          ImNew = ImZ + X
      Case 4
'         MyString = "Shunt Capacitor (pF)"
          X = -1 / (2 * pi * FGHz * Val1 * 0.001)
          ReNew = ReZIIZ(ReZ, ImZ, 0#, X)
          ImNew = ImZIIZ(ReZ, ImZ, 0#, X)
      Case 5
'         MyString = "Series Transmission Line (Zo, E.L.)"
          ReNew = ReTline(Val1, Val2, ReZ, ImZ)
          ImNew = ImTline(Val1, Val2, ReZ, ImZ)
      Case 6
'         MyString = "Open Ckt Stub (Zo, E.L.)"
          X = ImTline(Val1, Val2, 1E+20, 0)
          ReNew = ReZIIZ(ReZ, ImZ, 0#, X)
          ImNew = ImZIIZ(ReZ, ImZ, 0#, X)
      Case 7
'         MyString = "Short Ckt Stub (Zo, E.L.)"
          X = ImTline(Val1, Val2, 0, 0)
          ReNew = ReZIIZ(ReZ, ImZ, 0#, X)
          ImNew = ImZIIZ(ReZ, ImZ, 0#, X)
      Case 8
'         MyString = "Series Resistor (ohms)"
          ReNew = ReZ + Val1
          ImNew = ImZ
      Case 9
'         MyString = "Shunt Resistor (ohms)"
          ReNew = ReZIIZ(ReZ, ImZ, Val1, 0)
          ImNew = ImZIIZ(ReZ, ImZ, Val1, 0)
      Case Else   ' Other values.
'         MyString = "Not a valid code"
          ReNew = ReZ
          ImNew = ImZ
   End Select
   ImLadder = ImNew
End Function


'  Function to correlate a format name with a code number
Function format_name(code As Integer)
   Select Case code  ' Evaluate Code.
      Case 1
         MyString = "Unnormalized Z=R+jX"
      Case 2
         MyString = "Normalized z=r+jx"
      Case 3
         MyString = "Rectangular Gamma"
      Case 4
         MyString = "Polar Gamma"
      Case 5
         MyString = "Unnormalized Y=G+jB"
      Case 6
         MyString = "Normalized y=g+jb"
      Case Else   ' Other values.
         MyString = "Invalid code"
   End Select
   format_name = MyString
End Function

'  Function to match name of real part of format to code number
Function Re_format_name(code As Integer)
   Select Case code  ' Evaluate Code.
      Case 1
         MyString = "R"
      Case 2
         MyString = "r"
      Case 3
         MyString = "Re(G)"
      Case 4
         MyString = "Mag(G)"
      Case 5
         MyString = "G"
      Case 6
         MyString = "g"
      Case Else   ' Other values.
         MyString = "~"
   End Select
   Re_format_name = MyString
End Function

'  Function to match name of imaginary part of format to code number
Function Im_format_name(code As Integer)
   Select Case code  ' Evaluate Code.
      Case 1
         MyString = "X"
      Case 2
         MyString = "x"
      Case 3
         MyString = "Im(G)"
      Case 4
         MyString = "Ang(G)"
      Case 5
         MyString = "B"
      Case 6
         MyString = "b"
      Case Else   ' Other values.
         MyString = "~"
   End Select
   Im_format_name = MyString
End Function

'  Function to convert a given format to an impedance
Function Re_format_convert(code As Integer, X1 As Double, X2 As Double, Zo As Double)
   pi = 3.14159265358979
   Select Case code  ' Evaluate Code.
      Case 1
         result = X1
      Case 2
         result = X1 * Zo
      Case 3
         result = GtoReZ(X1, X2, Zo)
      Case 4
         result = GtoReZ(X1 * Cos(X2 * pi / 180), X1 * Sin(X2 * pi / 180), Zo)
      Case 5
         result = ZtoG(X1, X2)
      Case 6
         result = ZtoG(X1 * Zo, X2 * Zo)
      Case Else   ' Other values.
         result = -99
   End Select
   Re_format_convert = result
End Function

'  Function to convert a given format to an impedance
Function Im_format_convert(code As Integer, X1 As Double, X2 As Double, Zo As Double)
   pi = 3.14159265358979
   Select Case code  ' Evaluate Code.
      Case 1
         result = X2
      Case 2
         result = X2 * Zo
      Case 3
         result = GtoImZ(X1, X2, Zo)
      Case 4
         result = GtoImZ(X1 * Cos(X2 * pi / 180), X1 * Sin(X2 * pi / 180), Zo)
      Case 5
         result = ZtoB(X1, X2)
      Case 6
         result = ZtoB(X1 * Zo, X2 * Zo)
      Case Else   ' Other values.
         result = -99
   End Select
   Im_format_convert = result
End Function
