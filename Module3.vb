Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
' Smith Chart Plotting Module
Const pi = 3.14159265358979
'Dim Pi
'Pi = 4# * Atn(1)  ' Calculate the value of pi.




Function SmithX(code As Integer, Zo As Double, FGHz As Double, ReZ As Double, ImZ As Double, Val1 As Double, Val2 As Double, PtNum As Integer, MaxPt As Integer)
'   DefDbl A-Z
   Dim CW, Valid As Boolean
   Dim ReNew, ImNew, X, X1, Y1, X2, Y2 As Double
   Dim Xctr, Yctr, Radius, Ang1, Ang2 As Double
   
   Valid = True
      
   Select Case code  ' Evaluate Code.
      Case 1
'         MyString = "Series Inductor (nH)"
'          X = 2 * pi * FGHz * Val1
'          ReNew = ReZ
'          ImNew = ImZ + X
'         Calculate center and radius of constant R circle
          Xctr = ((ReZ / Zo) / ((ReZ / Zo) + 1))
          Yctr = 0
          Radius = 1 / ((ReZ / Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZ, ImZ + (2 * pi * FGHz * Val1), Zo)
          Y2 = GammaY(ReZ, ImZ + (2 * pi * FGHz * Val1), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          CW = True
      
      Case 2
'         MyString = "Shunt Inductor (nH)"
'          X = 2 * pi * FGHz * Val1
'          ReNew = ReZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1)
'          ImNew = ImZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1)
'         Calculate center and radius of constant R circle
          MyR = ReZ / (ReZ ^ 2 + ImZ ^ 2)
          Xctr = -((Zo * MyR) / ((Zo * MyR) + 1))
          Yctr = 0
          Radius = 1 / ((MyR * Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1), ImZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1), Zo)
          Y2 = GammaY(ReZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1), ImZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          CW = False
       
      Case 3
'         MyString = "Series Capacitor (pF)"
'          X = -1 / (2 * pi * FGHz * Val1 * 0.001)
'          ReNew = ReZ
'          ImNew = ImZ + X
'         Calculate center and radius of constant R circle
          Xctr = ((ReZ / Zo) / ((ReZ / Zo) + 1))
          Yctr = 0
          Radius = 1 / ((ReZ / Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZ, ImZ - 1 / (2 * pi * FGHz * Val1 * 0.001), Zo)
          Y2 = GammaY(ReZ, ImZ - 1 / (2 * pi * FGHz * Val1 * 0.001), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          CW = False
        
      Case 4
'         MyString = "Shunt Capacitor (pF)"
'          X = -1 / (2 * pi * FGHz * Val1 * 0.001)
'          ReNew = ReZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001))
'          ImNew = ImZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001))
'         Calculate center and radius of constant R circle
          MyR = ReZ / (ReZ ^ 2 + ImZ ^ 2)
          Xctr = -((Zo * MyR) / ((Zo * MyR) + 1))
          Yctr = 0
          Radius = 1 / ((MyR * Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001)), ImZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001)), Zo)
          Y2 = GammaY(ReZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001)), ImZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001)), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          CW = True
                  
      Case 5
'         MyString = "Series Transmission Line (Zo, E.L.)"
'          ReNew = ReTline(Val1, Val2, ReZ, ImZ)
'          ImNew = ImTline(Val1, Val2, ReZ, ImZ)
'         Calculate center and radius of circle
          Xctr = 0
          Yctr = 0
          Radius = GammaMag(ReZ, ImZ, Val1)
'         Calculate locations of start and end points on Smith Chart
'          X1 = GammaX(ReZ, ImZ, Zo)
'          Y1 = GammaY(ReZ, ImZ, Zo)
'          X2 = GammaX(ReTline(Val1, Val2, ReZ, ImZ), ImTline(Val1, Val2, ReZ, ImZ), Zo)
'          Y2 = GammaY(ReTline(Val1, Val2, ReZ, ImZ), ImTline(Val1, Val2, ReZ, ImZ), Zo)
          X1 = GammaX(ReZ, ImZ, Val1)
          Y1 = GammaY(ReZ, ImZ, Val1)
          X2 = GammaX(ReTline(Val1, Val2, ReZ, ImZ), ImTline(Val1, Val2, ReZ, ImZ), Val1)
          Y2 = GammaY(ReTline(Val1, Val2, ReZ, ImZ), ImTline(Val1, Val2, ReZ, ImZ), Val1)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          CW = True
                  
      Case 6
'         MyString = "Open Ckt Stub (Zo, E.L.)"
'          X = ImTline(Val1, Val2, 1E+20, 0)
'          ReNew = ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0))
'          ImNew = ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0))
'         Calculate center and radius of constant R circle
          MyR = ReZ / (ReZ ^ 2 + ImZ ^ 2)
          Xctr = -((Zo * MyR) / ((Zo * MyR) + 1))
          Yctr = 0
          Radius = 1 / ((MyR * Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0)), ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0)), Zo)
          Y2 = GammaY(ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0)), ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0)), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          If (ImTline(Val1, Val2, 1E+20, 0) < 0) Then
             CW = True
          Else
             CW = False
          End If
                 
      Case 7
'         MyString = "Short Ckt Stub (Zo, E.L.)"
'          X = ImTline(Val1, Val2, 0, 0)
'          ReNew = ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0))
'          ImNew = ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0))
'         Calculate center and radius of constant R circle
          MyR = ReZ / (ReZ ^ 2 + ImZ ^ 2)
          Xctr = -((Zo * MyR) / ((Zo * MyR) + 1))
          Yctr = 0
          Radius = 1 / ((MyR * Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0)), ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0)), Zo)
          Y2 = GammaY(ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0)), ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0)), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          If (ImTline(Val1, Val2, 0, 0) < 0) Then
             CW = True
          Else
             CW = False
          End If
          
      Case 8
'         MyString = "Series Resistor (ohms)"
'         Calculate center and radius of constant X circle
          Xctr = 1
          If (ImZ <> 0) Then
             Yctr = Zo / ImZ
             Radius = Abs(Zo / ImZ)
          Else
             Yctr = 1000
             Radius = 1000
          End If
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZ + Val1, ImZ, Zo)
          Y2 = GammaY(ReZ + Val1, ImZ, Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          If (ImZ >= 0) Then
             CW = False
          Else
             CW = True
          End If
          
      Case 9
'         MyString = "Shunt Resistor (ohms)"
'         Calculate center and radius of constant B circle
          Xctr = -1
          If (ImZ <> 0) Then
             Yctr = (ReZ ^ 2 + ImZ ^ 2) / (Zo * ImZ)
             Radius = Abs(Yctr)
          Else
             Yctr = 1000
             Radius = 1000
          End If
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZIIZ(ReZ, ImZ, Val1, 0), ImZIIZ(ReZ, ImZ, Val1, 0), Zo)
          Y2 = GammaY(ReZIIZ(ReZ, ImZ, Val1, 0), ImZIIZ(ReZ, ImZ, Val1, 0), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          If (Yctr >= 0) Then
             CW = True
          Else
             CW = False
          End If
          
      Case Else   ' Other values.
'         MyString = "Not a valid code"
          ReNew = ReZ
          ImNew = ImZ
          Valid = False
          
   End Select
   
   If (Valid) Then
      If (CW) Then
         SmithX = CW_CircleX((Xctr), (Yctr), (Radius), (Ang1), (Ang2), PtNum, MaxPt)
      Else
         SmithX = CCW_CircleX((Xctr), (Yctr), (Radius), (Ang1), (Ang2), PtNum, MaxPt)
      End If
    Else
       SmithX = GammaX(ReZ, ImZ, Zo)
    End If
      
End Function





Function SmithY(code As Integer, Zo As Double, FGHz As Double, ReZ As Double, ImZ As Double, Val1 As Double, Val2 As Double, PtNum As Integer, MaxPt As Integer)
   Dim CW, Valid As Boolean
   Dim ReNew, ImNew, X, X1, Y1, X2, Y2 As Double
   Dim Xctr, Yctr, Radius, Ang1, Ang2 As Double
   
   Valid = True
      
   Select Case code  ' Evaluate Code.
      Case 1
'         MyString = "Series Inductor (nH)"
'          X = 2 * pi * FGHz * Val1
'          ReNew = ReZ
'          ImNew = ImZ + X
'         Calculate center and radius of constant R circle
          Xctr = ((ReZ / Zo) / ((ReZ / Zo) + 1))
          Yctr = 0
          Radius = 1 / ((ReZ / Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZ, ImZ + (2 * pi * FGHz * Val1), Zo)
          Y2 = GammaY(ReZ, ImZ + (2 * pi * FGHz * Val1), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          CW = True
      
      Case 2
'         MyString = "Shunt Inductor (nH)"
'          X = 2 * pi * FGHz * Val1
'          ReNew = ReZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1)
'          ImNew = ImZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1)
'         Calculate center and radius of constant R circle
          MyR = ReZ / (ReZ ^ 2 + ImZ ^ 2)
          Xctr = -((Zo * MyR) / ((Zo * MyR) + 1))
          Yctr = 0
          Radius = 1 / ((MyR * Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1), ImZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1), Zo)
          Y2 = GammaY(ReZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1), ImZIIZ(ReZ, ImZ, 0#, 2 * pi * FGHz * Val1), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          CW = False
     
      Case 3
'         MyString = "Series Capacitor (pF)"
'          X = -1 / (2 * pi * FGHz * Val1 * 0.001)
'          ReNew = ReZ
'          ImNew = ImZ + X
'         Calculate center and radius of constant R circle
          Xctr = ((ReZ / Zo) / ((ReZ / Zo) + 1))
          Yctr = 0
          Radius = 1 / ((ReZ / Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZ, ImZ - 1 / (2 * pi * FGHz * Val1 * 0.001), Zo)
          Y2 = GammaY(ReZ, ImZ - 1 / (2 * pi * FGHz * Val1 * 0.001), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          CW = False
        
      Case 4
'         MyString = "Shunt Capacitor (pF)"
'          X = -1 / (2 * pi * FGHz * Val1 * 0.001)
'          ReNew = ReZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001))
'          ImNew = ImZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001))
'         Calculate center and radius of constant R circle
          MyR = ReZ / (ReZ ^ 2 + ImZ ^ 2)
          Xctr = -((Zo * MyR) / ((Zo * MyR) + 1))
          Yctr = 0
          Radius = 1 / ((MyR * Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001)), ImZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001)), Zo)
          Y2 = GammaY(ReZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001)), ImZIIZ(ReZ, ImZ, 0#, -1 / (2 * pi * FGHz * Val1 * 0.001)), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          CW = True
                  
      Case 5
'         MyString = "Series Transmission Line (Zo, E.L.)"
'          ReNew = ReTline(Val1, Val2, ReZ, ImZ)
'          ImNew = ImTline(Val1, Val2, ReZ, ImZ)
'         Calculate center and radius of circle
          Xctr = 0
          Yctr = 0
          Radius = GammaMag(ReZ, ImZ, Val1)
'         Calculate locations of start and end points on Smith Chart
'          X1 = GammaX(ReZ, ImZ, Zo)
'          Y1 = GammaY(ReZ, ImZ, Zo)
'          X2 = GammaX(ReTline(Val1, Val2, ReZ, ImZ), ImTline(Val1, Val2, ReZ, ImZ), Zo)
'          Y2 = GammaY(ReTline(Val1, Val2, ReZ, ImZ), ImTline(Val1, Val2, ReZ, ImZ), Zo)
          X1 = GammaX(ReZ, ImZ, Val1)
          Y1 = GammaY(ReZ, ImZ, Val1)
          X2 = GammaX(ReTline(Val1, Val2, ReZ, ImZ), ImTline(Val1, Val2, ReZ, ImZ), Val1)
          Y2 = GammaY(ReTline(Val1, Val2, ReZ, ImZ), ImTline(Val1, Val2, ReZ, ImZ), Val1)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          CW = True
                  
      Case 6
'         MyString = "Open Ckt Stub (Zo, E.L.)"
'          X = ImTline(Val1, Val2, 1E+20, 0)
'          ReNew = ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0))
'          ImNew = ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0))
'         Calculate center and radius of constant R circle
          MyR = ReZ / (ReZ ^ 2 + ImZ ^ 2)
          Xctr = -((Zo * MyR) / ((Zo * MyR) + 1))
          Yctr = 0
          Radius = 1 / ((MyR * Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0)), ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0)), Zo)
          Y2 = GammaY(ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0)), ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 1E+20, 0)), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          If (ImTline(Val1, Val2, 1E+20, 0) < 0) Then
             CW = True
          Else
             CW = False
          End If
                 
      Case 7
'         MyString = "Short Ckt Stub (Zo, E.L.)"
'          X = ImTline(Val1, Val2, 0, 0)
'          ReNew = ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0))
'          ImNew = ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0))
'         Calculate center and radius of constant R circle
          MyR = ReZ / (ReZ ^ 2 + ImZ ^ 2)
          Xctr = -((Zo * MyR) / ((Zo * MyR) + 1))
          Yctr = 0
          Radius = 1 / ((MyR * Zo) + 1)
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0)), ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0)), Zo)
          Y2 = GammaY(ReZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0)), ImZIIZ(ReZ, ImZ, 0#, ImTline(Val1, Val2, 0, 0)), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          If (ImTline(Val1, Val2, 0, 0) < 0) Then
             CW = True
          Else
             CW = False
          End If
          
      Case 8
'         MyString = "Series Resistor (ohms)"
'         Calculate center and radius of constant X circle
          Xctr = 1
          If (ImZ <> 0) Then
             Yctr = Zo / ImZ
             Radius = Abs(Zo / ImZ)
          Else
             Yctr = 1000
             Radius = 1000
          End If
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZ + Val1, ImZ, Zo)
          Y2 = GammaY(ReZ + Val1, ImZ, Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          If (ImZ >= 0) Then
             CW = False
          Else
             CW = True
          End If
          
      Case 9
'         MyString = "Shunt Resistor (ohms)"
'         Calculate center and radius of constant B circle
          Xctr = -1
          If (ImZ <> 0) Then
             Yctr = (ReZ ^ 2 + ImZ ^ 2) / (Zo * ImZ)
             Radius = Abs(Yctr)
          Else
             Yctr = 1000
             Radius = 1000
          End If
'         Calculate locations of start and end points on Smith Chart
          X1 = GammaX(ReZ, ImZ, Zo)
          Y1 = GammaY(ReZ, ImZ, Zo)
          X2 = GammaX(ReZIIZ(ReZ, ImZ, Val1, 0), ImZIIZ(ReZ, ImZ, Val1, 0), Zo)
          Y2 = GammaY(ReZIIZ(ReZ, ImZ, Val1, 0), ImZIIZ(ReZ, ImZ, Val1, 0), Zo)
'         Now determine angles relative to center of constant-R circle
          Ang1 = Application.Atan2(X1 - Xctr, Y1 - Yctr) * 180 / pi
          Ang2 = Application.Atan2(X2 - Xctr, Y2 - Yctr) * 180 / pi
'         Now Specify direction of the circle
          If (Yctr >= 0) Then
             CW = True
          Else
             CW = False
          End If
          
      Case Else   ' Other values.
'         MyString = "Not a valid code"
          ReNew = ReZ
          ImNew = ImZ
          Valid = False
        
      End Select
             
   If (Valid) Then
      If (CW) Then
         SmithY = CW_CircleY((Xctr), (Yctr), (Radius), (Ang1), (Ang2), PtNum, MaxPt)
      Else
         SmithY = CCW_CircleY((Xctr), (Yctr), (Radius), (Ang1), (Ang2), PtNum, MaxPt)
      End If
    Else
       SmithY = GammaY(ReZ, ImZ, Zo)
    End If
    
End Function




Function GammaX(R As Double, X As Double, Z0 As Double) As Double
'  convert impedance to reflection coefficient
   ReG = ((R - Z0) * (R + Z0) + X ^ 2) / ((R + Z0) ^ 2 + X ^ 2)
   ImG = (2 * X * Z0) / ((R + Z0) ^ 2 + X ^ 2)
'   MagG = (ReG ^ 2 + ImG ^ 2) ^ 0.5
'   AngG = Atan2(ReG, ImG) * 180 / Pi
   GammaX = ReG
End Function

Function GammaY(R As Double, X As Double, Z0 As Double) As Double
'  convert impedance to reflection coefficient
   ReG = ((R - Z0) * (R + Z0) + X ^ 2) / ((R + Z0) ^ 2 + X ^ 2)
   ImG = (2 * X * Z0) / ((R + Z0) ^ 2 + X ^ 2)
'   MagG = (ReG ^ 2 + ImG ^ 2) ^ 0.5
'   AngG = Atan2(ReG, ImG) * 180 / Pi
   GammaY = ImG
End Function

Function GammaMag(R As Double, X As Double, Z0 As Double) As Double
'   Pi = 3.14159265358979
'  convert impedance to reflection coefficient
   ReG = ((R - Z0) * (R + Z0) + X ^ 2) / ((R + Z0) ^ 2 + X ^ 2)
   ImG = (2 * X * Z0) / ((R + Z0) ^ 2 + X ^ 2)
   MagG = (ReG ^ 2 + ImG ^ 2) ^ 0.5
   AngG = Application.Atan2(ReG, ImG) * 180 / pi
   GammaMag = MagG
End Function

Function GammaAng(R As Double, X As Double, Z0 As Double) As Double
'   Pi = 3.14159265358979
'  convert impedance to reflection coefficient
   ReG = ((R - Z0) * (R + Z0) + X ^ 2) / ((R + Z0) ^ 2 + X ^ 2)
   ImG = (2 * X * Z0) / ((R + Z0) ^ 2 + X ^ 2)
   MagG = (ReG ^ 2 + ImG ^ 2) ^ 0.5
   AngG = Application.Atan2(ReG, ImG) * 180 / pi
   GammaAng = AngG
End Function


Function CCW_CircleX(Cx As Double, Cy As Double, Radius As Double, StartAng As Double, StopAng As Double, PtNum As Integer, MaxPt As Integer)
'  counterclockwise motion from start to stop angles.
'  StartAng and StopAng are in degrees
'  Angle 0° is at 3 o'clock, 90° is at 12 o'clock, 180° is at 9 o'clock, 270° is at 6 o'clock
   P = PtNum
   If (PtNum < 1) Then P = 1
   If (PtNum > MaxPt) Then P = MaxPt
   If (StopAng > StartAng) Then
      MyAng = StartAng + (P - 1) * (StopAng - StartAng) / (MaxPt - 1)
   Else
      MyAng = StartAng + (P - 1) * (StopAng + 360 - StartAng) / (MaxPt - 1)
   End If
   MyX = Cx + Radius * Cos(MyAng * pi / 180)
   MyY = Cy + Radius * Sin(MyAng * pi / 180)
   CCW_CircleX = MyX
End Function

Function CCW_CircleY(Cx As Double, Cy As Double, Radius As Double, StartAng As Double, StopAng As Double, PtNum As Integer, MaxPt As Integer)
'  counterclockwise motion from start to stop angles.
'  StartAng and StopAng are in degrees
'  Angle 0° is at 3 o'clock, 90° is at 12 o'clock, 180° is at 9 o'clock, 270° is at 6 o'clock
   P = PtNum
   If (PtNum < 1) Then P = 1
   If (PtNum > MaxPt) Then P = MaxPt
   If (StopAng > StartAng) Then
      MyAng = StartAng + (P - 1) * (StopAng - StartAng) / (MaxPt - 1)
   Else
      MyAng = StartAng + (P - 1) * (StopAng + 360 - StartAng) / (MaxPt - 1)
   End If
   MyX = Cx + Radius * Cos(MyAng * pi / 180)
   MyY = Cy + Radius * Sin(MyAng * pi / 180)
   CCW_CircleY = MyY
End Function

Function CW_CircleX(Cx As Double, Cy As Double, Radius As Double, StartAng As Double, StopAng As Double, PtNum As Integer, MaxPt As Integer)
'  clockwise motion from start to stop angles.
'  StartAng and StopAng are in degrees
'  Angle 0° is at 3 o'clock, 90° is at 12 o'clock, 180° is at 9 o'clock, 270° is at 6 o'clock
   P = PtNum
   If (PtNum < 1) Then P = 1
   If (PtNum > MaxPt) Then P = MaxPt
   If (StopAng < StartAng) Then
      MyAng = StartAng - (P - 1) * (StartAng - StopAng) / (MaxPt - 1)
   Else
      MyAng = StartAng - (P - 1) * (StartAng + 360 - StopAng) / (MaxPt - 1)
   End If
   MyX = Cx + Radius * Cos(MyAng * pi / 180)
   MyY = Cy + Radius * Sin(MyAng * pi / 180)
   CW_CircleX = MyX
End Function

Function CW_CircleY(Cx As Double, Cy As Double, Radius As Double, StartAng As Double, StopAng As Double, PtNum As Integer, MaxPt As Integer)
'  clockwise motion from start to stop angles.
'  StartAng and StopAng are in degrees
'  Angle 0° is at 3 o'clock, 90° is at 12 o'clock, 180° is at 9 o'clock, 270° is at 6 o'clock
   P = PtNum
   If (PtNum < 1) Then P = 1
   If (PtNum > MaxPt) Then P = MaxPt
   If (StopAng < StartAng) Then
      MyAng = StartAng - (P - 1) * (StartAng - StopAng) / (MaxPt - 1)
   Else
      MyAng = StartAng - (P - 1) * (StartAng + 360 - StopAng) / (MaxPt - 1)
   End If
   MyX = Cx + Radius * Cos(MyAng * pi / 180)
   MyY = Cy + Radius * Sin(MyAng * pi / 180)
   CW_CircleY = MyY
End Function


