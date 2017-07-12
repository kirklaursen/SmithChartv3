Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Rem
Rem      Functions written by Kirk Laursen
Rem        March 18, 1999
Rem
Rem

DefDbl A-Z
Rem
Rem    ============================================================================
Rem    A set of functions to convert between impedance and reflection coefficient
Rem    ============================================================================
Rem
Rem   function to calculate real part of reflection coefficient from impedance
Function ZtoReG(R As Double, X As Double, Z0 As Double)
   ZtoReG = ((R - Z0) * (R + Z0) + X ^ 2) / ((R + Z0) ^ 2 + X ^ 2)
End Function
 
Rem   function to calculate imaginary part of reflection coefficient from impedance
Function ZtoImG(R As Double, X As Double, Z0 As Double)
   ZtoImG = (2 * X * Z0) / ((R + Z0) ^ 2 + X ^ 2)
End Function
 
Rem   function to calculate real part of impedance from reflection coefficient
Function GtoReZ(U As Double, V As Double, Z0 As Double)
   GtoReZ = Z0 * ((1 - U ^ 2 - V ^ 2) / ((1 - U) ^ 2 + V ^ 2))
End Function
 
Rem   function to calculate real part of impedance from reflection coefficient
Function GtoImZ(U As Double, V As Double, Z0 As Double)
   GtoImZ = Z0 * ((2 * V) / ((1 - U) ^ 2 + V ^ 2))
End Function

Rem   function to invert a complex number (i.e. admittance <---> impedance)
Rem                  1
Rem     G + jB = ---------
Rem                R + jX
Function ZtoG(R, X)
   ZtoG = R / (R ^ 2 + X ^ 2)
End Function

Rem   function to invert a complex number (i.e. admittance <---> impedance)
Rem                  1
Rem     G + jB = ---------
Rem                R + jX
Function ZtoB(R, X)
   ZtoB = -X / (R ^ 2 + X ^ 2)
End Function

Rem   function to convert imaginary impedance to a capacitance or inductance value
Function XtoLC(X As Double, F_GHz As Double) As String
   If (X > 0) Then
      L_nH = X / (2 * 3.14159265358979 * F_GHz)
      result$ = Str(Format(L_nH, "0.000")) & " nH Inductor at " & Str(Format(F_GHz, "0.000")) & " GHz"
   End If
   If (X < 0) Then
      C_pF = -1000 / (2 * 3.14159265358979 * F_GHz * X)
      result$ = Str(Format(C_pF, "0.000")) & " pF Capacitor at " & Str(Format(F_GHz, "0.000")) & " GHz"
   End If
   If (X = 0) Then
      result$ = "No reactive component necessary"
   End If
   XtoLC = result$
End Function

Rem  function to convert complex numbers to text format
Function Ztext(R As Double, X As Double) As String
   If (X < 0) Then
      result$ = Str(Format(R, "0.000")) & " - j" & Str(Format(Abs(X), "0.000"))
   Else
      result$ = Str(Format(R, "0.000")) & " + j" & Str(Format(Abs(X), "0.000"))
   End If
   Ztext = result$
End Function
      

Rem function to combine impedances in parallel
Function ReZIIZ(R1, X1, R2, X2)
   g = ZtoG(R1, X1) + ZtoG(R2, X2)
   b = ZtoB(R1, X1) + ZtoB(R2, X2)
   ReZIIZ = ZtoG(g, b)
End Function

Rem function to combine impedances in parallel
Function ImZIIZ(R1, X1, R2, X2)
   g = ZtoG(R1, X1) + ZtoG(R2, X2)
   b = ZtoB(R1, X1) + ZtoB(R2, X2)
   ImZIIZ = ZtoB(g, b)
End Function

Rem function to combine R (ohms), L (nH), and C (pF) in parallel at a frequency F (GHz)
Function ReRLC(F As Double, R As Double, L As Double, C As Double)
   If (L <= 0) Then LnH = 9E+20 Else LnH = L * 0.000000001
   If (C <= 0) Then CpF = 0 Else CpF = C * 0.000000000001
   If (R <= 0) Then Rohm = 9E+20 Else Rohm = R
   If (F <= 0) Then FGHz = 1 Else FGHz = F * 1000000000#
   w1 = 2 * 3.14159265358979 * FGHz
   realpart = 1 / Rohm
   imagpart = w1 * CpF - (1 / (w1 * LnH))
   If (R <= 0) And (L <= 0) And (C <= 0) Then
      g = 0
   Else
      g = ZtoG(realpart, imagpart)
   End If
   ReRLC = g
End Function


Rem function to combine R (ohms), L (nH), and C (pF) in parallel at a frequency F (GHz)
Function ImRLC(F, R, L, C)
   If (L <= 0) Then LnH = 9E+20 Else LnH = L * 0.000000001
   If (C <= 0) Then CpF = 0 Else CpF = C * 0.000000000001
   If (R <= 0) Then Rohm = 9E+20 Else Rohm = R
   If (F <= 0) Then FGHz = 1 Else FGHz = F * 1000000000#
   w1 = 2 * 3.14159265358979 * FGHz
   realpart = 1 / Rohm
   imagpart = w1 * CpF - (1 / (w1 * LnH))
   If (R <= 0) And (L <= 0) And (C <= 0) Then
      b = 0
   Else
      b = ZtoB(realpart, imagpart)
   End If
   ImRLC = b
End Function

Rem function to transform impedance through a transmission line
Function ReTline(Z0, DegLen, RL, XL)
   ReNum = Z0 * RL
   ImNum = Z0 * (XL + Z0 * Tan(DegLen * 3.14159265358979 / 180))
   ReDen = (Z0 - XL * Tan(DegLen * 3.14159265358979 / 180))
   ImDen = RL * Tan(DegLen * 3.14159265358979 / 180)
   ReTline = (ReNum * ReDen + ImNum * ImDen) / (ReDen ^ 2 + ImDen ^ 2)
End Function

Rem function to transform impedance through a transmission line
Function ImTline(Z0, DegLen, RL, XL)
   ReNum = Z0 * RL
   ImNum = Z0 * (XL + Z0 * Tan(DegLen * 3.14159265358979 / 180))
   ReDen = (Z0 - XL * Tan(DegLen * 3.14159265358979 / 180))
   ImDen = RL * Tan(DegLen * 3.14159265358979 / 180)
   ImTline = (ReDen * ImNum - ImDen * ReNum) / (ReDen ^ 2 + ImDen ^ 2)
End Function

Rem
Rem
Rem    ============================================================================
Rem    A set of functions related to amplifier stability
Rem    ============================================================================
Rem
Rem    Function to determine stability status of an s-matrix
Rem      Based on Maas, Nonlinear Microwave Circuits, pp 324-325
Function Stability(MS11, AS11, MS21, AS21, MS12, AS12, MS22, AS22)
   ReS21S12 = (MS21 * MS12) * Cos((AS21 + AS12) * 3.14159265358979 / 180)
   ImS21S12 = (MS21 * MS12) * Sin((AS21 + AS12) * 3.14159265358979 / 180)
   MagS21S12 = (ReS21S12 ^ 2 + ImS21S12 ^ 2) ^ 0.5
   ReS11S22 = (MS11 * MS22) * Cos((AS11 + AS22) * 3.14159265358979 / 180)
   ImS11S22 = (MS11 * MS22) * Sin((AS11 + AS22) * 3.14159265358979 / 180)
   ' calculate the determinant of the S-matrix
   ReDetS = ReS11S22 - ReS21S12
   ImDetS = ImS11S22 - ImS21S12
   MagDetS = (ReDetS ^ 2 + ImDetS ^ 2) ^ 0.5
   ' calculate the Linville stability factor, K
   K = (1 - MS11 ^ 2 - MS22 ^ 2 + MagDetS ^ 2) / (2 * MagS21S12)
   
   If (K > 1) And (MagDetS < 1) Then
      result$ = "Unconditionally Stable"
   Else
      result$ = "Potentially Unstable"
   End If
   Stability = result$
End Function


Rem    Function to determine Linville stability factor of an S-matrix
Rem      Based on Maas, Nonlinear Microwave Circuits, pp 324-325
Function Linville(MS11, AS11, MS21, AS21, MS12, AS12, MS22, AS22)
   ReS21S12 = (MS21 * MS12) * Cos((AS21 + AS12) * 3.14159265358979 / 180)
   ImS21S12 = (MS21 * MS12) * Sin((AS21 + AS12) * 3.14159265358979 / 180)
   MagS21S12 = (ReS21S12 ^ 2 + ImS21S12 ^ 2) ^ 0.5
   ReS11S22 = (MS11 * MS22) * Cos((AS11 + AS22) * 3.14159265358979 / 180)
   ImS11S22 = (MS11 * MS22) * Sin((AS11 + AS22) * 3.14159265358979 / 180)
   ' calculate the determinant of the S-matrix
   ReDetS = ReS11S22 - ReS21S12
   ImDetS = ImS11S22 - ImS21S12
   MagDetS = (ReDetS ^ 2 + ImDetS ^ 2) ^ 0.5
   ' calculate the Linville stability factor, K
   K = (1 - MS11 ^ 2 - MS22 ^ 2 + MagDetS ^ 2) / (2 * MagS21S12)
   
   Linville = K
End Function


Rem    Function to calculate the magnitude of the determinant of an S-matrix
Rem      Based on Maas, Nonlinear Microwave Circuits, pp 324-325
Function MagDeterminant(MS11, AS11, MS21, AS21, MS12, AS12, MS22, AS22)
   ReS21S12 = (MS21 * MS12) * Cos((AS21 + AS12) * 3.14159265358979 / 180)
   ImS21S12 = (MS21 * MS12) * Sin((AS21 + AS12) * 3.14159265358979 / 180)
   MagS21S12 = (ReS21S12 ^ 2 + ImS21S12 ^ 2) ^ 0.5
   ReS11S22 = (MS11 * MS22) * Cos((AS11 + AS22) * 3.14159265358979 / 180)
   ImS11S22 = (MS11 * MS22) * Sin((AS11 + AS22) * 3.14159265358979 / 180)
   ' calculate the determinant of the S-matrix
   ReDetS = ReS11S22 - ReS21S12
   ImDetS = ImS11S22 - ImS21S12
   MagDeterminant = (ReDetS ^ 2 + ImDetS ^ 2) ^ 0.5
End Function
Rem
Rem    ============================================================================
Rem    ============================================================================
Rem

Rem    Function to Convert Degrees into Radians
Function DegtoRad(fDegrees As Double)
   DegtoRad = fDegrees * 3.14159265358979 / 180
End Function

'    Function to Convert Radians to Degrees
Function RadtoDeg(fRadians As Double)
   RadtoDeg = fRadians * 180 / 3.14159265358979
End Function

'
' ======================================================================================
'

'    Function to return the Real part of a complex multiplication
Function ReCxMult(ReX As Double, ImX As Double, ReY As Double, ImY As Double)
   ReCxMult = ReX * ReY - ImX * ImY
End Function
'    Function to return the Imaginary part of a complex multiplication
Function ImCxMult(ReX As Double, ImX As Double, ReY As Double, ImY As Double)
   ImCxMult = ReY * ImX + ImY * ReX
End Function

'    Function to return the Real part of a complex division (ReX +j ImX) / (ReY + j ImY)
Function ReCxDiv(ReX As Double, ImX As Double, ReY As Double, ImY As Double)
   ReCxDiv = (ReX * ReY + ImX * ImY) / (ReY ^ 2 + ImY ^ 2)
End Function
'    Function to return the Imaginary part of a complex division (ReX +j ImX) / (ReY + j ImY)
Function ImCxDiv(ReX As Double, ImX As Double, ReY As Double, ImY As Double)
   ImCxDiv = (ReY * ImX - ImY * ReX) / (ReY ^ 2 + ImY ^ 2)
End Function

Function MagCx(ReX, ImX)
   MagCx = (ReX ^ 2 + ImX ^ 2) ^ 0.5
End Function

'
' ======================================================================================
'
'    Functions to calculate the determinant of the S matrix
'
Function ReDeltaS(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' calculate the determinant of the S matrix (denoted Delta(s))
   ReDeltaS = ReCxMult(ReS11, ImS11, ReS22, ImS22) - ReCxMult(ReS21, ImS21, ReS12, ImS12)
End Function
'    Function to calculate the determinant of the S matrix
Function ImDeltaS(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' calculate the determinant of the S matrix (denoted Delta(s))
   ImDeltaS = ImCxMult(ReS11, ImS11, ReS22, ImS22) - ImCxMult(ReS21, ImS21, ReS12, ImS12)
End Function
'    Function to calculate the determinant of the S matrix
Function MagDeltaS(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' calculate the determinant of the S matrix (denoted Delta(s))
   RealDeltaS = ReCxMult(ReS11, ImS11, ReS22, ImS22) - ReCxMult(ReS21, ImS21, ReS12, ImS12)
   ImagDeltaS = ImCxMult(ReS11, ImS11, ReS22, ImS22) - ImCxMult(ReS21, ImS21, ReS12, ImS12)
   MagDeltaS = MagCx(RealDeltaS, ImagDeltaS)
End Function
'
' ======================================================================================
'
'    Function to calculate the Rollett Stability Factor, K
'      refer to Hayward, Radio Frequency Design, pp. 196
'            or Maas, Nonlinear Microwave Circuits, (c)1997, pp 325
'
'      device is unconditionally stable if K > +1
Function Rollett(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   MagS11 = MagCx(ReS11, ImS11)
   MagS22 = MagCx(ReS22, ImS22)
   ' calculate the value denoted delta(s)
'   ReDeltaS = (ReS11 * ReS22 - ImS11 * ImS22) - (ReS21 * ReS12 - ImS21 * ImS12)
'   ImDeltaS = (ImS11 * ReS22 + ReS11 * ImS22) - (ImS21 * ReS12 - ReS21 * ImS12)
   MagDelS = MagDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ' calculate the value of S12*S21
   ReS12S21 = ReCxMult(ReS21, ImS21, ReS12, ImS12)
   ImS12S21 = ImCxMult(ReS21, ImS21, ReS12, ImS12)
   MagS12S21 = MagCx(ReS12S21, ImS12S21)
   Rollett = (1 - MagS11 ^ 2 - MagS22 ^ 2 + MagDelS ^ 2) / (2 * MagS12S21)
End Function
'
' ======================================================================================
'
'
'
'   Function to calculate input reflection coefficient of a network
'    ReGL + j ImGL is the complex reflection coefficient at the load side of the network
'     Based on Maas, Nonlinear Microwave Circuits, (c)1997, pp 325
'
Function ReGammaIn(ReGL As Double, ImGL As Double, ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' First multiply S12*S21
   ReS1221 = ReCxMult(ReS12, ImS12, ReS21, ImS21)
   ImS1221 = ImCxMult(ReS12, ImS12, ReS21, ImS21)
   ' Multiply the Result by GL to get the numerator
   ReNum = ReCxMult(ReS1221, ImS1221, ReGL, ImGL)
   ImNum = ImCxMult(ReS1221, ImS1221, ReGL, ImGL)
   ' Calculate the denominator = (1 - S22*GL)
   ReDen = 1 - ReCxMult(ReS22, ImS22, ReGL, ImGL)
   ImDen = -ImCxMult(ReS22, ImS22, ReGL, ImGL)
   ' Now calculate the fraction
   ReFrac = ReCxDiv(ReNum, ImNum, ReDen, ImDen)
   ImFrac = ImCxDiv(ReNum, ImNum, ReDen, ImDen)
   ' And finally the end result
   ReGammaIn = ReS11 + ReFrac
'   ImGammaIn = ImS11 + ImFrac
End Function
'   Function to calculate input reflection coefficient of a network
'    ReGL + j ImGL is the complex reflection coefficient at the load side of the network
Function ImGammaIn(ReGL As Double, ImGL As Double, ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' First multiply S12*S21
   ReS1221 = ReCxMult(ReS12, ImS12, ReS21, ImS21)
   ImS1221 = ImCxMult(ReS12, ImS12, ReS21, ImS21)
   ' Multiply the Result by GL to get the numerator
   ReNum = ReCxMult(ReS1221, ImS1221, ReGL, ImGL)
   ImNum = ImCxMult(ReS1221, ImS1221, ReGL, ImGL)
   ' Calculate the denominator = (1 - S22*GL)
   ReDen = 1 - ReCxMult(ReS22, ImS22, ReGL, ImGL)
   ImDen = -ImCxMult(ReS22, ImS22, ReGL, ImGL)
   ' Now calculate the fraction
   ReFrac = ReCxDiv(ReNum, ImNum, ReDen, ImDen)
   ImFrac = ImCxDiv(ReNum, ImNum, ReDen, ImDen)
   ' And finally the end result
'   ReGammaIn = ReS11 + ReFrac
   ImGammaIn = ImS11 + ImFrac
End Function
'
' ======================================================================================
'
'
'
'   Function to calculate output reflection coefficient of a network
'    ReGL + j ImGL is the complex reflection coefficient at the load side of the network
'     Based on Maas, Nonlinear Microwave Circuits, (c)1997, pp 325
'
Function ReGammaOut(ReGS As Double, ImGS As Double, ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' First multiply S12*S21
   ReS1221 = ReCxMult(ReS12, ImS12, ReS21, ImS21)
   ImS1221 = ImCxMult(ReS12, ImS12, ReS21, ImS21)
   ' Multiply the Result by GS to get the numerator
   ReNum = ReCxMult(ReS1221, ImS1221, ReGS, ImGS)
   ImNum = ImCxMult(ReS1221, ImS1221, ReGS, ImGS)
   ' Calculate the denominator = (1 - S11*GS)
   ReDen = 1 - ReCxMult(ReS11, ImS11, ReGS, ImGS)
   ImDen = -ImCxMult(ReS11, ImS11, ReGS, ImGS)
   ' Now calculate the fraction
   ReFrac = ReCxDiv(ReNum, ImNum, ReDen, ImDen)
   ImFrac = ImCxDiv(ReNum, ImNum, ReDen, ImDen)
   ' And finally the end result
   ReGammaOut = ReS22 + ReFrac
'   ImGammaOut = ImS22 + ImFrac
End Function
'   Function to calculate input reflection coefficient of a network
'    ReGL + j ImGL is the complex reflection coefficient at the load side of the network
Function ImGammaOut(ReGS As Double, ImGS As Double, ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' First multiply S12*S21
   ReS1221 = ReCxMult(ReS12, ImS12, ReS21, ImS21)
   ImS1221 = ImCxMult(ReS12, ImS12, ReS21, ImS21)
   ' Multiply the Result by GS to get the numerator
   ReNum = ReCxMult(ReS1221, ImS1221, ReGS, ImGS)
   ImNum = ImCxMult(ReS1221, ImS1221, ReGS, ImGS)
   ' Calculate the denominator = (1 - S11*GS)
   ReDen = 1 - ReCxMult(ReS11, ImS11, ReGS, ImGS)
   ImDen = -ImCxMult(ReS11, ImS11, ReGS, ImGS)
   ' Now calculate the fraction
   ReFrac = ReCxDiv(ReNum, ImNum, ReDen, ImDen)
   ImFrac = ImCxDiv(ReNum, ImNum, ReDen, ImDen)
   ' And finally the end result
'   ReGammaOut = ReS22 + ReFrac
   ImGammaOut = ImS22 + ImFrac
End Function
'
' ======================================================================================
'
'     Based on Maas, Nonlinear Microwave Circuits, (c)1997, pp 324-326
'
Function TransducerGain(ReGS As Double, ImGS As Double, ReGL As Double, ImGL As Double, ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   MagS21 = MagCx(ReS21, ImS21)
   MagGS = MagCx(ReGS, ImGS)
   MagGL = MagCx(ReGL, ImGL)
   numer = MagS21 ^ 2 * (1 - MagGS ^ 2) * (1 - MagGL ^ 2)
   
   ReSrce = 1 - ReCxMult(ReS11, ImS11, ReGS, ImGS)
   ImSrce = -ImCxMult(ReS11, ImS11, ReGS, ImGS)
   ReLd = 1 - ReCxMult(ReS22, ImS22, ReGL, ImGL)
   ImLd = -ImCxMult(ReS22, ImS22, ReGL, ImGL)
   Re2 = ReCxMult(ReSrce, ImSrce, ReLd, ImLd)
   Im2 = ImCxMult(ReSrce, ImSrce, ReLd, ImLd)
   
   ReA = ReCxMult(ReS21, ImS21, ReS12, ImS12)
   ImA = ImCxMult(ReS21, ImS21, ReS12, ImS12)
   ReB = ReCxMult(ReGL, ImGL, ReGS, ImGS)
   ImB = ImCxMult(ReGL, ImGL, ReGS, ImGS)
   Re3 = ReCxMult(ReA, ImA, ReB, ImB)
   Im3 = ImCxMult(ReA, ImA, ReB, ImB)
   
   ReDen = Re2 - Re3
   ImDen = Im2 - Im3
   Magden = MagCx(ReDen, ImDen)

   TransducerGain = numer / (Magden ^ 2)
End Function

'
' ======================================================================================
'
'  Calculation for the output stability circle
'     Based on Maas, Nonlinear Microwave Circuits, (c)1997, pp 324-326
'
Function ReOutputStabCtr(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ReNum = ReS22 - ReCxMult(ReDs, ImDs, ReS11, -ImS11)
   ImNum = -(ImS22 - ImCxMult(ReDs, ImDs, ReS11, -ImS11))
   Den = MagCx(ReS22, ImS22) ^ 2 - MagCx(ReDs, ImDs) ^ 2
   ReCL = ReNum / Den      ' the real part of the center of the output stability circle
'   ImCL = ImNum / Den      ' the imaginary part of the center of the output stability circle
   ReOutputStabCtr = ReCL
'   ReNum2 = ReCxMult(ReS12, ImS12, ReS21, ImS21)
'   Imnum2 = ImCxMult(ReS12, ImS12, ReS21, ImS21)
'   radius = MagCx(ReNum2 / Den, Imnum2 / Den)
End Function
Function ImOutputStabCtr(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ReNum = ReS22 - ReCxMult(ReDs, ImDs, ReS11, -ImS11)
   ImNum = -(ImS22 - ImCxMult(ReDs, ImDs, ReS11, -ImS11))
   Den = MagCx(ReS22, ImS22) ^ 2 - MagCx(ReDs, ImDs) ^ 2
'   ReCL = ReNum / Den      ' the real part of the center of the output stability circle
   ImCL = ImNum / Den      ' the imaginary part of the center of the output stability circle
   ImOutputStabCtr = ImCL
'   ReNum2 = ReCxMult(ReS12, ImS12, ReS21, ImS21)
'   Imnum2 = ImCxMult(ReS12, ImS12, ReS21, ImS21)
'   radius = MagCx(ReNum2 / Den, Imnum2 / Den)
End Function
Function OutputStabRad(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
'   ReNum = ReS22 - ReCxMult(ReDS, ImDS, ReS11, -ImS11)
'   ImNum = -(ImS22 - ImCxMult(ReDS, ImDS, ReS11, -ImS11))
   Den = MagCx(ReS22, ImS22) ^ 2 - MagCx(ReDs, ImDs) ^ 2
'   ReCL = ReNum / Den      ' the real part of the center of the output stability circle
'   ImCL = ImNum / Den      ' the imaginary part of the center of the output stability circle
   ReNum2 = ReCxMult(ReS12, ImS12, ReS21, ImS21)
   ImNum2 = ImCxMult(ReS12, ImS12, ReS21, ImS21)
   Radius = MagCx(ReNum2 / Den, ImNum2 / Den)
   OutputStabRad = Radius
End Function


'
' ======================================================================================
'
'  Calculation for the input stability circle
'     Based on Maas, Nonlinear Microwave Circuits, (c)1997, pp 324-326
'
Function ReInputStabCtr(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ReNum = ReS11 - ReCxMult(ReDs, ImDs, ReS22, -ImS22)
   ImNum = -(ImS11 - ImCxMult(ReDs, ImDs, ReS22, -ImS22))
   Den = MagCx(ReS11, ImS11) ^ 2 - MagCx(ReDs, ImDs) ^ 2
   ReCL = ReNum / Den      ' the real part of the center of the input stability circle
'   ImCL = ImNum / Den      ' the imaginary part of the center of the input stability circle
   ReInputStabCtr = ReCL
'   ReNum2 = ReCxMult(ReS12, ImS12, ReS21, ImS21)
'   Imnum2 = ImCxMult(ReS12, ImS12, ReS21, ImS21)
'   radius = MagCx(ReNum2 / Den, Imnum2 / Den)
End Function
Function ImInputStabCtr(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ReNum = ReS11 - ReCxMult(ReDs, ImDs, ReS22, -ImS22)
   ImNum = -(ImS11 - ImCxMult(ReDs, ImDs, ReS22, -ImS22))
   Den = MagCx(ReS11, ImS11) ^ 2 - MagCx(ReDs, ImDs) ^ 2
'   ReCL = ReNum / Den      ' the real part of the center of the input stability circle
   ImCL = ImNum / Den      ' the imaginary part of the center of the input stability circle
   ImInputStabCtr = ImCL
'   ReNum2 = ReCxMult(ReS12, ImS12, ReS21, ImS21)
'   Imnum2 = ImCxMult(ReS12, ImS12, ReS21, ImS21)
'   radius = MagCx(ReNum2 / Den, Imnum2 / Den)
End Function
Function InputStabRad(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
'   ReNum = ReS11 - ReCxMult(ReDS, ImDS, ReS22, -ImS22)
'   ImNum = -(ImS11 - ImCxMult(ReDS, ImDS, ReS22, -ImS22))
   Den = MagCx(ReS11, ImS11) ^ 2 - MagCx(ReDs, ImDs) ^ 2
'   ReCL = ReNum / Den      ' the real part of the center of the input stability circle
'   ImCL = ImNum / Den      ' the imaginary part of the center of the input stability circle
   ReNum2 = ReCxMult(ReS12, ImS12, ReS21, ImS21)
   ImNum2 = ImCxMult(ReS12, ImS12, ReS21, ImS21)
   Radius = MagCx(ReNum2 / Den, ImNum2 / Den)
   InputStabRad = Radius
End Function
'
' ======================================================================================
'
'  Calculation for simultaneous conjugate match
'     Based on Maas, Nonlinear Microwave Circuits, (c)1997, pp 326-330
'
Function ReGammaSMatch(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   Magds = MagCx(ReDs, ImDs)
   MagS11 = MagCx(ReS11, ImS11)
   MagS22 = MagCx(ReS22, ImS22)
   B1 = 1 + (MagS11 ^ 2) - (MagS22 ^ 2) - (Magds ^ 2)
'   B2 = 1 + (MagS22 ^ 2) - (MagS11 ^ 2) - (MagDS ^ 2)
   ReC1 = ReS11 - ReCxMult(ReDs, ImDs, ReS22, -ImS22)
   ImC1 = ImS11 - ImCxMult(ReDs, ImDs, ReS22, -ImS22)
   MagC1 = MagCx(ReC1, ImC1)
'   ReC2 = ReS22 - ReCxMult(ReDs, ImDs, ReS11, -ImS11)
'   ImC2 = ImS22 - ImCxMult(ReDs, ImDs, ReS11, -ImS11)
'   MagC2 = MagCx(ReC2, ImC2)
   '   Now calculate the source reflection coefficient
   A1 = B1 ^ 2 - (4 * (MagC1 ^ 2))
   ReNum1 = B1 + ReSqrRoot(A1, 0)
   ImNum1 = 0 + ImSqrRoot(A1, 0)
   MagNum1 = MagCx(ReNum1, ImNum1)
   ReNum2 = B1 - ReSqrRoot(A1, 0)
   ImNum2 = 0 - ImSqrRoot(A1, 0)
   MagNum2 = MagCx(ReNum2, ImNum2)
   If MagNum1 < MagNum2 Then
      ReGS = ReCxDiv(ReNum1, ImNum1, 2 * ReC1, 2 * ImC1)
      ImGS = ImCxDiv(ReNum1, ImNum1, 2 * ReC1, 2 * ImC1)
   Else
      ReGS = ReCxDiv(ReNum2, ImNum2, 2 * ReC1, 2 * ImC1)
      ImGS = ImCxDiv(ReNum2, ImNum2, 2 * ReC1, 2 * ImC1)
   End If
   ReGammaSMatch = ReGS
End Function
Function ImGammaSMatch(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   Magds = MagCx(ReDs, ImDs)
   MagS11 = MagCx(ReS11, ImS11)
   MagS22 = MagCx(ReS22, ImS22)
   B1 = 1 + MagS11 ^ 2 - MagS22 ^ 2 - Magds ^ 2
'   B2 = 1 + MagS22 ^ 2 - MagS11 ^ 2 - MagDS ^ 2
   ReC1 = ReS11 - ReCxMult(ReDs, ImDs, ReS22, -ImS22)
   ImC1 = ImS11 - ImCxMult(ReDs, ImDs, ReS22, -ImS22)
   MagC1 = MagCx(ReC1, ImC1)
'   ReC2 = ReS22 - ReCxMult(ReDs, ImDs, ReS11, -ImS11)
'   ImC2 = ImS22 - ImCxMult(ReDs, ImDs, ReS11, -ImS11)
'   MagC2 = MagCx(ReC2, ImC2)
   '   Now calculate the source reflection coefficient
   A1 = B1 ^ 2 - (4 * (MagC1 ^ 2))
   ReNum1 = B1 + ReSqrRoot(A1, 0)
   ImNum1 = 0 + ImSqrRoot(A1, 0)
   MagNum1 = MagCx(ReNum1, ImNum1)
   ReNum2 = B1 - ReSqrRoot(A1, 0)
   ImNum2 = 0 - ImSqrRoot(A1, 0)
   MagNum2 = MagCx(ReNum2, ImNum2)
   If MagNum1 < MagNum2 Then
      ReGS = ReCxDiv(ReNum1, ImNum1, 2 * ReC1, 2 * ImC1)
      ImGS = ImCxDiv(ReNum1, ImNum1, 2 * ReC1, 2 * ImC1)
   Else
      ReGS = ReCxDiv(ReNum2, ImNum2, 2 * ReC1, 2 * ImC1)
      ImGS = ImCxDiv(ReNum2, ImNum2, 2 * ReC1, 2 * ImC1)
   End If
   ImGammaSMatch = ImGS
End Function


Function ReGammaLMatch(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   Magds = MagCx(ReDs, ImDs)
   MagS11 = MagCx(ReS11, ImS11)
   MagS22 = MagCx(ReS22, ImS22)
'   B1 = 1 + (MagS11 ^ 2) - (MagS22 ^ 2) - (MagDS ^ 2)
   B2 = 1 + (MagS22 ^ 2) - (MagS11 ^ 2) - (Magds ^ 2)
'   ReC1 = ReS11 - ReCxMult(ReDS, ImDS, ReS22, -ImS22)
'   ImC1 = ImS11 - ImCxMult(ReDS, ImDS, ReS22, -ImS22)
'   MagC1 = MagCx(ReC1, ImC1)
   ReC2 = ReS22 - ReCxMult(ReDs, ImDs, ReS11, -ImS11)
   ImC2 = ImS22 - ImCxMult(ReDs, ImDs, ReS11, -ImS11)
   MagC2 = MagCx(ReC2, ImC2)
   '   Now calculate the source reflection coefficient
   A2 = B2 ^ 2 - (4 * (MagC2 ^ 2))
   ReNum1 = B2 + ReSqrRoot(A2, 0)
   ImNum1 = 0 + ImSqrRoot(A2, 0)
   MagNum1 = MagCx(ReNum1, ImNum1)
   ReNum2 = B2 - ReSqrRoot(A2, 0)
   ImNum2 = 0 - ImSqrRoot(A2, 0)
   MagNum2 = MagCx(ReNum2, ImNum2)
   If MagNum1 < MagNum2 Then
      ReGL = ReCxDiv(ReNum1, ImNum1, 2 * ReC2, 2 * ImC2)
      ImGL = ImCxDiv(ReNum1, ImNum1, 2 * ReC2, 2 * ImC2)
   Else
      ReGL = ReCxDiv(ReNum2, ImNum2, 2 * ReC2, 2 * ImC2)
      ImGL = ImCxDiv(ReNum2, ImNum2, 2 * ReC2, 2 * ImC2)
   End If
   ReGammaLMatch = ReGL
End Function
Function ImGammaLMatch(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   Magds = MagCx(ReDs, ImDs)
   MagS11 = MagCx(ReS11, ImS11)
   MagS22 = MagCx(ReS22, ImS22)
'   B1 = 1 + MagS11 ^ 2 - MagS22 ^ 2 - MagDS ^ 2
   B2 = 1 + MagS22 ^ 2 - MagS11 ^ 2 - Magds ^ 2
'   ReC1 = ReS11 - ReCxMult(ReDs, ImDs, ReS22, -ImS22)
'   ImC1 = ImS11 - ImCxMult(ReDs, ImDs, ReS22, -ImS22)
'   MagC1 = MagCx(ReC1, ImC1)
   ReC2 = ReS22 - ReCxMult(ReDs, ImDs, ReS11, -ImS11)
   ImC2 = ImS22 - ImCxMult(ReDs, ImDs, ReS11, -ImS11)
   MagC2 = MagCx(ReC2, ImC2)
   '   Now calculate the source reflection coefficient
   A2 = B2 ^ 2 - (4 * (MagC2 ^ 2))
   ReNum1 = B2 + ReSqrRoot(A2, 0)
   ImNum1 = 0 + ImSqrRoot(A2, 0)
   MagNum1 = MagCx(ReNum1, ImNum1)
   ReNum2 = B2 - ReSqrRoot(A2, 0)
   ImNum2 = 0 - ImSqrRoot(A2, 0)
   MagNum2 = MagCx(ReNum2, ImNum2)
   If MagNum1 < MagNum2 Then
      ReGL = ReCxDiv(ReNum1, ImNum1, 2 * ReC2, 2 * ImC2)
      ImGL = ImCxDiv(ReNum1, ImNum1, 2 * ReC2, 2 * ImC2)
   Else
      ReGL = ReCxDiv(ReNum2, ImNum2, 2 * ReC2, 2 * ImC2)
      ImGL = ImCxDiv(ReNum2, ImNum2, 2 * ReC2, 2 * ImC2)
   End If
   ImGammaLMatch = ImGL
End Function




Function ReSqrRoot(ReX, ImX)
   MagX = MagCx(ReX, ImX)
   AngleX = Application.Atan2(ReX, ImX)
   ReY = (MagX ^ 0.5) * Cos(AngleX * 0.5)
   ImY = (MagX ^ 0.5) * Sin(AngleX * 0.5)
   ReSqrRoot = ReY
End Function
Function ImSqrRoot(ReX, ImX)
   MagX = MagCx(ReX, ImX)
   AngleX = Application.Atan2(ReX, ImX)
   ReY = (MagX ^ 0.5) * Cos(AngleX * 0.5)
   ImY = (MagX ^ 0.5) * Sin(AngleX * 0.5)
   ImSqrRoot = ImY
End Function


'
' ======================================================================================
'
'  Calculation for the Power Gain Circles
'     Based on Maas, Nonlinear Microwave Circuits, (c)1997, pp 329
'
Function RePwrGainCtr(GpdB As Double, ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' determinant of the s-matrix
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   Magds = MagCx(ReDs, ImDs)
   ' magnitude of S22
   MagS22 = MagCx(ReS22, ImS22)
   ' magnitude of S12*S21
   MagS12S21 = MagCx(ReCxMult(ReS12, ImS12, ReS21, ImS21), ImCxMult(ReS12, ImS12, ReS21, ImS21))
   ' Rollett stability factor
   K = Rollett(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ' normalized power gain as a ratio
   gp = 10 ^ (GpdB / 10) / (MagCx(ReS21, ImS21) ^ 2)
   ' numerator of formula  (see eqn 8.1.22 in Maas)
   ReNum = gp * (ReS22 - ReCxMult(ReDs, ImDs, ReS11, -ImS11))
   ImNum = -gp * (ImS22 - ImCxMult(ReDs, ImDs, ReS11, -ImS11))
   ' denominator of formula
   Den = 1 + (gp * (MagS22 ^ 2 - Magds ^ 2))
   ' Final result for the center of the power gain circle
   ReCp = ReNum / Den
   ImCp = ImNum / Den
   ' Calculate numerator for determining radius (see eqn 8.1.23)
'   Num2 = (1 - (2 * K * gp * MagS12S21) + ((gp * MagS12S21) ^ 2)) ^ 0.5
'   radius = Num2 / Den
   
   RePwrGainCtr = ReCp
'   ImPwrGainCtr = ImCp
'   PwrGainRad = radius
End Function
Function ImPwrGainCtr(GpdB As Double, ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' determinant of the s-matrix
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   Magds = MagCx(ReDs, ImDs)
   ' magnitude of S22
   MagS22 = MagCx(ReS22, ImS22)
   ' magnitude of S12*S21
   MagS12S21 = MagCx(ReCxMult(ReS12, ImS12, ReS21, ImS21), ImCxMult(ReS12, ImS12, ReS21, ImS21))
   ' Rollett stability factor
   K = Rollett(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ' normalized power gain as a ratio
   gp = 10 ^ (GpdB / 10) / (MagCx(ReS21, ImS21) ^ 2)
   ' numerator of formula  (see eqn 8.1.22 in Maas)
   ReNum = gp * (ReS22 - ReCxMult(ReDs, ImDs, ReS11, -ImS11))
   ImNum = -gp * (ImS22 - ImCxMult(ReDs, ImDs, ReS11, -ImS11))
   ' denominator of formula
   Den = 1 + (gp * (MagS22 ^ 2 - Magds ^ 2))
   ' Final result for the center of the power gain circle
   ReCp = ReNum / Den
   ImCp = ImNum / Den
   ' Calculate numerator for determining radius (see eqn 8.1.23)
'   Num2 = (1 - (2 * K * gp * MagS12S21) + ((gp * MagS12S21) ^ 2)) ^ 0.5
'   radius = Num2 / Den
   
'   RePwrGainCtr = ReCp
   ImPwrGainCtr = ImCp
'   PwrGainRad = radius
End Function
Function PwrGainRad(GpdB As Double, ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' determinant of the s-matrix
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   Magds = MagCx(ReDs, ImDs)
   ' magnitude of S22
   MagS22 = MagCx(ReS22, ImS22)
   ' magnitude of S12*S21
   MagS12S21 = MagCx(ReCxMult(ReS12, ImS12, ReS21, ImS21), ImCxMult(ReS12, ImS12, ReS21, ImS21))
   ' Rollett stability factor
   K = Rollett(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ' normalized power gain as a ratio
   gp = 10 ^ (GpdB / 10) / (MagCx(ReS21, ImS21) ^ 2)
   ' numerator of formula  (see eqn 8.1.22 in Maas)
   ReNum = gp * (ReS22 - ReCxMult(ReDs, ImDs, ReS11, -ImS11))
   ImNum = -gp * (ImS22 - ImCxMult(ReDs, ImDs, ReS11, -ImS11))
   ' denominator of formula
   Den = 1 + (gp * (MagS22 ^ 2 - Magds ^ 2))
   ' Final result for the center of the power gain circle
'   ReCp = ReNum / Den
'   ImCp = ImNum / Den
   ' Calculate numerator for determining radius (see eqn 8.1.23)
   Num2 = (1 - (2 * K * gp * MagS12S21) + ((gp * MagS12S21) ^ 2)) ^ 0.5
   Radius = Num2 / Den
   
'   RePwrGainCtr = ReCp
'   ImPwrGainCtr = ImCp
   PwrGainRad = Radius
End Function

' ======================================================================================
'
'  Calculation for the Available Gain Circles
'     Based on Maas, Nonlinear Microwave Circuits, (c)1997, pp 329
'
Function ReAvGainCtr(GadB As Double, ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' determinant of the s-matrix
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   Magds = MagCx(ReDs, ImDs)
   ' magnitude of S22
   MagS11 = MagCx(ReS11, ImS11)
   ' magnitude of S12*S21
   MagS12S21 = MagCx(ReCxMult(ReS12, ImS12, ReS21, ImS21), ImCxMult(ReS12, ImS12, ReS21, ImS21))
   ' Rollett stability factor
   K = Rollett(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ' normalized power gain as a ratio
   ga = 10 ^ (GadB / 10) / (MagCx(ReS21, ImS21) ^ 2)
   ' numerator of formula  (see eqn 8.1.22 in Maas)
   ReNum = ga * (ReS11 - ReCxMult(ReDs, ImDs, ReS22, -ImS22))
   ImNum = -ga * (ImS11 - ImCxMult(ReDs, ImDs, ReS22, -ImS22))
   ' denominator of formula
   Den = 1 + (ga * (MagS11 ^ 2 - Magds ^ 2))
   ' Final result for the center of the power gain circle
   ReCa = ReNum / Den
   ImCa = ImNum / Den
   ' Calculate numerator for determining radius (see eqn 8.1.23)
'   Num2 = (1 - (2 * K * ga * MagS12S21) + ((ga * MagS12S21) ^ 2)) ^ 0.5
'   radius = Num2 / Den
   
   ReAvGainCtr = ReCa
'   ImAvGainCtr = ImCa
'   AvGainRad = radius
End Function
Function ImAvGainCtr(GadB As Double, ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' determinant of the s-matrix
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   Magds = MagCx(ReDs, ImDs)
   ' magnitude of S22
   MagS11 = MagCx(ReS11, ImS11)
   ' magnitude of S12*S21
   MagS12S21 = MagCx(ReCxMult(ReS12, ImS12, ReS21, ImS21), ImCxMult(ReS12, ImS12, ReS21, ImS21))
   ' Rollett stability factor
   K = Rollett(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ' normalized power gain as a ratio
   ga = 10 ^ (GadB / 10) / (MagCx(ReS21, ImS21) ^ 2)
   ' numerator of formula  (see eqn 8.1.22 in Maas)
   ReNum = ga * (ReS11 - ReCxMult(ReDs, ImDs, ReS22, -ImS22))
   ImNum = -ga * (ImS11 - ImCxMult(ReDs, ImDs, ReS22, -ImS22))
   ' denominator of formula
   Den = 1 + (ga * (MagS11 ^ 2 - Magds ^ 2))
   ' Final result for the center of the power gain circle
   ReCa = ReNum / Den
   ImCa = ImNum / Den
   ' Calculate numerator for determining radius (see eqn 8.1.23)
'   Num2 = (1 - (2 * K * ga * MagS12S21) + ((ga * MagS12S21) ^ 2)) ^ 0.5
'   radius = Num2 / Den
   
'   ReAvGainCtr = ReCa
   ImAvGainCtr = ImCa
'   AvGainRad = radius
End Function
Function AvGainRad(GadB As Double, ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   ' determinant of the s-matrix
   ReDs = ReDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ImDs = ImDeltaS(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   Magds = MagCx(ReDs, ImDs)
   ' magnitude of S22
   MagS11 = MagCx(ReS11, ImS11)
   ' magnitude of S12*S21
   MagS12S21 = MagCx(ReCxMult(ReS12, ImS12, ReS21, ImS21), ImCxMult(ReS12, ImS12, ReS21, ImS21))
   ' Rollett stability factor
   K = Rollett(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   ' normalized power gain as a ratio
   ga = 10 ^ (GadB / 10) / (MagCx(ReS21, ImS21) ^ 2)
   ' numerator of formula  (see eqn 8.1.22 in Maas)
   ReNum = ga * (ReS11 - ReCxMult(ReDs, ImDs, ReS22, -ImS22))
   ImNum = -ga * (ImS11 - ImCxMult(ReDs, ImDs, ReS22, -ImS22))
   ' denominator of formula
   Den = 1 + (ga * (MagS11 ^ 2 - Magds ^ 2))
   ' Final result for the center of the power gain circle
   ReCa = ReNum / Den
   ImCa = ImNum / Den
   ' Calculate numerator for determining radius (see eqn 8.1.23)
   Num2 = (1 - (2 * K * ga * MagS12S21) + ((ga * MagS12S21) ^ 2)) ^ 0.5
   Radius = Num2 / Den
   
'   ReAvGainCtr = ReCa
'   ImAvGainCtr = ImCa
   AvGainRad = Radius
End Function


Function MaxStableGain(ReS11 As Double, ImS11 As Double, ReS21 As Double, ImS21 As Double, ReS12 As Double, ImS12 As Double, ReS22 As Double, ImS22 As Double)
   K = Rollett(ReS11, ImS11, ReS21, ImS21, ReS12, ImS12, ReS22, ImS22)
   If (K < 1) Then
      realpart = ReCxDiv(ReS21, ImS21, ReS12, ImS12)
      imagpart = ImCxDiv(ReS21, ImS21, ReS12, ImS12)
      Gain = MagCx(realpart, imagpart)
   Else
      Gain = -99
   End If
 
   MaxStableGain = Gain
End Function
