Attribute VB_Name = "Module1"
Function BSTV(A As Double, B As Double, C As Double, D As Double, E As Double, F As Integer)


d11 = D * (E ^ (1 / 2))

d12 = 1 / d11

d13 = Application.WorksheetFunction.Ln(A / B)

d14 = D ^ (2)

d15 = d14 / 2

d16 = d15 + C

d17 = d16 * E

d18 = d13 + d17

d1 = d12 * d18

d2 = d1 - (D * (E ^ (1 / 2)))

CALL1 = Application.WorksheetFunction.Norm_Dist(d1, 0, 1, True)

CALL2 = CALL1 * A

CALL3 = Application.WorksheetFunction.Norm_Dist(d2, 0, 1, True)

CALL4 = Exp((-C * E))

CALL5 = CALL3 * CALL4 * B

CALL0 = CALL2 - CALL5

PUT1 = Application.WorksheetFunction.Norm_Dist(-d2, 0, 1, True)

PUT2 = PUT1 * CALL4 * B

PUT3 = Application.WorksheetFunction.Norm_Dist(-d1, 0, 1, True)

PUT4 = PUT3 * A

PUT0 = PUT2 - PUT4


V = (F * CALL0) + ((1 - F) * PUT0)


BSTV = V
   
   

End Function

Function PnL(A As Double, B As Double, C As Double, D As Double, E As Double, F As Integer, G As Integer, H As Double)

d11 = D * (E ^ (1 / 2))

d12 = 1 / d11

d13 = Application.WorksheetFunction.Ln(A / B)

d14 = D ^ (2)

d15 = d14 / 2

d16 = d15 + C

d17 = d16 * E

d18 = d13 + d17

d1 = d12 * d18

d2 = d1 - (D * (E ^ (1 / 2)))

CALL1 = Application.WorksheetFunction.Norm_Dist(d1, 0, 1, True)

CALL2 = CALL1 * A

CALL3 = Application.WorksheetFunction.Norm_Dist(d2, 0, 1, True)

CALL4 = Exp((-C * E))

CALL5 = CALL3 * CALL4 * B

CALL0 = CALL2 - CALL5


PUT1 = Application.WorksheetFunction.Norm_Dist(-d2, 0, 1, True)

PUT2 = PUT1 * CALL4 * B

PUT3 = Application.WorksheetFunction.Norm_Dist(-d1, 0, 1, True)

PUT4 = PUT3 * A

PUT0 = PUT2 - PUT4


V = (F * CALL0) + ((1 - F) * PUT0)

D = V - H

P = D * 50 * G


PnL = P


   
   

End Function




Function Delta1(A As Double, B As Double, C As Double, D As Double, E As Double, F As Integer, G As Integer)



d11 = D * (E ^ (1 / 2))


d12 = 1 / d11


d13 = Application.WorksheetFunction.Ln(A / B)


d14 = D ^ (2)


d15 = d14 / 2


d16 = d15 + C


d17 = d16 * E


d18 = d13 + d17


d1 = d12 * d18

T1 = Application.WorksheetFunction.Norm_Dist(d1, 0, 1, True)

T2 = -Application.WorksheetFunction.Norm_Dist(-d1, 0, 1, True)

TF = (F * T1) + ((1 - F) * T2)

TFF = TF * G


Delta1 = TFF



End Function


Function Gamma(A As Double, B As Double, C As Double, D As Double, E As Double, F As Double)


d11 = D * (E ^ (1 / 2))


d12 = 1 / d11


d13 = Application.WorksheetFunction.Ln(A / B)


d14 = D ^ (2)


d15 = d14 / 2


d16 = d15 + C


d17 = d16 * E


d18 = d13 + d17


d1 = d12 * d18

K1 = Application.WorksheetFunction.Norm_Dist(d1, 0, 1, False)

K2 = A * D

K3 = K2 * (E ^ (1 / 2))

K = K1 / K3

KF = K * F


Gamma = KF



End Function


Function Vega(A As Double, B As Double, C As Double, D As Double, E As Double, F As Double)

d11 = D * (E ^ (1 / 2))


d12 = 1 / d11


d13 = Application.WorksheetFunction.Ln(A / B)


d14 = D ^ (2)


d15 = d14 / 2


d16 = d15 + C


d17 = d16 * E


d18 = d13 + d17


d1 = d12 * d18

V1 = Application.WorksheetFunction.Norm_Dist(d1, 0, 1, False)

V1 = (2 * Application.WorksheetFunction.Pi) ^ (1 / 2)

V2 = 1 / V1

V3 = Exp(-(d1 ^ (2)) / 2)

V4 = E ^ (1 / 2)

V = V2 * V3 * V4 * A

VF = V * F

Vega = VF



End Function


Function Theta(A As Double, B As Double, C As Double, D As Double, E As Double, F As Integer, G As Integer)

d11 = D * (E ^ (1 / 2))

d12 = 1 / d11

d13 = Application.WorksheetFunction.Ln(A / B)

d14 = D ^ (2)

d15 = d14 / 2

d16 = d15 + C

d17 = d16 * E

d18 = d13 + d17

d1 = d12 * d18

d2 = d1 - (D * (E ^ (1 / 2)))

T1 = Application.WorksheetFunction.Norm_Dist(d1, 0, 1, False)

T2 = Application.WorksheetFunction.Norm_Dist(d2, 0, 1, True)

T3 = Application.WorksheetFunction.Norm_Dist(-d2, 0, 1, True)

T4 = A * T1 * D

T5 = E ^ (1 / 2)

T6 = Exp((-C * E))

T7 = T4 / (2 * T5)

T8 = C * B * T6 * T2

TC = -T7 - T8

T9 = C * B * T6 * T3

TP = -T7 + T9

TF = (F * TC) + ((1 - F) * TP)

T = TF * G

Theta = T


End Function

Function PnLF(A As Double, B As Double, C As Double)

F = (A - B) * 50 * C


PnLF = F


End Function

