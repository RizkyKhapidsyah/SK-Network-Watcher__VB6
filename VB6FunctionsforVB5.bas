Attribute VB_Name = "VB6FunctionsforVB5"
Option Explicit

'Found this function at:
'http://www.xbeat.net/vbspeed/c_Round.htm
'There are also a few other functions, to help those of us without VB6, on this site.

Public Static Function Round(dblNumber As Double, Optional ByVal numDecimalPlaces As Long) As Double
' by Donald, donald@xbeat.net, 20001018
  
  Dim fInit As Boolean
  Dim numDecimalPlacesPrev As Long
  Dim dFac As Double
  Dim dFacInv As Double
  Dim dTmp As Double
  
  ' calc factor once for this depth of rounding
  If Not fInit Or numDecimalPlacesPrev <> numDecimalPlaces Then
    dFac = 10 ^ numDecimalPlaces
    dFacInv = 10 ^ -numDecimalPlaces
    numDecimalPlacesPrev = numDecimalPlaces
    fInit = True
  End If
  
  If dblNumber >= 0 Then
    dTmp = dblNumber * dFac + 0.5
    Round = Int(dTmp) * dFacInv
  Else
    dTmp = -dblNumber * dFac + 0.5
    Round = -Int(dTmp) * dFacInv
  End If
  
End Function


