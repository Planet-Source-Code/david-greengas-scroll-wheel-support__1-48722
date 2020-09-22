Attribute VB_Name = "modHiWord"
Option Explicit
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
' Combines two integers into a long integer

Public Function MAKELONG(wLow As Long, wHigh As Long) As Long
  MAKELONG = LOWORD(wLow) Or (&H10000 * LOWORD(wHigh))
End Function

' Combines two integers into a long integer

Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
  MAKELPARAM = MAKELONG(wLow, wHigh)
End Function

' Returns the low 16-bit integer from a 32-bit long integer

Public Function LOWORD(dwValue As Long) As Integer
  MoveMemory LOWORD, dwValue, 2
End Function

' Returns the low 16-bit integer from a 32-bit long integer

Public Function HIWORD(dwValue As Long) As Integer
  MoveMemory HIWORD, ByVal VarPtr(dwValue) + 2, 2
End Function

