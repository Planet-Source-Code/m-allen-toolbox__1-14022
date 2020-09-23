Attribute VB_Name = "Module1"
Public Cel As Long
Public Kel As Long
Public Far As Long
Public Ran As Long
Public Billion As Long

Public Dia As Double
Public Length As Double
Public CuInIn As Double
Public ToCuIn As Double
Public ToCuFt As Double
Public Const Pi As Double = 3.14159
Public Radias As Double
Public Const StdCuFt As Double = 1728
Public Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)


Public Function FindVolt()
volt = ohm * amp

End Function




Public Function ConvCel()
Kel = Cel + 273.15
Far = (Cel * 1.8) + 32
Ran = (Cel * 1.8) + 491.67
End Function

Public Function ConvKel()
Cel = Kel - 273.15
Ran = Kel * 1.8
Far = (Kel * 1.8) - 459.67
End Function

Public Function ConvRan()
Cel = (Ran - 491.67) * 1.8
Far = (Ran - 459.67)
Kel = (Ran * 1.8)
End Function

Public Function ConvFar()
Cel = (Far - 32) / 1.8
Kel = (Far + 459.67) / 1.8
Ran = (Far + 459.67)
End Function

Public Function ClrList()
frmToolBox.lstAmortize.Clear


End Function

