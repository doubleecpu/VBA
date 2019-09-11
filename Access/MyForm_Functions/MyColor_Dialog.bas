Attribute VB_Name = "MyColor_Dialog"
Option Compare Database
Option Explicit
'Establishes Color dialog reference in code

Declare Sub MyColorDialog _
  Lib "msaccess.exe" _
    Alias "#53" (ByVal Hwnd As Long, lngRGB As Long)
 
Public Function ChooseVBAColor(MyDefaultColor As Variant) As String
  Dim Color_Code As Long
  Color_Code = CLng("&H" & Right("000000" + _
                  Replace(Nz(MyDefaultColor, ""), "#", ""), 6))
  MyColorDialog Screen.ActiveForm.Hwnd, Color_Code
  ChooseVBAColor = "#" & Right("000000" & Hex(Color_Code), 6)
End Function
'Call this function within a form using following code:
'
'Me!txtYourColor = ChooseWebColor(Me!txtYourColor)

Public Function Color_Hex_To_Long(Hex_Code As Variant) As Long
    Color_Hex_To_Long = CLng(("&H" & Right(Hex_Code, Len(Hex_Code) - 1)))
End Function
Public Function Color_Long_To_RGB(ByVal lng_Color As Long) As String
    Dim int_Red As Integer
    Dim int_Blue As Integer
    Dim int_Green As Integer
    int_Red = lng_Color Mod &H100
    lng_Color = lng_Color \ &H100
    int_Green = lng_Color Mod &H100
    lng_Color = lng_Color \ &H100
    int_Blue = lng_Color Mod &H100
    Color_Long_To_RGB = CStr(int_Red & "," & int_Green & "," & int_Blue)
End Function
