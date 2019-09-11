VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Color_Legend_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Code for Form That uses MyColor_Dialog module for picking colors 

Private Sub Detail_Paint()
    Paint_Text_Box
End Sub
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
    Paint_Text_Box
End Sub
Private Sub Paint_Text_Box()
    If IsNull(Txt_Color_Number) Then
        If IsNull(Txt_Color_Hex) Then
            Txt_Legend_Color.BackColor = 0
        Else
            Txt_Legend_Color.BackColor = ColorConstants.vbWhite
        End If
    Else
        Txt_Legend_Color.BackColor = Txt_Color_Number
    End If
End Sub
Private Sub Txt_Color_Hex_DblClick(Cancel As Integer)
    Get_New_Color
End Sub
Private Sub Txt_Color_Number_DblClick(Cancel As Integer)
    Get_New_Color
End Sub
Private Sub Txt_Color_RGB_DblClick(Cancel As Integer)
    Get_New_Color
End Sub
Private Sub Get_New_Color()
    Txt_Color_Hex = MyColor_Dialog.ChooseVBAColor(Txt_Legend_Color.BackColor)
    'Convert to long Number
    Txt_Color_Number = Color_Hex_To_Long(Txt_Color_Hex)
    'Convert to RGB
    Txt_Color_RGB = Color_Long_To_RGB(Txt_Color_Number)
    Txt_Color_Description.SetFocus
    Txt_Date_Updated = Now()
End Sub

