Option Compare Database
Option Explicit
'
' Name:
' Date: 09/08/2019
' Author: Gilbert Medel
' Current Version: 3.1.0
' Called by: Report_[Report_Name]
'
' Notes: Used to create Workbooks and Worksheets, Will call Classes used to fill in worsheets and charts.
' Uses Factory Class to Create Objects
'
' Public Variables
Public Excel_App As Object
Public Excel_WorkBook As Object
Public Excel_WorkSheet As Object
Public MyVBA_App As Object
Public Parent As Object
Public Self As Object
Public NOTES As Variant
'
' Private Variables
Private Data_Recset As Object
Private Export_Schedule As Object
Private File_Test As Object
Private Workbook_Name As Variant
Private Worksheet_Name As Variant
Private Const Object_Type As String = "Export_Excel" 'Identifies Class object Type
'Unique Variables
'
' Public properties
Public Property Get Name() As String
    Name = Object_Type
End Property
'
' Public Functions that return values
Public Function Create_Report(Created_By As Object) As Boolean
    Dim Successful As Boolean
    Successful = False
    'Reads Object Name of whatever created class and created Class Used for Generating Report
    'Need to add methods for calling Report.Name to case to generate other excel files 
    If IsNull(Created_By.Name) Then
        Msg_Invalid_Name "Name_of_Export"
    Else
        ' Sets to Report.Application    
        Set MyVBA_App = Created_By.Application
        Export_Excel_File_Create
        Select Case Created_By.Name
            Case "Report_Name"
                 Export_Report_Name
                  Successful = True
            Case Else
               Msg_Invalid_Name "Name_of_Export"
               Successful = False
        End Select
    End If
    Create_Report = Successful
End Function
'
' Public Sub Routines
'
' Private Functions
'
' Private Subs Routines
Private Sub Export_Excel_File_Create()
'Creates File
  Set Excel_App = MyVBA_Module.Factory.Create_Office_Object("Excel.Application")
  Excel_App.Visible = True
  Excel_App.Workbooks.Add
  Set Excel_WorkBook = Excel_App.ActiveWorkbook
  Set Excel_WorkSheet = Excel_WorkBook.Worksheets.Add
End Sub
Private Sub Export_Report_Name()
    'Create Export Schedule Objects
    Set Export_Schedule = MyVBA_Module.Factory.Create_MyVBA_Object("Export_To_Excel_Report_Name")
    Set Export_Schedule.Parent = Self
    Export_Schedule.Report_Name Self
End Sub
Private Sub Msg_Invalid_Name(Invalid_Name As Variant)
    If IsNull(Invalid_Name) Then
        MsgBox "Missing Report Name", vbCritical, "Error Exporting"
    Else
         MsgBox "Report Name Not Valid: " & Invalid_Name, vbCritical, "Error Exporting"
    End If
End Sub
Private Sub Check_Export_Status(Recursion As Boolean)
    'Runs_Recursively until File Name does not match
    Set File_Test = MyVBA_Module.Factory.Create_SYNC_Object("SYNC_Connection_Class")
    If File_Test.Check_Export_Exists(Self, Workbook_Name) Then
        If Recursion Then
             '-4 removes extension from name
             Workbook_Name = Left(Workbook_Name, Len(Workbook_Name) - 4) + "1" + ".ppt"
             Check_Export_Status True
        Else
            Select Case MsgBox("File Exists do you want to delete and Replace?", vbYesNo, "Export To Excel")
                Case vbYes
                    'Delete Existing File
                    Set File_Test = Nothing
                    Set File_Test = MyVBA_Module.Factory.Create_Office_Object("Scripting.FileSystemObject")
                    File_Test.DeleteFile Parent.Application.CurrentProject.Path & "\" & Workbook_Name
                Case Else
                    Workbook_Name = Left(Workbook_Name, Len(Workbook_Name) - 4) + "1" + ".ppt"
                    Check_Export_Status True
            End Select
        End If
    End If
End Sub
Public Sub Zeroize()
    Set Excel_App = Nothing
    Set Excel_WorkBook = Nothing
    Set Excel_WorkSheet = Nothing
    Set Parent = Nothing
    Set Self = Nothing
    Set Export_Schedule = Nothing
    Workbook_Name = Empty
    Worksheet_Name = Empty
    NOTES = Empty
End Sub
Private Sub Class_Initialize()
    Zeroize
End Sub
Private Sub Class_Terminate()
    Zeroize
End Sub
'
'End Code
