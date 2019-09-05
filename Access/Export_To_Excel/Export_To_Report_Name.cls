VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Export_To_Report_Name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Date: 2019/09/07
' Author: Gilbert Medel
' Current Version: 3.1.0
' Called by: Export To Excel, Sets up data to export, Formats report and fills in all data
'
' Public Variables
Public Local_ID As Variant
Public Network_PKEY As Variant
Public NOTES As Variant
Public Self As Object
Public Parent As Object
'
' Private Variables
Private Data_Recset As Object
Private Const Object_Type As String = "Export_To_Excel_Report_Name" 'Identifies Class object Type
Private Const Query_Name As String = "Q_Export_To_Excel_Report_Name"
Private Const Sheet_Name As String = "Report_Name Details"
'Unique Variables
'
' Public properties
Public Property Get Name() As String
    Name = Object_Type
End Property
'
' Public Functions that return values
Public Function Create_Report(Excel_Export_Parent As Object) As Boolean
     Parent = Excel_Export_Parent
     Parent.Excel_WorkSheet.Name = Sheet_Name
    'Set Up Records
    Prepare_For_Export
    Setup_RecordSet
    Export_Report_Title
    Export_Report_Footer
    'Return to Parent
End Function
'
' Public Sub Routines
'
' Private Functions
Private Function Setup_RecordSet() As Boolean
    'Data Records For Title color coding Legend
    Set Data_Recset = Parent.Parent.Application.CurrentDb.QueryDefs("Q_Export_To_Excel_Color_Legend").OpenRecordset
    
    'Data Records for Report
   Set Data_Recset = Parent.Parent.Application.CurrentDb.QueryDefs(Query_Name).OpenRecordset
End Function
Private Function Prepare_For_Export() As Boolean
    'Formats The Cell Layouts for
    
End Function
Private Function Export_Report_Title() As Boolean
    'Fills in Title Rows
    
End Function
Private Function Export_Report_Footer() As Boolean
    'Fills in Distrobution statement
    
End Function
Private Function Export_Report_Sheet1() As Boolean
    'Fills out Install Details
    
End Function
Private Function Export_Report_Sheet2() As Boolean
    'Fills out Install Details
    
End Function
Private Function Export_Report_Sheet3() As Boolean
    'Fills out Install Details
    
End Function
'
' Private Subs Routines
Public Sub Zeroize()
    Local_ID = Null
    Network_PKEY = Null
    NOTES = Null
    Set Parent = Nothing
End Sub
Private Sub Class_Initialize()
    Zeroize
End Sub
Private Sub Class_Terminate()
    Zeroize
End Sub
'
'End Code
