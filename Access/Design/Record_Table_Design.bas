Attribute VB_Name = "Record_Table_Design"
Option Compare Database
Option Explicit
' Date: 08/08/2019
' Author: Gilbert Medel
' Current Version: 3.1.0
' Notes: Used To Create Record_Table_Structure class and Run Schema Recorder 
'
' Public Variables
'
' Private Variables
'
' Public properties
'
' Public Functions
'
' Public Sub
Public Sub Record_Schema(Parent_Reference As Object)
    Dim Factory_ As New DB_Factory
    Dim Message As Variant
    Dim Result As Integer
    Dim Schema_Recorder As Object
    ' Use Factory To Set Record_Table_Structure
    Set Schema_Recorder = Factory_.Create_DB_Object("Record_Table_Structure")
    Schema_Recorder.Record_Table_Data Parent_Reference
    Select Case Schema_Recorder.Result
        Case 0
            Message = "No Errors"
        Case 1
            Message = "Parent Reference Invalid"
        Case 2
            Message = "Table Data Errors"
        Case Else
            Message = "Unknown Error"
    End Select
    
    Set Factory_ = Nothing
    Set Schema_Recorder = Nothing
End Sub
'
' Private Functions
'
' Private Subs
'
'End Code
