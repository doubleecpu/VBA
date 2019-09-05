VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Record_Table_Structure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
' Date: 09/03/2019
' Author: Gilbert Medel
' Current Version: 3.1.0
' Notes: This will record the table structure of the Access file selected
' Needs a reference to the Access.Application object Can be expanded to capture more properties/attributes by adding fields to table list or design tables
'
' Public Variables
'  Objects
Public Parent As Object
Public Self As Object
Public Result As Integer
'
' Private Variables
'  Objects
Private Access_Application As Object 'Application
Private Current_File As Object 'Database
Private Design_File As Object 'Database
Private Table_Design_Records As Object 'Recordset2
Private Table_List_Records As Object 'Recordset2
Private Table_from_list As Object 'TableDef
Private Table_Field As Object 'Field2
'  Variables
Private Table_List_Key As Variant
Private Table_Design_Key As Variant
'
'  Initialize
Private Sub Class_Initialize()

End Sub

'
' Public Properties
'
' Public Functions
'
' Public Sub
Public Sub Record_Table_Data(Parent_Reference As Object)
    If Parent_Reference Is Nothing Then
        MsgBox "No Parent Application Reference", vbCritical, "Record_Table_Structure"
        Result = 1 'No parent Error
    Else
        Set Parent = Parent_Reference
        'If Variables are set Get Table Structure
        If Set_Variables And Open_Design_Database Then
            Generate_Table_List
            Table_Design_Records.Close
            Table_List_Records.Close
            Result = 0 'No Errors
        Else
             Result = 2 'Data Error
        End If
    End If
End Sub
'
' Private Functions
Private Function Tables_Exist() As Boolean
    Dim Tables_Found(1) As Boolean
    If Current_File Is Nothing Then
        Tables_Exist = False
    Else
        For Each Table_from_list In Current_File.TableDefs
            If Table_from_list.Name = "Table_List" Then
                Tables_Found(0) = True
            End If
            If Table_from_list.Name = "Table_Design" Then
                Tables_Found(1) = True
            End If
        Next Table_from_list
    End If
    If Tables_Found(0) Then
        'Exists
    Else
        'Create table
        Set Table_from_list = Nothing
        Set Table_from_list = Current_File.CreateTableDef("Table_List")
        'Create Local_ID as Auto incrementing primary key
        Set Table_Field = Table_from_list.CreateField("Local_ID", DB_LONG)
        Table_Field.Attributes = DB_AUTOINCRFIELD
        Table_from_list.Fields.Append Table_Field
        Set Table_List_Key = Table_from_list.CreateIndex("Table_ID")
        Table_List_Key.Primary = True
        Set Table_Field = Table_List_Key.CreateField("Local_ID")
        Table_List_Key.Fields.Append Table_Field
        Table_from_list.Indexes.Append Table_List_Key
        
        Set Table_Field = Table_from_list.CreateField("Table_Name", DB_TEXT, 255)
        Table_from_list.Fields.Append Table_Field
        Set Table_Field = Table_from_list.CreateField("Version", DB_TEXT, 255)
        Table_from_list.Fields.Append Table_Field
        Set Table_Field = Table_from_list.CreateField("Date_Updated", DB_Date)
        Table_from_list.Fields.Append Table_Field
        Set Table_Field = Table_from_list.CreateField("Date_Created", DB_Date)
        Table_from_list.Fields.Append Table_Field
        Current_File.TableDefs.Append Table_from_list
    End If
    Current_File.TableDefs.Refresh
    If Tables_Found(1) Then
        'Exists
    Else
        'Create table
        Set Table_from_list = Current_File.CreateTableDef("Table_Design")
        'Create Local_ID as Auto incrementing primary key
        Set Table_Field = Table_from_list.CreateField("Local_ID", DB_LONG)
        Table_Field.Attributes = DB_AUTOINCRFIELD
        Table_from_list.Fields.Append Table_Field
        Set Table_List_Key = Table_from_list.CreateIndex("Table_ID")
        Table_List_Key.Primary = True
        Set Table_Field = Table_List_Key.CreateField("Local_ID")
        Table_List_Key.Fields.Append Table_Field
        Table_from_list.Indexes.Append Table_List_Key
        
        Set Table_Field = Table_from_list.CreateField("Table_List_FKEY", DB_LONG)
        Table_from_list.Fields.Append Table_Field
        Set Table_Field = Table_from_list.CreateField("Field_Name", DB_TEXT, 255)
        Table_from_list.Fields.Append Table_Field
        Set Table_Field = Table_from_list.CreateField("Data_Type", DB_TEXT, 255)
        Table_from_list.Fields.Append Table_Field
         Set Table_Field = Table_from_list.CreateField("Date_Updated", DB_Date)
        Table_from_list.Fields.Append Table_Field
         Set Table_Field = Table_from_list.CreateField("Date_Created", DB_Date)
        Table_from_list.Fields.Append Table_Field
        Current_File.TableDefs.Append Table_from_list
    End If
    Current_File.TableDefs.Refresh
    Tables_Exist = True
End Function
Private Function Set_Variables() As Boolean
    Set_Variables = False
    If Parent.Application.Name = "Microsoft Access" Then
        Set Access_Application = Parent.Application
        Set Current_File = Parent.Application.CurrentDb
       'Set Recordset to store Design
        If Tables_Exist Then
            Set Table_List_Records = Current_File.TableDefs("Table_List").OpenRecordset
            Set Table_Design_Records = Current_File.TableDefs("Table_Design").OpenRecordset
            Set_Variables = True
        Else
            Set_Variables = False
        End If
    End If
End Function
Private Function Open_Design_Database() As Boolean
    Dim Select_File As Object
    Set Select_File = Access_Application.FileDialog(3) '3 File Selection dialog
    With Select_File
        .AllowMultiSelect = False
        '.Application
        '.ButtonName = Folder
        '.Creator
        '.DialogType
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        .FilterIndex = 1
        '.InitialView
        '.Item
        '.Parent
        '.SelectedItems
        .Title = "Select File For Processing"
        .Show
        If .SelectedItems.Count = 1 Then
            If Right(.SelectedItems(1), 5) = "accdb" Then
                Set Design_File = Access_Application.DBEngine.OpenDatabase(.SelectedItems(1))
                Open_Design_Database = True
            Else
                Open_Design_Database = False
            End If
        Else
            Open_Design_Database = False
        End If
    End With
End Function
Private Function Is_Not_TID_Field(Field_Name As Variant) As Boolean
    If Field_Name = "Date_Created" Or Field_Name = "Data_Lock" Or Field_Name = "Date_Locked" Or Field_Name = "Date_Updated" Or Field_Name = "Data_Matched" Or Field_Name = "Data_Exported" Or Field_Name = "Local_ID" Or Field_Name = "Network_PKEY" Then
        Is_Not_TID_Field = False
    Else
        Is_Not_TID_Field = True
    End If
End Function
'
' Private Subs
Private Sub Zeroize()
    'Clear objects
    Set Parent = Nothing
    Set Self = Nothing
    Set Access_Application = Nothing
    Set Current_File = Nothing
    Set Design_File = Nothing
    Set Table_from_list = Nothing
    Set Table_Field = Nothing
    Set Table_Design_Records = Nothing
    Set Table_List_Records = Nothing
    ' Clear Variables
    Table_List_Key = Null
    Table_Design_Key = Null
End Sub
Private Sub Generate_Table_List()
    For Each Table_from_list In Design_File.TableDefs
        'check if table is not a access application table, temp table, paste error, error table
        If Left(Table_from_list.Name, 4) <> "MSys" And Left(Table_from_list.Name, 2) <> "~T" And Left(Table_from_list.Name, 4) <> "Past" And Left(Table_from_list.Name, 3) <> "Err" Then
            'Checks if table is already listed
            If IsNull(DLookup("Table_Name", "Table_List", "Table_Name = " & Chr(39) & Table_from_list.Name & Chr(39) & " AND Version = " & Chr(39) & Design_File.Properties("AppTitle") & Chr(39))) Then
                Table_List_Records.AddNew
                Table_List_Records.Fields("Table_Name") = Table_from_list.Name
                Table_List_Records.Fields("Version") = Design_File.Properties("AppTitle")
                Table_List_Records.Fields("Date_Updated") = Now
                Generate_Field_List
                Table_List_Records.Update
           Else
                'No Update Needed
            End If
        End If
    Next Table_from_list
End Sub
Private Sub Generate_Field_List()
    'Send Name
    For Each Table_Field In Table_from_list.Fields
        'Recor Field Name if not a TID Design Field
        If Is_Not_TID_Field(Table_Field.Name) Then
            Table_Design_Records.AddNew
            Table_Design_Records.Fields("Table_List_FKEY") = Table_List_Records.Fields("Local_ID")
            Table_Design_Records.Fields("Field_Name") = Table_Field.Name
            Table_Design_Records.Fields("Data_Type") = DLookup("DAO_Name", "DataTypeEnumeration_DAO", "Enumeration = " & Table_Field.Type)
            Table_Design_Records.Fields("Date_Updated") = Now
            Table_Design_Records.Update
        End If
    Next Table_Field
End Sub
'
' Terminate
Private Sub Class_Terminate()
    Zeroize
End Sub
'
' End Code
