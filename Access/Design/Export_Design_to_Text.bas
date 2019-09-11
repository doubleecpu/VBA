Attribute VB_Name = "Export_Design_to_Text"
Option Compare Database
Option Explicit

Private Const VBA_MODULE  As Integer = 1
Private Const VBA_CLASS As Integer = 2
Private Const VBA_FORM  As Integer = 100
Private Const Extension_TABLE As String = ".tbl"
Private Const Extension_QUERY  As String = ".qry"
Private Const Extension_MODULE  As String = ".bas"
Private Const Extension_CLASS As String = ".cls"
Private Const Extension_FORM  As String = ".frm"
Private Const Code_From  As String = "Code_From_"
Private Const Ok_To_Save As Boolean = True               ' False: just generate the script
'
Public Sub saveAllAsText()
Dim Table_Definition As Object
Dim Query_Definition  As Object
Dim Container_Object  As Object
Dim Form_Object  As Object
Dim Module_Object As Object
Dim FileSystem_Object As Object

Dim Current_Path As String
Dim Reference_Name As String
Dim File_Name As String
'
On Error GoTo errHandler
    Current_Path = CurrentProject.Path
    Set FileSystem_Object = CreateObject("Scripting.FileSystemObject")
    Current_Path = addFolder(FileSystem_Object, Current_Path, Code_From & Application.CurrentProject.Name)
    Current_Path = addFolder(FileSystem_Object, Current_Path, Format(Date, "yyyy.mm.dd"))

'    For Each Table_Definition In CurrentDb.TableDefs
'        Reference_Name = Table_Definition.Name
'        If Left(Reference_Name, 4) <> "MSys" Then
'            File_Name = Current_Path & "\" & Reference_Name & Extension_TABLE
'            If Ok_To_Save Then
'                Application.ExportXML _
'                acExportTable, _
'                Reference_Name, _
'                File_Name, _
'                File_Name & ".XSD", _
'                File_Name & ".XSL", , _
'                acUTF8, _
'                acEmbedSchema + acExportAllTableAndFieldProperties
'            Else
'                Debug.Print "Application.ImportXML """ & File_Name & """, acStructureAndData"
'            End If
'        End If
'    Next
'
'    For Each Query_Definition In CurrentDb.QueryDefs
'        Reference_Name = Query_Definition.Name
'        If Left(Reference_Name, 1) <> "~" Then
'            File_Name = Current_Path & "\" & Reference_Name & Extension_QUERY
'            If Ok_To_Save Then
'                Application.SaveAsText _
'                acQuery, _
'                Reference_Name, _
'                File_Name
'            Else
'                Debug.Print "Application.LoadFromText acQuery, """ & Reference_Name & """, """ & File_Name & """"
'            End If
'        End If
'    Next

    Set Container_Object = Nothing
    Set Container_Object = Workspaces(0).Databases(0).Containers("Forms")
    For Each Form_Object In Container_Object.Documents
        Reference_Name = Form_Object.Name
        File_Name = Current_Path & "\" & Reference_Name & Extension_FORM
        If Ok_To_Save Then
            Application.SaveAsText _
            acForm, _
            Reference_Name, _
            File_Name
        Else
            Debug.Print "Application.LoadFromText acForm, """ & Reference_Name & """, """ & File_Name & """"
        End If
    Next

    Current_Path = addFolder(FileSystem_Object, Current_Path, "modules")
    For Each Module_Object In Application.VBE.ActiveVBProject.VBComponents
        Reference_Name = Module_Object.Name
        File_Name = Current_Path & "\" & Reference_Name
        Select Case Module_Object.Type
            Case VBA_MODULE
                If Ok_To_Save Then
                    Module_Object.Export File_Name & Extension_MODULE
                Else
                    Debug.Print "Application.VBE.ActiveVBProject.VBComponents.Import """ & File_Name & Extension_MODULE; """"
                End If
            Case VBA_CLASS
                If Ok_To_Save Then
                    Module_Object.Export File_Name & Extension_CLASS
                Else
                    Debug.Print "Application.VBE.ActiveVBProject.VBComponents.Import """ & File_Name & Extension_CLASS; """"
                End If
            Case VBA_FORM
                ' Do not export form modules (already exported the complete forms)
            Case Else
                Debug.Print "Unknown module type: " & Module_Object.Type, Module_Object.Name
        End Select
    Next
    If Ok_To_Save Then
        MsgBox "Files saved in  " & Current_Path, vbOKOnly, "Export Complete"
    End If
Exit Sub
errHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf
    Stop: Resume
End Sub
'
' Create a folder when necessary. Append the folder name to the given path.
Private Function addFolder(ByRef FileSystem_Object As Object, ByVal Current_Path As String, ByVal Add_To_Path As String) As String
    addFolder = Current_Path & "\" & Add_To_Path
    If Not FileSystem_Object.FolderExists(addFolder) Then MkDir addFolder
End Function
