Sub ExportAllDatabaseObjects()
    Dim obj As AccessObject
    Dim dbs As Object
    Dim strPath As String
    Dim fso As Object
    Dim objFile As Object
    
    ' SET YOUR EXPORT PATH HERE
    ' Examples:
    ' "C:\AccessExport\"
    ' "D:\Projects\AccessBackup\"
    ' Environ("USERPROFILE") & "\Desktop\AccessExport\"
    strPath = "C:\AccessExport\"
    
    ' Create File System Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create main folder and subfolders
    If Not fso.FolderExists(strPath) Then fso.CreateFolder strPath
    If Not fso.FolderExists(strPath & "Forms\") Then fso.CreateFolder strPath & "Forms\"
    If Not fso.FolderExists(strPath & "Modules\") Then fso.CreateFolder strPath & "Modules\"
    If Not fso.FolderExists(strPath & "Reports\") Then fso.CreateFolder strPath & "Reports\"
    If Not fso.FolderExists(strPath & "Queries\") Then fso.CreateFolder strPath & "Queries\"
    If Not fso.FolderExists(strPath & "Macros\") Then fso.CreateFolder strPath & "Macros\"
    
    Set dbs = Application.CurrentProject
    
    ' Export all FORMS
    Debug.Print "=== EXPORTING FORMS ==="
    For Each obj In dbs.AllForms
        Application.SaveAsText acForm, obj.Name, strPath & "Forms\" & obj.Name & ".txt"
        Debug.Print "Form exported: " & obj.Name
    Next obj
    
    ' Export all MODULES (VBA Code)
    Debug.Print "=== EXPORTING MODULES ==="
    For Each obj In dbs.AllModules
        Application.SaveAsText acModule, obj.Name, strPath & "Modules\" & obj.Name & ".txt"
        Debug.Print "Module exported: " & obj.Name
    Next obj
    
    ' Export all REPORTS
    Debug.Print "=== EXPORTING REPORTS ==="
    For Each obj In dbs.AllReports
        Application.SaveAsText acReport, obj.Name, strPath & "Reports\" & obj.Name & ".txt"
        Debug.Print "Report exported: " & obj.Name
    Next obj
    
    ' Export all QUERIES
    Debug.Print "=== EXPORTING QUERIES ==="
    Dim qdf As Object
    Dim qFile As Object
    On Error Resume Next
    For Each obj In CurrentData.AllQueries
        If Left(obj.Name, 1) <> "~" Then ' Skip system queries
            Set qdf = CurrentDb.QueryDefs(obj.Name)
            If Err.Number = 0 Then
                Set qFile = fso.CreateTextFile(strPath & "Queries\" & obj.Name & ".sql", True)
                qFile.WriteLine qdf.SQL
                qFile.Close
                Debug.Print "Query exported: " & obj.Name
            Else
                Debug.Print "Query skipped (error): " & obj.Name
                Err.Clear
            End If
        End If
    Next obj
    On Error GoTo 0
    
    ' Export all MACROS
    Debug.Print "=== EXPORTING MACROS ==="
    For Each obj In dbs.AllMacros
        Application.SaveAsText acMacro, obj.Name, strPath & "Macros\" & obj.Name & ".txt"
        Debug.Print "Macro exported: " & obj.Name
    Next obj
    
    ' Create summary file
    Set objFile = fso.CreateTextFile(strPath & "ExportSummary.txt", True)
    objFile.WriteLine "Access Database Export Summary"
    objFile.WriteLine "=============================="
    objFile.WriteLine "Database: " & CurrentDb.Name
    objFile.WriteLine "Export Date: " & Now()
    objFile.WriteLine ""
    objFile.WriteLine "Forms exported: " & dbs.AllForms.Count
    objFile.WriteLine "Modules exported: " & dbs.AllModules.Count
    objFile.WriteLine "Reports exported: " & dbs.AllReports.Count
    objFile.WriteLine "Queries exported: " & CurrentData.AllQueries.Count
    objFile.WriteLine "Macros exported: " & dbs.AllMacros.Count
    objFile.Close
    
    ' Cleanup
    Set qdf = Nothing
    Set fso = Nothing
    
    MsgBox "Export Complete!" & vbCrLf & vbCrLf & _
           "Forms: " & dbs.AllForms.Count & vbCrLf & _
           "Modules: " & dbs.AllModules.Count & vbCrLf & _
           "Location: " & strPath, vbInformation
    
    ' Open the export folder
    Shell "explorer.exe " & strPath, vbNormalFocus
End Sub
