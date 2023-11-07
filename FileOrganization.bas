Attribute VB_Name = "模块3"
Option Explicit
Sub GetFloder_Shell()
    Dim objshell, objFolder
    Dim count As Long, l As Long
    count = 1
    l = 1
    
    Set objFolder = Application.FileDialog(msoFileDialogFolderPicker)
    If objFolder.Show = -1 Then
        ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(count, 1), Address:=objFolder.SelectedItems(1), _
        TextToDisplay:=objFolder.SelectedItems(1)
        count = count + 1
        Call GetFolderFile(objFolder.SelectedItems(1), count, l)
    Else
        Exit Sub
    End If
    
    Set objFolder = Nothing
    MsgBox "遍历完毕"
End Sub

Sub GetFolderFile(ByVal nPath As String, ByRef iCount As Long, level As Long)
    Dim iFileSys
    Dim iFile, gFile
    Dim iFolder, sFolder, nFolder
    
    
    Set iFileSys = CreateObject("Scripting.FileSystemObject")
    Set iFolder = iFileSys.GetFolder(nPath)
    Set sFolder = iFolder.SubFolders
    Set iFile = iFolder.Files
  
    With ActiveSheet
        If Not iFile Is Nothing Then
            For Each gFile In iFile
                .Hyperlinks.Add Anchor:=.Cells(iCount, 1), Address:=gFile.Path, TextToDisplay:= _
                "│" & Excel.Application.WorksheetFunction.Rept("    │", level - 1) & Space(4) & gFile.Name
                iCount = iCount + 1
            Next gFile
        End If
      
    '递归遍历所有子文件夹
        If Not sFolder Is Nothing Then
            For Each nFolder In sFolder
                .Hyperlinks.Add Anchor:=.Cells(iCount, 1), Address:=nFolder.Path, TextToDisplay:= _
                "│" & Excel.Application.WorksheetFunction.Rept("    │", level) & "─" & nFolder.Name
                iCount = iCount + 1
                Call GetFolderFile(nFolder.Path, iCount, level + 1)
            Next nFolder
        End If
    End With
    
    Set iFileSys = Nothing
    Set iFolder = Nothing
    Set sFolder = Nothing
    Set iFile = Nothing
End Sub

