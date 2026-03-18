Attribute VB_Name = "模块1"
Sub Export1080PSlides_Simple()
    
    Set ppt = ActivePresentation
    Dim exportPath As String
    
    ' 让用户选择输出路径
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "选择图片保存文件夹"
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
        
        If .Show = -1 Then
            exportPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With
    
    
    ' 导出图片
    For i = 1 To ppt.Slides.Count
        ppt.Slides(i).Export exportPath & Format(i, "000") & ".png", "PNG", 1920, 1080
        DoEvents
    Next i
    
    MsgBox "成功导出 " & ppt.Slides.Count & " 张1080P图片！", vbInformation
    Shell "explorer.exe """ & exportPath & """", vbNormalFocus
    
End Sub
