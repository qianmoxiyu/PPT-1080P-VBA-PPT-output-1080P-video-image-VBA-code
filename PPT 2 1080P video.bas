Attribute VB_Name = "模块1"

Sub CustomizedVideoExport()

    Set myDocument = ActivePresentation
    Dim myPath As String
    
        ' 让用户选择输出路径
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "选择图片保存文件夹"
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
        
        If .Show = -1 Then
            myPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With
    
    boxResponse = MsgBox("要将当前演示文稿导出至 " & myPath & " 吗？", vbOKCancel, "导出确认")
    
    If boxResponse = vbOK And _
    myDocument.CreateVideoStatus <> ppMediaTaskStatusInProgress Then
    
        myDocument. _
            CreateVideo FileName:=myPath, _
            UseTimingsAndNarrations:=True, _
            DefaultSlideDuration:=1, _
            VertResolution:=1080, _
            FramesPerSecond:=30, _
            Quality:=100
        
        '显示导出完成弹窗
        MsgBox "视频导出已完成！" & vbCrLf & vbCrLf & _
               "文件位置: " & myPath & vbCrLf & _
               "分辨率: 1080p" & vbCrLf & _
               "帧率: 30fps", _
               vbOKOnly + vbInformation, "导出完成"
        
        '可选：打开导出目标文件夹
        'Shell Environ("windir") & "\explorer.exe """ & Left(myPath, InStrRev(myPath, "\")) & """", vbNormalFocus
            
    ElseIf boxResponse = vbCancel Then
        MsgBox "导出已取消", vbOKOnly, "导出取消"
    Else
        MsgBox "正在导出另一个视频，无法同时导出多个视频", vbOKOnly, "无法导出"
        
    End If

End Sub



