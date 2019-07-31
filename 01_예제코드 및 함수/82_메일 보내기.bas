--------------------------------------------------------------------------------
2) 메일 보내기

Private Function Mailto(strAddress As String, strSubjectTitle As String, strBodyString As String, strAttachFile As String) As Boolean
    Dim oOutlook    As Outlook.Application
    Dim oMailItem   As Outlook.MailItem
    Dim iAppVer     As Integer
    Dim varAttachFile As Variant

    On Error GoTo ErrProcedure
    iAppVer = Val(Outlook.Application.Version)

    Set oOutlook = CreateObject("Outlook.Application" & IIf(iAppVer < 12, "", "." & iAppVer)) ' 버전별 설정
    Set oMailItem = oOutlook.CreateItem(olMailItem) ' 메일설정
    With oMailItem
         .To = strAddress ' 받는사람 주소
         .Subject = strSubjectTitle ' 제목
         .Body = strSubjectTitle & vbNewLine & strBodyString & vbNewLine ' 내용
         .Importance = olImportanceHigh
         If InStr(strAttachFile, strSpliter) Then ' 여러개면
            For Each varAttachFile In Split(strAttachFile, strSpliter)
                .Attachments.Add varAttachFile, olByValue, , varAttachFile ' 첨부파일
            Next
        Else ' 하나면
            .Attachments.Add strAttachFile, olByValue, , strAttachFile ' 첨부파일
        End If
        .Send
    End With

ErrProcedure: ' 에러시 이곳으로 점프
    Set oMailItem = Nothing
    Set oOutlook = Nothing
    If Err = 0 Then
        Application.StatusBar = "메일이 성공적으로 발송되었습니다."
    ElseIf Err = 429 Then
        MsgBox "Microsoft Outlook 개체를 작성할 수 없습니다.", vbExclamation, "메일 발송 실패"
    Else
        MsgBox Err.Description, vbExclamation, "메일 발송 실패"
    End If
End Function


--------------------------------------------------------------------------------
