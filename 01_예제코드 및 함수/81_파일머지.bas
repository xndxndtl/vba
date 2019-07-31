Option Explicit

Sub Main_Merge()

    Dim i As Integer
    Dim cntR As Long, cntRR As Long, cntC As Integer
    Dim Fd As FileDialog

    If MsgBox("파일을 합치시겠습니까?", vbQuestion + vbOKCancel, "") = vbCancel Then Exit Sub   '매크로 진행 유무 선택 (취소 -> 프로시저 종료)

    Set Fd = Application.FileDialog(msoFileDialogFilePicker)    'Fd 변수에 FilePicker 개체 할당
    With Fd
        .AllowMultiSelect = True    '다중 선택 허용
        .Show   'FileDialog 열기

        If .SelectedItems.Count = 0 Then Exit Sub   '선택 한 파일이 없다면, 프로시저 종료
    End With

Application.ScreenUpdating = False  '화면 업데이트 중지

    On Error Resume Next    '만약, Result 시트가 없을때 Delete하면 에러가 나기때문에 On error resume next 처리
    Application.DisplayAlerts = False   'Delete시, 삭제 하겠냐는 메세지 무시
    Sheets("Result").Delete 'Result 시트 삭제
    Application.DisplayAlerts = True
    On Error GoTo 0

    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Result"     'Result 시트 생성

    For i = 1 To Fd.SelectedItems.Count '선택 한 파일만큼 순환
        Workbooks.Open Fd.SelectedItems(i)  '선택 한 파일 열기
        cntR = Cells(1048576, 1).End(xlUp).Row  '행의 개수 파악
        cntC = Cells(1, 1000).End(xlToLeft).Column  '열의 개수 파악

        If i = 1 Then   '첫 번째 파일은 제목행까지 Copy
            Range(Cells(1, 1), Cells(cntR, cntC)).Copy ThisWorkbook.Sheets("Result").Cells(1, 1)
        Else    '두 번째 파일부터는 2행부터 Copy
            Range(Cells(1, 1), Cells(cntR, cntC)).Copy ThisWorkbook.Sheets("Result").Cells(cntRR, 1)
        End If

        ActiveWorkbook.Close savechanges:=False '파일 닫기
        cntRR = Cells(1048576, 1).End(xlUp).Row + 1 'Result 시트의 행 개수 파악
    Next i

Application.ScreenUpdating = True   '화면 업데이트 활성화

     MsgBox "Complete", vbOKOnly, ""

End Sub
