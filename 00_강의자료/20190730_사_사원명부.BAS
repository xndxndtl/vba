
Sub 사원명부_입력()

Dim 이름 As String
Dim 작업일수 As Integer
Dim 행 As Integer
Dim 버튼 As Byte


시작:

    이름 = InputBox("이름?", "이름묻기", "이름입력주세요(예:김또깡)")
    작업일수 = InputBox("작업일수?", "일수묻기")

    Range("a10000").End(xlUp).Offset(1, 0).Select

    With Selection
        .Value = Range("a10000").End(xlUp) + 1
        .Offset(0, 1).Value = 이름
        .Offset(0, 2).Value = 100000
        .Offset(0, 3).Value = 작업일수
        .Offset(0, 4).Value = 작업일수 * 100000
    End With

    버튼 = MsgBox("계속하시겠습니까?", vbYesNo)

    If 버튼 = vbYes Then
        GoTo 시작
    End If



End Sub

Sub 사원명부_찾기()

Dim 검색번호 As Integer
Dim 행 As Byte

    검색번호 = InputBox("번호입력")
    행 = 2

    Do While Cells(행, 1) <> ""
        If Cells(행, 1).Value = 검색번호 Then
            MsgBox Cells(행, 2) & "님의 수당은 " & Cells(행, 5)
        End If
        행 = 행 + 1
    Loop

End Sub

Sub 사원명부_삭제()

Dim 삭제번호 As Integer
Dim 행 As Byte

    삭제번호 = InputBox("삭제번호입력")
    행 = 2

    Do While Cells(행, 1) <> ""
        If Cells(행, 1).Value = 삭제번호 Then
            Rows(행).Delete
        End If
        행 = 행 + 1
    Loop

End Sub
