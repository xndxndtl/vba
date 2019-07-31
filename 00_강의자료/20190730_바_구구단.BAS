Option Explicit

Sub 구구단()

Worksheets.Add ', Worksheets("제어문")

'ActiveSheet.Name = "구구단출력"

Dim 단 As Byte

시작:  'GOTO 분기문
    단 = InputBox("몇단 출력?", "구구단질문", "1~9단까지 입력해주세요")

    If 단 = 0 Then
        MsgBox "입력된 값이 없음"
        Cells.Clear

    End If

Dim 행 As Byte
Dim 버튼 As Byte

    행 = 1
    Do While 행 <= 9
        Cells(행, 1).Value = 단
        Cells(행, 2).Value = "*"
        Cells(행, 3).Value = 행
        Cells(행, 4).Value = "="
        Cells(행, 5).Value = 단 * 행

        행 = 행 + 1

    Loop

    버튼 = MsgBox("구구단을 출력하였습니다. 계속할까요?", vbYesNo)

    If 버튼 = vbYes Then
        GoTo 시작
    End If


End Sub
