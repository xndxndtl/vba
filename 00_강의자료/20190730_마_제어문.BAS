Option Explicit

Dim 행 As Byte

Sub 제어문()


    Dim result As Byte

    result = MsgBox("당신은 남자입니까?", vbYesNoCancel, "남녀구분")

    If result = 6 Then
        MsgBox "당신은 남"
    ElseIf result = vbNo Then
        MsgBox "당신은 여"
    Else
        MsgBox "당신은...?"
    End If


End Sub

Sub 제어문2()


    Dim 점수 As Byte


    행 = Application.WorksheetFunction.CountA(Range("b:b")) + 1
    점수 = InputBox("점수를 입력하세요", , 0)
    Cells(행, 2).Value = 점수

    If 점수 >= 70 Then
        Cells(행, 3).Value = "합격"
    Else
        Cells(행, 3).Value = "불합격"
    End If


    조건문 '아래 조건먼 함수 호출


End Sub

Sub 조건문()

    Select Case Cells(행, 2).Value
        Case Is > 90: Cells(행, 4).Value = "수"
        Case Is > 80: Cells(행, 4).Value = "우"
        Case Is > 70: Cells(행, 4).Value = "미"
        Case Is > 60: Cells(행, 4).Value = "양"
        Case Else: Cells(행, 4).Value = "가"
    End Select

End Sub

Sub 반복문()

 Dim 최종행 As Byte


    최종행 = Application.WorksheetFunction.CountA(Range("b:b")) + 1

    For 행 = 2 To 최종행 Step 1

        Cells(행, 1).Value = 행 - 1
    Next

End Sub

Sub 반복문_DoWhileLoop()

    Dim 더하기 As Byte

    행 = 2
    더하기 = 0
    Do While Cells(행, 1).Value <> ""
        더하기 = 더하기 + Cells(행, 1).Value
        행 = 행 + 1
    Loop
    MsgBox 더하기

End Sub


Sub 반복문_DoWhileLoop2()

    Dim 점수합 As Byte

    행 = 2
    점수합 = 0

    Do While Cells(행, 2).Value <> ""
        점수합 = 점수합 + Cells(행, 2).Value
        행 = 행 + 1
    Loop
    MsgBox "점수평균은" & 점수합 / (행 - 2)

End Sub

Sub 반복문_DoWhileLoop3()

    행 = 2

    Do While Cells(행, 2).Value <> "" '★b칼럼 빈셀을 찾을때까지 반복.

        If Cells(행, 2).Value >= 90 Then
            Cells(행, 4).Value = "수"
        Else
            Cells(행, 4).Value = "우미양가"
        End If

    행 = 행 + 1
    Loop

    'do while 루프는 최종 값이 결정되지 않을 때 쓰기 유용(위처럼)
     for 루프는 1~10까지 등 영역이 정해져있을 경우 유용

End Sub
