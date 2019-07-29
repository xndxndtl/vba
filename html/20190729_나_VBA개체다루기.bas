
Sub 개체_Application()   '1) 가장 상위 단계인 Application(엑셀 그 자체) 다루기

    MsgBox Application.Name & Application.Version

    Range("a1:c5").Value = 100

    MsgBox Application.WorksheetFunction.Sum(Range("a1:c5"))


    'Application.WorksheetFunction.
    '엑셀 워크시트에서 사용하는 함수를 사용하기 위해서는 "Application.WorksheetFunction." 을 입력 후 사용할 수 있다.

    MsgBox "데이터를 추가할 행의 번호는 " & _
    Application.WorksheetFunction.CountA(Range("a:a")) + 1
    ' _는 한줄로 쭉 쓴다는 뜻이고 :는 한줄로 썼더라도 enter를 쳤다는 것으로 인식된다.

  '  Application.Quit

End Sub

Sub 개체_Workbooks() '2) workbooks 연습

Application.ScreenUpdating = False

 '   Workbooks.Add

    Workbooks.Open ("C:\Users\82103\Desktop\김선묘\2. 파이썬\70. VBA\배포용_기초데이터\test.xlsx")
    MsgBox "활성화된 파일명은 " & ActiveWorkbook.Name & _
            "   작업중인 파일명은 " & ThisWorkbook.Name
    Range("a1").Value = 500
    Workbooks("VBA기본.xlsm").Worksheets(2).Range("a1").Value = 700

    Workbooks.Close

End Sub

Sub 개체_Worksheets() '3) worksheet 연습

    Worksheets(2).Name = "ddd"

    MsgBox Worksheets(2).Name

    Sheets.Add , Worksheets("ddd")
    ActiveSheet.Name = "fff"

    Worksheets(2).Range("a1").Value = 10
    'Worksheets(2).Range("a1:c5").Select  '==> 에러가 남. ★매우중요★ select는 활성화된 영역(시트)에서만 사용 가능함.

    '그래서 아래와 같이 activate 후 하면 올바르게 실행됨

    Worksheets(2).Activate
    Worksheets(2).Range("a1:c5").Select

    '강사 曰 select와 activate는 가능한 쓰지 않는게 좋다. 사용하면 프로그램 속도가 현저히 떨어짐. 왠만하면 다른 방법을 사용하라!

End Sub



Sub 개체_Range_Cells() '4) range 및 cell 연습

    Range("a1").Select
    Range("a1:c5").Select   ' ==> 범위로 선택 = Range("a1", "c5").Select
    Range("a1,c5").Select ' ==> a1, c5를 각각 선택
    Range("데이터").Select  '==> '셀의 특정영역을 선택 후 이름을 지정한 뒤 왼쪽과 같이 이름을 불러와서 범위 선택 가능.

    Cells(5, 3).Select '==> 셀은 한 셀만 선택 가능. 셀 순서설정은 행,열로 설정.

    '그럼 언제 range를 쓰고 cell을 쓰느냐?
    '반복문 등을 통해 셀영역을 변경시킬 때 cell 안의 행,열 번호를 변수로 지정하여 손쉽게 선택 가능.
    'range는 그럴때는 불-편.

    Selection.Offset(2, 1).Value = "행2줄 열1칸 이동"   '==>위에서 select를 했기 때문에 selection을 쓸 수 있음. f5로 실행해보면 암.

    'Range("a1").End(xlDown).Select
    Range("a1").End(xlDown).Offset(1, 0).Value = "데이터를 추가해야할 위치_위에서 아래로 탐색" '==> 마지막 열 + 1(offset활용) 을 통해 데이터 입력 열 탐색
    '★그런데! 보통 위에서부터 탐색하면 에러가 나거나 하는 경우 발생(예로 1행에만 데이터가 있을 경우 등).
    '★그.래.서 데이터를 맨 아래서 위로 탐색해서 마지막 행을 찾는 것이 더 좋음! (아래 참조)

    Range("a100000").End(xlUp).Offset(1, 0).Value = "데이터추가위치_아래서 위로 탐색)"

    Selection.EntireRow.Select
    Selection.EntireColumn.Select

    Rows(5).Select
    Columns(5).Select

    Range("a1").CurrentRegion.Select '==> 인접한 데이터 전체 선택


End Sub
