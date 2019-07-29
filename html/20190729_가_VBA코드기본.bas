Public Sub 가_VBA코드기본()

Range("a1").Value = "VBA코드연습중222"

MsgBox "안녕하세요" & _
        Range("a1").Value & _
        "입니다.": MsgBox "테스트"

End Sub

Public Function 점유율구하기(부분, 전체)

    점유율구하기 = 부분 / 전체

End Function
