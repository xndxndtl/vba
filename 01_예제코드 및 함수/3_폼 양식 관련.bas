1) 워크시트 내 양식 및 데이터를 초기에 세팅해놓고자 하면
   1. 모듈이 아닌 워크시트 개체에다가 아래와 같이 코드를 쓰면 된다.
      아래 예제는 b2셀에 a칼럼의 데이터가 들어있는 마지막 값을 입력해 주는 것이다.

   Private Sub Worksheet_SelectionChange(ByVal Target As Range)

     Range("b2").Value = Range("a10000").End(xlUp) + 1

   End Sub

2) actveX 콤보박스 내 목록 리스트 항목 동적 범위 설정
   위 1번과 연계하여, 시트 내 삽입된 콤보박스의 데이터를 곧바로 설정

   Private Sub Worksheet_SelectionChange(ByVal Target As Range)

       ComboBox1.ListFillRange = "'급여대장'!a11:" & _
           Worksheets("급여대장").Range("A100000").End(xlUp).Address

   End Sub

2) 스핀단추 연계
