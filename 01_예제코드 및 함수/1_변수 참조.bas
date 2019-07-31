5) 변수 그냥 설명

    프로시저 내
     - DIM, STATIC ==> 온니 프로시저 내 사용되는 변수 (함수 내 지역변수)
    모듈선두
     - DIM, PRIVATE ==> 모듈 내의 모든 프로시저에 적용되는 변수 (모듈 내 전역변수)
     - PUBLIC ==> 통합문서 전체 모듈에 적용되는 변수 (문서 내 전역변수)

    '모든 프로시저의 첫 행에 아래 행을 쓰면 '변수를 반드시 선언하며 프로그래밍하는 것이 규칙'이라는 것을 룰로 받아들여서, 변수를 안쓰면 에러로 받아들임.
    '좋은 프로그래밍을 위해 습관을 들이면 좋음.
Option Explicit

Public a As Integer
Dim b As Boolean
Public Const PI As Integer = 3.14159

Sub Macro1()
    Dim c As String
    a = (10 + 4) / 2
    b = (3 = 4)
    c = "한글엑셀" & 2010
    macro2 3
    c = c Like "*2007"


End Sub

Sub macro2(x As Integer)

    a = PI * x

End Sub
