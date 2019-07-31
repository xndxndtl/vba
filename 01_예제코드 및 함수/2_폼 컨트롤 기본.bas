#VBA 유저 폼 기본 콘트롤 예시 코드

1) 텍스트상자

만약 폼에 데이터를 넣은 상태로 띄우고 싶다면,
아래와 같이 폼 바탕화면에서 미리 텍스트상자에 데이터를 입력해 둔다.

  Private Sub UserForm_activate()

      txt등록일자 = Date

  End Sub

2) 콤보상자

위와 마찬가지로, 콤보상자에 데이터도 미리 폼창이 실행되었을 경우 셋팅되어야 할 경우,


  Private Sub UserForm_activate()
    with com거주지
      .additem "역삼동"
      .additem "도곡동"
      .additem "삼성동"
      .additem "대치동"
      .additem "기타"
    end With
  End Sub


3) 이미지 삽입

  Private Sub com사진_Click()

  Dim 사진경로 As Variant

    사진경로 = Application.GetOpenFilename
    img사진.Picture = LoadPicture(사진경로)
    txt사진경로 = 사진경로

  End Sub

4) 동글뱅이 라디오 버튼

  '라디오버튼 이름을 opt남, opt여로 이름 설정'
  If opt남 = True Then
      Worksheets("고객목록").Cells(행, 3) = "남"
  ElseIf opt여 = True Then
      Worksheets("고객목록").Cells(행, 3) = "여"
  End If

5) 체크박스
  '체크박스 이름 chk헬스,골프으로 총 2개 항목 체크박스 생성

    If chk헬스 = True Then
        Worksheets("고객목록").Cells(행, 8) = "O"
    Else
        Worksheets("고객목록").Cells(행, 8) = "X"
    End If

    If chk골프 = True Then
        Worksheets("고객목록").Cells(행, 9) = "O"
    Else
        Worksheets("고객목록").Cells(행, 9) = "X"
    End If


  4) 스핀단추 연계



  5) 체크박스 연계

  8) 연속탭은 무쓸모(겹쳐나옴). 쓰려면 다중페이지로 작성.

  9) 폼창 띄우기 닫기 예제 (버튼)

    9-1) 띄우기

      Sub 신규고객등록()

      신규고객등록.Show

      End Sub

    9-2) 닫기

      Private Sub CommandButton6_Click()

      Me.Hide

      End Sub
