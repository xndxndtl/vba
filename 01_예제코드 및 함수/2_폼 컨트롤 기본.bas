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
