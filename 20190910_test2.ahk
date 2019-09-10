
gui, add, text, x30 y5 w110 h20, 매크로 프로그램
gui, add, text, x60 y25 w50 h20 vA, 준비!!
gui, add, text, x60 y50 h20 w50 vB, 0회
gui, add, button, x20 y80 w110 h20, 시작
gui, add, button, x20 y110 w110 h20, 종료
gui, show

return

button시작:
{
	매크로시작 :=true
	Loop
	{
		;여기에 무한 반복할 작업의 코드를 작성합니다.
		;msgbox, 고우 
		ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *50 에러.png
		;msgbox, %errorlevel%
		if (errorlevel = 0)
		{
			;msgbox, 뒤로가기 눌러야 하는 상황
			ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *50 크롬뒤로가기.bmp
			if (errorlevel = 0)
			{
				;msgbox, 뒤로가기 버튼 탐색 완료
				sleep, 1000
				send {click %foundX% %foundY%}
				MouseMove, 100, 100
			}
			
			
		}
		if (errorlevel = 1)
		{
			;msgbox, 암것도 못찾음
			sleep, 20000			
		}
	}

}

button종료:
{
	매크로시작 := false
	exitapp
}
return
