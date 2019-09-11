
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
----
gui, add, text, x30 y5 w110 h20, 매크로 프로그램
gui, add, text, x60 y25 w50 h20 vA, 준비!!
gui, add, text, x60 y50 h20 w50 vB, 0회
gui, add, button, x20 y80 w110 h20, 시작
gui, add, button, x20 y110 w110 h20, 종료
gui, show

return


BUTTON시작:
{
	매크로시작 :=true
	sleep, 3000
	Loop
	{
		;에러 시 뒤로가기


		ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 조회하기.png
		msgbox,,,조회하기_%errorlevel%,1

		if (errorlevel = 0)
		{
			send {click %foundX% %foundY%}
			sleep, 2500
			; 스크롤 움직여야함.

			ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예매.png
			msgbox,,,예매_%errorlevel%,1

					if (errorlevel = 0)
					{
								send {click %foundX% %foundY%}
								sleep, 2500

								ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예매계속진행하기.png
								msgbox,,,예매진행_%errorlevel%,1

								if (errorlevel = 0)
								{
											send {click %foundX% %foundY%}
											sleep, 2500

											ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예약확인.png
											if (errorlevel = 0)
											{
														send {click %foundX% %foundY%}
														sleep, 2500
											}
								}
					}

			sleep, 1500

			ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예매.png
			msgbox,,,예매_%errorlevel%,1

					if (errorlevel = 0)
					{
								send {click %foundX% %foundY%}
								sleep, 2500

								ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예매계속진행하기.png
								msgbox,,,예매진행_%errorlevel%,1

								if (errorlevel = 0)
								{
											send {click %foundX% %foundY%}
											sleep, 2500

											ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예약확인.png
											if (errorlevel = 0)
											{
														send {click %foundX% %foundY%}
														sleep, 2500
														exitapp
											}
								}
					}

					sleep, 1500
					ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예매.png
					msgbox,,,예매_%errorlevel%,1

							if (errorlevel = 0)
							{
										send {click %foundX% %foundY%}
										sleep, 2500

										ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예매계속진행하기.png
										msgbox,,,예매진행_%errorlevel%,1

										if (errorlevel = 0)
										{
													send {click %foundX% %foundY%}
													sleep, 2500

													ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예약확인.png
													if (errorlevel = 0)
													{
																send {click %foundX% %foundY%}
																sleep, 2500
													}
										}
							}

							sleep, 1000
							ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예매.png
							msgbox,,,예매_%errorlevel%,1

									if (errorlevel = 0)
									{
												send {click %foundX% %foundY%}
												sleep, 2500

												ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예매계속진행하기.png
												msgbox,,,예매진행_%errorlevel%,1

												if (errorlevel = 0)
												{
															send {click %foundX% %foundY%}
															sleep, 2500

															ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, 예약확인.png
															if (errorlevel = 0)
															{
																		send {click %foundX% %foundY%}
																		sleep, 2500
															}
												}
									}


		}
		if(errorlevel = 1)
		{
			msgbox,,,조회하기 실패,1
			sleep, 1000

			send {click, 941, 607 }

			ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *50 에러.bmp
			;msgbox, %errorlevel%
			if (errorlevel = 0)
			{
				msgbox,,,에러검출,1
				ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *50 크롬뒤로가기.bmp

				if (errorlevel = 0)
				{
					;msgbox, 뒤로가기 버튼 탐색 완료
					sleep, 1000
					send {click %foundX% %foundY%}
					MouseMove, 100, 100
				}

			}
		}
	}
}

;F3::
BUTTON종료:
{
;	매크로시작 := false
	exitapp
}
return

