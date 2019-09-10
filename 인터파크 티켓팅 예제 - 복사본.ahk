RRR()

{

ImageSearch,vx,vy,14,101,1814,986, *40 예매하기.png  

if errorlevel=0

{

mouseclick,left,%vx%,%vy%

sleep, 200

ImageSearch,vx,vy,11,59,640,417, *40 인식.png  

if errorlevel=0

{

sleep, 100

mouseclick,left,157,277

sleep, 100

mouseclick,left,791,604

sleep, 200

ImageSearch,vx,vy,11,53,657,711, *40 렉방지.png  

if errorlevel=0

{

send, {F4 down}

}

else

{

WWW()

}

 

}

else

{

EEE()

}

}

else

{

RRR()

}

}

 

 

EEE()

{

ImageSearch,vx,vy,11,59,640,417, *40 인식.png  

if errorlevel=0

{

sleep, 100

mouseclick,left,157,277

sleep, 100

mouseclick,left,791,604

sleep, 200

ImageSearch,vx,vy,11,53,657,711, *40 렉방지.png  

if errorlevel=0

{

send, {F4 down}

}

else

{

WWW()

}

 

}

else

{

EEE()

}

}

 

 

 

WWW()

{

ImageSearch,vx,vy,11,53,657,711, *40 렉방지.png  

if errorlevel=0

{

send, {F4 down}

}

else

{

WWW()

}

}

 

 

 

 

 

 

 

 

 

 

ACC()

{

sleep, 300

send, {F5}

sleep, 1000

ImageSearch,vx,vy,14,101,1814,986, *40 예매하기.png  

if errorlevel=0

{

mouseclick,left,%vx%,%vy%

sleep, 200

ImageSearch,vx,vy,11,59,640,417, *40 인식.png  

if errorlevel=0

{

sleep, 100

mouseclick,left,157,277

sleep, 100

mouseclick,left,791,604

sleep, 200

ImageSearch,vx,vy,11,53,657,711, *40 렉방지.png  

if errorlevel=0

{

send, {F4 down}

}

else

{

WWW()

}

 

}

else

{

EEE()

}

}

else

{

RRR()

}

}

 

 

 

 

 

 

 

 

 

 

ABC()

{

PixelSearch,VX,VY,19,112,671,719,0xEE687B,6,fast

if errorlevel=0

{

return

}

else

{

mouseclick,left,840,646

send, {F4 down}

 

}

}

 

AUU()

{

Loop

{

현재시간 = %A_Hour%:%A_Min%:%A_Sec%

if(현재시간 = "21:15:00")

{

send, {F6}

} 

}

}

 

 

F6::

ACC()

 

F1::

AUU()

 

F4::

Loop

{

send, {F4 up}

PixelSearch,VX,VY,19,112,671,719,0xEE687B,6,fast

if errorlevel=0

{

mouseclick,left,%vx%,%vy%

PixelSearch,VX,VY,19,112,671,719,0x3B3B3B,10,fast

if errorlevel=0

{

mouseclick,left,789,616

sleep, 500

send, {anter down}

sleep, 100

send, {anter up}

PixelSearch,VX,VY,19,112,671,719,0xEE687B,6,fast

if errorlevel=0

{

return

}

else

{

mouseclick,left,840,646

send, {F4 down}

 

}

}

else

{

mouseclick,left,789,616

sleep, 500

send, {anter down}

   sleep, 100

   send, {anter up}

ABC()

 

}

}

else

{

mouseclick,left,840,646

sleep, 500

send, {F4 down}

}

 

return

}

SetControlDelay,-1 

SetDefaultMouseSpeed,-1

SetWinDelay,-1 

SetMouseDelay,-1 

SetBatchLines,-1

 

 

 

 

F2::

reload

return

 

 

F3::

exitapp

 