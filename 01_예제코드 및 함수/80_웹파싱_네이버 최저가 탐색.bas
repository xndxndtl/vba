Option Explicit

'Made by YouCELL

Sub Main_실시간()

    Dim i As Long, i2 As Long, cntR As Long
    Dim T As String, str_Temp As Variant, str_Temp_B As String, str_Title As String, str_Price As String


    If MsgBox("정말 웹파싱을 진행하시겠습니까?", vbQuestion + vbOKCancel, "") = vbCancel Then Exit Sub
    cntR = Cells(Rows.Count, 2).End(xlUp).Row

    Range(Cells(3, 3), Cells(10000, 4)).ClearContents
    If Cells(cntR, 2).Value = "" Then cntR = Cells(cntR, 2).End(xlUp).Row

    For i = 3 To cntR
        With CreateObject("WinHttp.WinHttpRequest.5.1")
            .Open "GET", "https://search.shopping.naver.com/search/all.nhn?query=" & Cells(i, 2).Value & "&pagingIndex=1&pagingSize=40&viewType=list&sort=price_asc&frm=NVSHATC&sps=Y&query=" & Cells(i, 2).Value
            .SetRequestHeader "Cookie", "BMR=; "
            .Send
            .WaitForResponse: DoEvents
            T = .ResponseText
        End With

        On Error Resume Next
        str_Temp = Split(T, "<div class=""info"">")
        For i2 = 1 To 1
            str_Title = Split(Split(str_Temp(i2), """ class")(0), "title=""")(1)
            str_Price = Split(Split(Split(str_Temp(i2), "data-reload-date")(1), "</span>")(0), ">")(1)
        Next i2

        If Err.Number = 0 Then
            Cells(i, 3).Value = str_Title
            Cells(i, 4).Value = str_Price
        Else
            Cells(i, 3).Value = "검색 안됨"
            Cells(i, 4).Value = ""
        End If

        On Error GoTo 0
    Next i
    MsgBox "상품이 업데이트 되었습니다.", vbInformation, ""

End Sub
