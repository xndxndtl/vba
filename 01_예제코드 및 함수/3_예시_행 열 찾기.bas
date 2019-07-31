
0) 고급필터 동적 범위설정 예시

[중요!] 고급필터 매크로 사용 시, 사전에 아래와 같이 빈시트? 같은 것을 하나 클릭해야됨.
        그렇게 안하면 무슨 추출 범위 어쩌구 에러 남.(왜나는지 잘 모르겠음.)

        Sheets("인쇄").Select   '--> 고급필터 에러 방지를 위한 더미 시트 셀렉트

        Sheets("고객목록").Range("A1", _
            Worksheets("고객목록").Range("K10000").End(xlUp).Address). _
            AdvancedFilter Action:=xlFilterCopy, _
            CriteriaRange:=Sheets("인쇄필터조건").Range("A1:D2"), CopyToRange:=Range("A5:G5") _
            , Unique:=False
---------------------------------------------------------------------------------
1) 시작셀 및 끝 칼럼열을 알 때, 마지막셀의 위치 찾기
    아래 예시는 고급필터 범위 설정 시 A2(시작셀) 기준으로 J열의 끝까지 설정하기 위함임.

    Range("A2", Range("j100000").End(xlUp).Address). _
    AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= Range("B12:J13"), Unique:=False

--------------------------------------------------------------------------------

2) 찾고자 하는 열, 행번호 찾기
  Function Find_Direction(Search_rng As Range, Search_Item As Variant, Direction As Integer)

      Dim rng As Range
      Dim str_Date As Date

      If IsDate(Search_Item) = True Then
          str_Date = DateSerial(Year(Search_Item), Month(Search_Item), Day(Search_Item))
          Set rng = Search_rng.Find(what:=str_Date, lookat:=xlWhole)
      Else
          Set rng = Search_rng.Find(what:=Search_Item, lookat:=xlWhole)
      End If

      If Not rng Is Nothing Then
          Find_Direction = IIf(Direction = 1, rng.Row, rng.Column)
      Else
          Find_Direction = 0
      End If
  End Function

--------------------------------------------------------------------------------

3) 중복값 제거

  Function Remove_Duplicate(rng As Range) As Variant

      Dim i As Double
      Dim C As New Collection
      Dim Var_Result() As Variant

      On Error Resume Next
      For Each rng In rng
          If rng.Value <> "" Then
              C.Add CStr(rng.Value), CStr(rng.Value)
          End If
      Next
      On Error GoTo 0

      If C.Count > 0 Then
          ReDim Preserve Var_Result(1 To C.Count)
          For i = 1 To C.Count
              Var_Result(i) = C.Item(i)
          Next i

          Remove_Duplicate = Var_Result
      Else
          Remove_Duplicate = -1
      End If

  End Function
'사용법

'========================================================
'Var_Db = Remove_Duplicate(Range(Cells(2, 4), Cells(15, 4)))
'If IsArray(Var_Db) = True Then
'    For i = 1 To UBound(Var_Db)

'    Next i
'End If
'========================================================

--------------------------------------------------------------------------------
4) 행의 개수 반환

  Function Count_Rows(Optional str_sheet As String, Optional str_Temp As String)

      Dim cntR As Double

      If str_sheet = "" Then
          cntR = Cells(Rows.Count, str_Temp).End(xlUp).Row
      Else
          cntR = Sheets(str_sheet).Cells(Rows.Count, str_Temp).End(xlUp).Row
      End If

      Count_Rows = cntR
  End Function

  '4. 파일 선택 주소 반환
  Function File_Select() As Variant

      Dim i As Double, i2 As Double
      Dim Fd As FileDialog
      Dim Var_File() As Variant

      Set Fd = Application.FileDialog(msoFileDialogFilePicker)
      With Fd
          .AllowMultiSelect = True
          .Show

          If .SelectedItems.Count = 0 Then
              File_Select = -1
              Exit Function
          End If
      End With

      ReDim Preserve Var_File(1 To Fd.SelectedItems.Count)
      For i = 1 To Fd.SelectedItems.Count
          Var_File(i) = Fd.SelectedItems(i)
      Next i

      File_Select = Var_File
  End Function

  Sub Create_Sheet(시트명 As String)
      Sheets.Add(after:=Sheets(Sheets.Count)).Name = 시트명
  End Sub

  Sub Copy_Sheet(시트명 As String, 변경시트명 As String)

      Dim ws As Worksheet

      If 시트명 = "" Then
          MsgBox "시트명이 공백입니다.", vbCritical, ""
          Exit Sub
      ElseIf 변경시트명 = "" Then
          MsgBox "변경 시트명이 공백입니다.", vbCritical, ""
          Exit Sub
      ElseIf 시트명 = 변경시트명 Then
          MsgBox "시트명과 변경시트명은 달라야 합니다.", vbCritical, ""
          Exit Sub
      End If

      Application.DisplayAlerts = False
      On Error Resume Next
      Set ws = Sheets(변경시트명)
          ws.Delete
      On Error GoTo 0

      Sheets(시트명).Copy after:=Sheets(Sheets.Count)
      ActiveSheet.Name = 변경시트명
      ActiveSheet.Tab.Color = vbBlue

  End Sub

  Sub Delete_Sheet_Keyword(str_Keyword As String)
      Dim i As Double

      Application.DisplayAlerts = False
      For i = Sheets.Count To 1 Step -1
          If InStr(Sheets(i).Name, str_Keyword) > 0 Then
              Sheets(i).Delete
          End If
      Next i
      Application.DisplayAlerts = True
  End Sub

  Sub Delete_Sheet_Color(Val_Col As Double)
      Dim i As Double

      Application.DisplayAlerts = False
      For i = Sheets.Count To 1 Step -1
          If Sheets(i).Tab.Color = Val_Col Then
              Sheets(i).Delete
          End If
      Next i
      Application.DisplayAlerts = True
  End Sub
