#시작셀을 알고 마지막셀의 위치를 찾는 것

# 1) 아래 예시는 고급필터 범위 설정 시 A2(시작셀) 기준으로 J열의 끝까지 설정하기 위함임.

    Range("A2", Range("j100000").End(xlUp).Address). _
    AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= Range("B12:J13"), Unique:=False
