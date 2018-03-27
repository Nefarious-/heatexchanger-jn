Function JnTubeSide(LD, Re)
    Select Case LD
        Case Is < 24
            JnTubeSide = "#REF!"
        Case Is = 24
            JnTubeSide = calc24(Re)
        Case Is = 48
            JnTubeSide = calc48(Re)
        Case Is = 120
            JnTubeSide = calc120(Re)
        Case Is = 240
            JnTubeSide = calc240(Re)
        Case Is = 500
            JnTubeSide = calc500(Re)
        Case 24 To 48
            JnTubeSide = calc48(Re) + (LD - 48) / (48 - 24) * (calc48(Re) - calc24(Re))
        Case 48 To 120
            JnTubeSide = calc120(Re) + (LD - 120) / (120 - 48) * (calc120(Re) - calc48(Re))
        Case 120 To 240
            JnTubeSide = calc240(Re) + (LD - 240) / (240 - 120) * (calc240(Re) - calc120(Re))
        Case 240 To 500
            JnTubeSide = calc500(Re) + (LD - 500) / (500 - 240) * (calc500(Re) - calc240(Re))
    End Select
End Function

Function JnShellSide25(Re)
    Select Case Re
        Case Is < 107.9
            JnShellSide25 = 0.7124 * Re ^ -0.576
        Case 107.9 To 100000
            JnShellSide25 = 0.4257 * Re ^ -0.466
        Case Is >= 100000
            JnShellSide25 = 34.142 * Re ^ -0.446
    End Select
End Function

Private Function calc24(Re)
    Select Case Re
        Case 10 To 2000
            calc24 = 0.6094 * Re ^ -0.649
        Case 2000 To 16000
            calc24 = -4.119E-27 * Re ^ 6 + 2.06905295E-22 * Re ^ 5 - 3.923006979957E-18 * Re ^ 4 + 3.39469830867911E-14 * Re ^ 3 - 1.18367065820436E-10 * Re ^ 2 - 7.45757747191883E-09 * Re + 4.62687996111435E-03
        Case Is >= 16000
            calc24 = 0.0276 * Re ^ -0.2
    End Select
End Function

Private Function calc48(Re)
    Select Case Re
        Case 10 To 2000
            calc48 = 0.5043 * Re ^ -0.651
        Case 2000 To 16000
            calc48 = -1.281E-27 * Re ^ 6 + 9.0399489E-23 * Re ^ 5 - 2.557154365649E-18 * Re ^ 4 + 3.69434807350357E-14 * Re ^ 3 - 2.8585596592455E-10 * Re ^ 2 + 1.11506556129343E-06 * Re + 2.30433714941214E-03
        Case Is >= 16000
            calc48 = 0.0276 * Re ^ -0.2
    End Select
End Function

Private Function calc120(Re)
    Select Case Re
        Case 10 To 2000
            calc120 = 0.3455 * Re ^ -0.631
        Case 2000 To 16000
            calc120 = -8.849E-27 * Re ^ 6 + 5.42846821E-22 * Re ^ 5 - 1.3258109349244E-17 * Re ^ 4 + 1.64191224375176E-13 * Re ^ 3 - 1.09006708961713E-09 * Re ^ 2 + 3.75750423938426E-06 * Re - 1.4666513610526E-03
        Case Is >= 16000
            calc120 = 0.0276 * Re ^ -0.2
    End Select
End Function

Private Function calc240(Re)
    Select Case Re
        Case 10 To 2000
            calc240 = 0.2628 * Re ^ -0.642
        Case 2000 To 16000
            calc240 = 3.8081059E-23 * Re ^ 5 - 1.970956883017E-18 * Re ^ 4 + 4.11137393127386E-14 * Re ^ 3 - 4.39401737179716E-10 * Re ^ 2 + 2.41902839156989E-06 * Re - 1.38196072777863E-03
        Case Is >= 16000
            calc240 = 0.0276 * Re ^ -0.2
    End Select
End Function

Private Function calc500(Re)
    Select Case Re
        Case 10 To 2000
            calc500 = 0.1971 * Re ^ -0.643
        Case 2000 To 16000
            calc500 = -2.7697E-26 * Re ^ 6 + 1.524690502E-21 * Re ^ 5 - 3.3207187518495E-17 * Re ^ 4 + 3.65737441270605E-13 * Re ^ 3 - 2.16802244579708E-09 * Re ^ 2 + 6.81536425904119E-06 * Re - 5.89440423703533E-03
        Case Is >= 16000
            calc500 = 0.0276 * Re ^ -0.2
    End Select
End Function
