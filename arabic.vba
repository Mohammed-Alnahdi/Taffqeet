Function DITAFQEET(ByVal X As Double) As String
'******     Author Mohammed K Alnahdi        ******'
'******     License GPL 2 2022-2040   ******'
'******     mohammed-alnahdi@protonmail.com                          ******'

        Dim dblOriginal As Double
        dblOriginal = X

        Dim Letter1, Letter2, Letter3, Letter4, Letter5, Letter6 As String
        Dim c As String
        c = Format(Application.WorksheetFunction.RoundDown(X, 0), "000000000000")
        Dim C1 As Double
        C1 = Val(Mid(c, 12, 1))
        Select Case C1
            Case Is = 1: Letter1 = "واحد"
            Case Is = 2: Letter1 = "اثنين"
            Case Is = 3: Letter1 = "ثلاثة"
            Case Is = 4: Letter1 = "أربعة"
            Case Is = 5: Letter1 = "خمسة"
            Case Is = 6: Letter1 = "ستة"
            Case Is = 7: Letter1 = "سبعة"
            Case Is = 8: Letter1 = "ثمانية"
            Case Is = 9: Letter1 = "تسعة"
        End Select

        Dim C2 As Double
        C2 = Val(Mid(c, 11, 1))
        Select Case C2
            Case Is = 1: Letter2 = "عشر"
            Case Is = 2: Letter2 = "عشرون"
            Case Is = 3: Letter2 = "ثلاثون"
            Case Is = 4: Letter2 = "أربعون"
            Case Is = 5: Letter2 = "خمسون"
            Case Is = 6: Letter2 = "ستون"
            Case Is = 7: Letter2 = "سبعون"
            Case Is = 8: Letter2 = "ثمانون"
            Case Is = 9: Letter2 = "تسعون"
        End Select

        If Letter1 <> "" And C2 > 1 Then Letter2 = Letter1 + " و" + Letter2
        If Letter2 = Empty Then
            Letter2 = Letter1
        End If
        If Letter2 = "" Then
            Letter2 = Letter1
        End If
        If C1 = 0 And C2 = 1 Then Letter2 = Letter2 + "ة"
        If C1 = 1 And C2 = 1 Then Letter2 = "احدى عشر"
        If C1 = 2 And C2 = 1 Then Letter2 = "إثنى عشر"
        If C1 > 2 And C2 = 1 Then Letter2 = Letter1 + " " + Letter2
        Dim C3 As Double
        C3 = Val(Mid(c, 10, 1))
        Select Case C3
            Case Is = 1: Letter3 = "مائة"
            Case Is = 2: Letter3 = "مئتان"
            Case Is > 2: Letter3 = Left(DITAFQEET(C3), Len(DITAFQEET(C3)) - 1) + "مائة"
        End Select
        If Letter3 <> "" And Letter2 <> "" Then Letter3 = Letter3 + " و" + Letter2
        If Letter3 = "" Then Letter3 = Letter2

        Dim C4 As Double
        C4 = Val(Mid(c, 7, 3))
        Select Case C4
            Case Is = 1: Letter4 = "الف"
            Case Is = 2: Letter4 = "الفان"
            Case 3 To 10: Letter4 = DITAFQEET(C4) + " آلاف"
            Case Is > 10: Letter4 = DITAFQEET(C4) + " الف"
        End Select
        If Letter4 <> "" And Letter3 <> "" Then Letter4 = Letter4 + " و" + Letter3
        If Letter4 = "" Then Letter4 = Letter3
        Dim C5 As Double
        C5 = Val(Mid(c, 4, 3))
        Select Case C5
            Case Is = 1: Letter5 = "مليون"
            Case Is = 2: Letter5 = "مليونان"
            Case 3 To 10: Letter5 = DITAFQEET(C5) + " ملايين"
            Case Is > 10: Letter5 = DITAFQEET(C5) + " مليون"
        End Select
        If Letter5 <> "" And Letter4 <> "" Then Letter5 = Letter5 + " و" + Letter4
        If Letter5 = "" Then Letter5 = Letter4

        Dim C6 As Double
        C6 = Val(Mid(c, 1, 3))
        Select Case C6
            Case Is = 1: Letter6 = "مليار"
            Case Is = 2: Letter6 = "ملياران"
            Case Is > 2: Letter6 = DITAFQEET(C6) + " مليار"
        End Select
        If Letter6 <> "" And Letter5 <> "" Then Letter6 = Letter6 + " و" + Letter5
        If Letter6 = "" Then Letter6 = Letter5
        
        Dim x2 As Double
        x2 = dblOriginal * 100
        
        
        DITAFQEET = Letter6 + IIf(CInt(Right(CStr(x2), 3)) <> 0, " ريال سعودي و" + Right(CStr(x2), 2) + " هللة ", "")

    End Function



