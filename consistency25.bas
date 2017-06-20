Attribute VB_Name = "Module1"
Sub consistency25()

' Letter2Colors Consistency 25
' Simple search and replace

    Dim rndNumber As Integer
    rndNummer = 1

    Dim a_color_1 As Long
    Dim a_color_2 As Long
    Dim a_color_3 As Long
    Dim a_color_4 As Long
    Dim e_color_1 As Long
    Dim e_color_2 As Long
    Dim e_color_3 As Long
    Dim e_color_4 As Long
    Dim n_color_1 As Long
    Dim n_color_2 As Long
    Dim n_color_3 As Long
    Dim n_color_4 As Long
    Dim r_color_1 As Long
    Dim r_color_2 As Long
    Dim r_color_3 As Long
    Dim r_color_4 As Long
    
    ' colors for A
    a_color_1 = RGB(230, 0, 0) 'red
    a_color_2 = RGB(255, 143, 0) 'orange
    a_color_3 = RGB(0, 181, 0) 'green
    a_color_4 = RGB(0, 155, 255) 'blue
    
    'colors for E
    e_color_1 = RGB(230, 0, 0) 'red
    e_color_2 = RGB(255, 143, 0) 'orange
    e_color_3 = RGB(0, 181, 0) 'green
    e_color_4 = RGB(0, 155, 255) 'blue
    
    ' colors for N
    n_color_1 = RGB(230, 0, 0) 'red
    n_color_2 = RGB(255, 143, 0) 'orange
    n_color_3 = RGB(0, 181, 0) 'green
    n_color_4 = RGB(0, 155, 255) 'blue
    
    ' colors for R
    r_color_1 = RGB(230, 0, 0) 'red
    r_color_2 = RGB(255, 143, 0) 'orange
    r_color_3 = RGB(0, 181, 0) 'green
    r_color_4 = RGB(0, 155, 255) 'blue


        ' this is for the letter a
        With Selection.Find
            .ClearFormatting
            .Forward = True
            .Wrap = wdFindContinue
            .Text = "a"
            .Format = True
            .MatchCase = True
            .MatchWholeWord = False
            .MatchByte = False
            .CorrectHangulEndings = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Replacement.Font.Color = wdColorDarkRed
        End With
        Do While Selection.Find.Execute
            
            rndNumber = CInt(Int((100 * Rnd()) + 1))    ' Generate random value between 1 and 100.
            
            If rndNumber <= 25 Then
                Selection.Font.Color = a_color_1
            ElseIf rndNumber > 25 And rndNumber <= 50 Then
                Selection.Font.Color = a_color_2
            ElseIf rndNumber > 50 And rndNumber <= 75 Then
                Selection.Font.Color = a_color_3
            Else
                Selection.Font.Color = a_color_4
            End If
                
        Loop
        
        ' this is for the letter e
        With Selection.Find
            .ClearFormatting
            .Forward = True
            .Wrap = wdFindContinue
            .Text = "e"
            .Format = True
            .MatchCase = True
            .MatchWholeWord = False
            .MatchByte = False
            .CorrectHangulEndings = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Replacement.Font.Color = wdColorDarkRed
        End With
        Do While Selection.Find.Execute
            
            rndNumber = CInt(Int((100 * Rnd()) + 1))    ' Generate random value between 1 and 100.
            
            If rndNumber <= 25 Then
                Selection.Font.Color = e_color_1
            ElseIf rndNumber > 25 And rndNumber <= 50 Then
                Selection.Font.Color = e_color_2
            ElseIf rndNumber > 50 And rndNumber <= 75 Then
                Selection.Font.Color = e_color_3
            Else
                Selection.Font.Color = e_color_4
            End If
                
        Loop
        
        ' this is for the letter n
        With Selection.Find
            .ClearFormatting
            .Forward = True
            .Wrap = wdFindContinue
            .Text = "n"
            .Format = True
            .MatchCase = True
            .MatchWholeWord = False
            .MatchByte = False
            .CorrectHangulEndings = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Replacement.Font.Color = wdColorDarkRed
        End With
        Do While Selection.Find.Execute
            
            rndNumber = CInt(Int((100 * Rnd()) + 1))    ' Generate random value between 1 and 100.
            
            If rndNumber <= 25 Then
                Selection.Font.Color = n_color_1
            ElseIf rndNumber > 25 And rndNumber <= 50 Then
                Selection.Font.Color = n_color_2
            ElseIf rndNumber > 50 And rndNumber <= 75 Then
                Selection.Font.Color = n_color_3
            Else
                Selection.Font.Color = n_color_4
            End If
                
        Loop
        
        ' this is for the letter r
        With Selection.Find
            .ClearFormatting
            .Forward = True
            .Wrap = wdFindContinue
            .Text = "r"
            .Format = True
            .MatchCase = True
            .MatchWholeWord = False
            .MatchByte = False
            .CorrectHangulEndings = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Replacement.Font.Color = wdColorDarkRed
        End With
        Do While Selection.Find.Execute
            
            rndNumber = CInt(Int((100 * Rnd()) + 1))    ' Generate random value between 1 and 100.
            
            If rndNumber <= 25 Then
                Selection.Font.Color = r_color_1
            ElseIf rndNumber > 25 And rndNumber <= 50 Then
                Selection.Font.Color = r_color_2
            ElseIf rndNumber > 50 And rndNumber <= 75 Then
                Selection.Font.Color = r_color_3
            Else
                Selection.Font.Color = r_color_4
            End If
                
        Loop
        
            
End Sub


