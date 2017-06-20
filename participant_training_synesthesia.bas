Attribute VB_Name = "Module1"
Sub Letter2Colors_Participant()

' Letter2Colors Participant Example
' Simple search and replace

    Dim a_color As Long
    Dim e_color As Long
    Dim n_color As Long
    Dim r_color As Long

    a_color = RGB(255, 143, 0) 'orange
    e_color = RGB(0, 155, 255) 'blue
    n_color = RGB(230, 0, 0) 'red
    r_color = RGB(0, 181, 0) 'green
  

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "a"
        .Replacement.Text = "a"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchByte = False
        .CorrectHangulEndings = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Replacement.Font.Color = a_color
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "e"
        .Replacement.Text = "e"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchByte = False
        .CorrectHangulEndings = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Replacement.Font.Color = e_color
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "n"
        .Replacement.Text = "n"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchByte = False
        .CorrectHangulEndings = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Replacement.Font.Color = n_color
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
  
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "r"
        .Replacement.Text = "r"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchByte = False
        .CorrectHangulEndings = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Replacement.Font.Color = r_color
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   
    
End Sub
