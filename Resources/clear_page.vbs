Attribute VB_Name = "clear_page"
Sub wipe():

    Dim ws As Integer
    For ws = 1 To Worksheets.Count
    Worksheets(ws).Select
    
    Range("I1:L22771") = " "
    Range("I1:L22771").Interior.ColorIndex = 0
    Range("N1:P5") = " "
    
    Next ws
    
End Sub
