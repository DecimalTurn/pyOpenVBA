Attribute VB_Name = "Module1"

'with_class fixture: exercises a class module and writes a sentinel file.
Sub RunFixture()
    Dim obj As New Class1
    obj.Message = "it works"

    Dim outPath As String
    outPath = ThisWorkbook.Path & "\" & Replace(ThisWorkbook.Name, ".xlsm", "") & ".txt"

    Open outPath For Output As #1
    Print #1, obj.Greet()
    Close #1
End Sub
