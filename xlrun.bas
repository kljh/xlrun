Attribute VB_Name = "xlrun"
' To Save in %AppData%\Microsoft\AddIns to load automatically
' ThisWorkbook.Workbook_Open will call XlRun routine below.

Option Explicit

Sub XlRun()
    ' Read XLRUN and XLRUN_OUT environment varibale
    ' - XLRUN should contain command formated as : -xlFileOpen MyWorkbook.xlsx  -xlRefreshLeftToRight  -xlRngGet Summary!TestStatus
    ' - XLRUN_OUT should contain the output file (default C:\temp\xlrun.out)
    
    Dim tokens() As String
    If Environ("XLRUN") = "" Then
        ' Nothing to do
        'Exit Sub
        tokens = Split("-xlFileOpen MyWorkbook.xlsx  -xlRefreshLeftToRight  -xlRngGet Summary!TestStatus ", " ")
        tokens = Split("-xlFileOpen MyMacrobook.xlsm  -xlEvalMacro MyMacro  -xlRngGet Summary!B4  -xlFileSave", " ")
        tokens = Split("-xlFileNew  -xlRngSet A1 1.0  -xlRngGet A1  -xlRngSet A2 =today()  -xlRngGet A2 -xlFileSaveAs MyGeneratedBook.xlsx", " ")
    Else
        tokens = Split(Environ("XLRUN"), " ")
    End If
    
    Dim fout_path As String, fout
    fout_path = Environ("XLRUN_OUT")
    If fout_path = "" Then fout_path = "C:\temp\xlrun.out"
    If fout_path <> "" Then
        ' Pure VBA
        ' fout = FreeFile
        ' Open Environ("XLRUN_OUT") For Output As #fout
        ' Write #fout, "xlRun"
        
        ' Using COM
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set fout = fso.CreateTextFile(fout_path, True)
        Call fout.WriteLine("xlrun")
    End If
    
    Dim i As Long, token As String, cmd As String
    For i = LBound(tokens) To UBound(tokens)
        If Left(tokens(i), 1) = "-" Then
            
            cmd = Mid(tokens(i), 2)
            
            If cmd = "xlFileOpen" Or cmd = "xlFilePath" Then
                Call xlFileOpen(fout, tokens(i + 1))
            ElseIf cmd = "xlFileNew" Then
                Call xlFileNew(fout)
            ElseIf cmd = "xlFileSave" Then
                Call xlFileSave(fout)
            ElseIf cmd = "xlFileSaveAs" Then
                Call xlFileSaveAs(fout, tokens(i + 1))
            ElseIf cmd = "xlRefreshLeftToRight" Then
                Call xlRefreshLeftToRight(fout)
            ElseIf cmd = "xlEvalMacro" Then
                Call xlEvalMacro(fout, tokens(i + 1))
            ElseIf cmd = "xlRngSet" Then
                Call xlRngSet(fout, tokens(i + 1), tokens(i + 2))
            ElseIf cmd = "xlRngGet" Then
                Call xlRngGet(fout, tokens(i + 1))
            Else
                Debug.Assert False
            End If
            
        End If
    Next i
    
    ActiveWorkbook.Close False
    PrintOut fout, "Done"
    
    ' If fout > 0 Then Close #fout
    Call fout.Close
End Sub

Function PrintOut(fout, msg As String)
    Debug.Print msg
    ' If fout > 0 Then Write #fout, msg
    fout.WriteLine msg
End Function

Function xlFileOpen(fout, xlpath As String)
    PrintOut fout, "xlFileOpen: " & xlpath
    Dim wbk As Workbook
    Set wbk = Workbooks.Open(xlpath)
    wbk.Activate
End Function

Function xlFileNew(fout)
    PrintOut fout, "xlFileNew"
    Dim wbk As Workbook
    Set wbk = Workbooks.Add
    wbk.Activate
End Function

Function xlFileSave(fout)
    PrintOut fout, "xlFileSave"
    Dim wbk As Workbook
    Set wbk = ActiveWorkbook
    wbk.Save
End Function

Function xlFileSaveAs(fout, xlpath As String)
    PrintOut fout, "xlFileSaveAs: " & xlpath
    Dim wbk As Workbook
    Set wbk = ActiveWorkbook
    Call wbk.SaveAs(xlpath)
End Function

Function xlRefreshLeftToRight(fout)
    Dim wsh As Worksheet
    For Each wsh In ActiveWorkbook.Worksheets
        PrintOut fout, "xlRefreshLeftToRight: " & wsh.Name
        wsh.Select
        wsh.Activate
    Next wsh
End Function

Function xlEvalMacro(fout, macroName As String)
    PrintOut fout, "xlEvalMacro: " & macroName
    Call Application.Run(macroName)
End Function

Function xlRngSet(fout, rngName As String, rngValue)
    PrintOut fout, "xlRngSet: " & rngName & " = " & rngValue
    Range(rngName) = rngValue
End Function

Function xlRngGet(fout, rngName As String)
    Dim rngValue
    rngValue = Range(rngName)
    PrintOut fout, "xlRngGet: " & rngName & " = " & rngValue
    ' to write to a file
End Function

