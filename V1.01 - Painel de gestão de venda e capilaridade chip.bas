Attribute VB_Name = "Módulo1"
Sub Geral()

    Application.ScreenUpdating = False

    Call Base_Vendas
    Call Chave_Con
    Call Base_Pistolagem
    Call Base_Capilaridade
    Call TD
    Call BD_Envio

    Sheets("MACROS").Select
    Range("B7").Select

    Application.ScreenUpdating = True

End Sub


Sub Base_Vendas()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BV INICIAL").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BD - BV").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BV INICIAL").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("O4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("O5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    Sheets("BD VENDAS CHIP").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
    
    Application.CutCopyMode = False
    Range("B5").Select
    Sheets("BV INICIAL").Select
    Range("O4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B4").Select
    Sheets("BD VENDAS CHIP").Select
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("N5:Q5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B5").Select
    Application.CutCopyMode = False

    Application.ScreenUpdating = True

End Sub

Sub Chave_Con()

    Application.ScreenUpdating = False

    Sheets("CHAVE - CON").Select
    
    Rows("1:1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Sheets("BD - CON").Columns("D:D").Copy
    Sheets("CHAVE - CON").Select
    Columns("B:B").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("C:XFD").Select
    Selection.Delete Shift:=xlToLeft
    Sheets("BD - CON").Select
    Columns("E:E").Select
    Selection.Copy
    Range("B6").Select
    Sheets("CHAVE - CON").Select
    Columns("C:C").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("D:D").Select
    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlFixedWidth, _
        OtherChar:="-", FieldInfo:=Array(Array(0, 1), Array(1, 1)), _
        TrailingMinusNumbers:=True
    Range("C5").Select
    Selection.Cut
    Range("E5").Select
    ActiveSheet.Paste
    Columns("C:D").Select
    Selection.Delete Shift:=xlToLeft
    Range("D5").Value = "Chave"
    Range("D6").Value = "=RC[-2]&RC[-1]"
    Range("E6").Value = "Sim"
    Range("D6:E6").Select
    Selection.Copy
    Range("C6").Select
    Selection.End(xlDown).Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("2:4").Select
    Selection.Delete Shift:=xlUp
    Cells.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.EntireColumn.AutoFit
    Range("B3").Select
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True

End Sub

Sub Base_Pistolagem()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BP INICIAL").Select
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "'CHAVE - CON'!D:E"
    Range("M4").Select
    ActiveCell.FormulaR1C1 = _
    "=IFERROR(VLOOKUP(RC[-1],'CHAVE - CON'!C[-9]:C[-8],2,FALSE),""Não"")"
    Range("M5").Select
    Sheets("BP INICIAL").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
    
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BD - BP").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BP INICIAL").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    Sheets("BD PISTOLAGEM CHIP").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
    
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BP INICIAL").Select
    ActiveSheet.Range("$B$3:$N$24569").AutoFilter Field:=13, Criteria1:="=1", _
        Operator:=xlAnd
    Range("B3:I3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BD PISTOLAGEM CHIP").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("K5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    Sheets("BP INICIAL").Select
    ActiveSheet.Range("$B$3:$N$24569").AutoFilter Field:=13
    Range("B4").Select
    Sheets("BD PISTOLAGEM CHIP").Select
 
    Application.ScreenUpdating = True
 
End Sub

Sub Base_Capilaridade()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BD CAP").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("B2").Value > 0 Then
        linhaf = linhai - Range("B2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("B2").Value < 0 Then
        linhaf = linhai + Range("B2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If

    Application.CutCopyMode = False
    Range("B5").Select
    Range("B5:M5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B5").Select
    Sheets("BD VENDAS CHIP").Select
    Range("B5:M5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B5").Select
    Sheets("BD CAP").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Sheets("BD PISTOLAGEM CHIP").Select
    Range("J4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B4").Select
    Sheets("BD CAP").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("N5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("N6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B5").Select
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True

End Sub

Sub TD()

    Application.ScreenUpdating = False

    Sheets("STATUS DE ABASTECIMENTO CHIP").Select
    ActiveWorkbook.RefreshAll
    Range("N6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("N7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Application.CutCopyMode = False

    Application.ScreenUpdating = True

End Sub

Sub BD_Envio()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim linhai As Double
    Dim linhaf As Double

    Sheets("BASE DE VENDAS").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If

    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BASE DE PISTOLAGEM").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
    
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BD VENDAS CHIP").Select
    Range("B5:M5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B5").Select
    Sheets("BASE DE VENDAS").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    Sheets("BD PISTOLAGEM CHIP").Select
    Range("J4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B4").Select
    Sheets("BASE DE PISTOLAGEM").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True

End Sub

Sub Arquivo_Envio()

    Application.ScreenUpdating = False
    
    ActiveWorkbook.Save
    ChDir _
        ActiveWorkbook.Path
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\" & Worksheets("MACROS").Range("C12").Value & " - Gestão de Abastecimento e Venda Chip - Dados até dia " & Worksheets("MACROS").Range("C13").Value & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Sheets("MACROS").Select
    Sheets("BASE RMV ABAS. CHIP").Visible = True
    Sheets("BASE RMV ABAS. CHIP").Select
    Sheets("HC").Visible = True
    Sheets("HC").Select
    Sheets("METAS - CAP").Visible = True
    Sheets("METAS - CAP").Select
    Sheets("METAS").Visible = True
    Sheets("METAS").Select
    Sheets("DE-PARA CHIP").Visible = True
    Sheets("DE-PARA CHIP").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=16
    Sheets("QUADRO DE PERFORMANCE").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("STATUS DE ABASTECIMENTO CHIP").Select
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.ScrollWorkbookTabs Sheets:=-16
    Sheets("BASE RMV ABAS. CHIP").Select
    Application.CutCopyMode = False
    Sheets(Array("BASE RMV ABAS. CHIP", "HC", "METAS - CAP", "METAS", "DE-PARA CHIP", _
        "MACROS", "BASE DIAS", "BD - BV", "BD - CON", "BD - BP", "BV INICIAL")).Select
    Sheets("BASE RMV ABAS. CHIP").Activate
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    Sheets(Array("BASE RMV ABAS. CHIP", "HC", "METAS - CAP", "METAS", "DE-PARA CHIP", _
        "MACROS", "BASE DIAS", "BD - BV", "BD - CON", "BD - BP", "BV INICIAL", _
        "BD VENDAS CHIP", "CHAVE - CON", "BP INICIAL", "BD PISTOLAGEM CHIP", "BD CAP", _
        "TD - VENDAS CHIP")).Select
    Sheets("BASE RMV ABAS. CHIP").Activate
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    Sheets(Array("BASE RMV ABAS. CHIP", "HC", "METAS - CAP", "METAS", "DE-PARA CHIP", _
        "MACROS", "BASE DIAS", "BD - BV", "BD - CON", "BD - BP", "BV INICIAL", _
        "BD VENDAS CHIP", "CHAVE - CON", "BP INICIAL", "BD PISTOLAGEM CHIP", "BD CAP", _
        "TD - VENDAS CHIP", "TD - STATUS DE ABASTECIMENTO")).Select
    Sheets("BASE RMV ABAS. CHIP").Activate
    ActiveWindow.ScrollWorkbookTabs Sheets:=2
    Sheets(Array("BASE RMV ABAS. CHIP", "HC", "METAS - CAP", "METAS", "DE-PARA CHIP", _
        "MACROS", "BASE DIAS", "BD - BV", "BD - CON", "BD - BP", "BV INICIAL", _
        "BD VENDAS CHIP", "CHAVE - CON", "BP INICIAL", "BD PISTOLAGEM CHIP", "BD CAP", _
        "TD - VENDAS CHIP", "TD - STATUS DE ABASTECIMENTO", "GRÁFICO DE ENVIO")).Select
    Sheets("GRÁFICO DE ENVIO").Activate
    ActiveWorkbook.Connections("WorksheetConnection_BD CAP!$B$4:$M$9139").Delete
    ActiveWindow.SelectedSheets.Delete
    Range("B1").Select
    Sheets("BASE DE PISTOLAGEM").Select
    Range("A1").Select
    Selection.Copy
    Range("B1:D1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B4").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("BASE DE VENDAS").Select
    Range("A1").Select
    Selection.Copy
    Range("B1:C1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B4").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("STATUS DE ABASTECIMENTO CHIP").Select
    Range("AK6").Select
    Range("B6").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("QUADRO DE PERFORMANCE").Select
    Range("B7").Select
    ActiveWindow.DisplayHeadings = False
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True

End Sub







