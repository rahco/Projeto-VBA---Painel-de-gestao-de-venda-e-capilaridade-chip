Attribute VB_Name = "Módulo1"
Sub Geral()

    Application.ScreenUpdating = False

    Call Base_Vendas
    Call Status_de_abastecimento_chip
    Call Base_de_vendas_envio

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
    ActiveWorkbook.RefreshAll

    Application.ScreenUpdating = True

End Sub


Sub Status_de_abastecimento_chip()

    Application.ScreenUpdating = False

    Sheets("STATUS DE ABASTECIMENTO CHIP").Select
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

Sub Base_de_vendas_envio()

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
    Sheets("METAS").Visible = True
    Sheets("METAS").Select
    Sheets("DE-PARA CHIP").Visible = True
    Sheets("DE-PARA CHIP").Select
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
    Sheets("BASE RMV ABAS. CHIP").Select
    Application.CutCopyMode = False
 
    Sheets(Array("BASE RMV ABAS. CHIP", "HC", "METAS", "DE-PARA CHIP", _
        "MACROS", "BASE DIAS", "BD - BV", "BV INICIAL", _
        "BD VENDAS CHIP", _
        "TD - VENDAS CHIP", "TD - STATUS DE ABASTECIMENTO", "GRÁFICO DE ENVIO")).Select
    
    ActiveWindow.SelectedSheets.Delete
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







