Private Sub UserForm_Initialize()
Application.Visible = False

limpar2
atualizacombos
atualizalista

With TabelaDeMovimentações
    .BorderStyle = ccFixedSingle
    .Gridlines = True
    .View = lvwReport
    .ColumnHeaders.Add Text:="Item", Width:=190
    .ColumnHeaders.Add Text:="Movimentação", Width:=50, Alignment:=2
    .ColumnHeaders.Add Text:="Quantidade", Width:=50, Alignment:=2
    .ColumnHeaders.Add Text:="Lote", Width:=60, Alignment:=2
    .ColumnHeaders.Add Text:="Validade", Width:=60, Alignment:=2
    .ColumnHeaders.Add Text:="Câmara", Width:=60, Alignment:=2
    .ColumnHeaders.Add Text:="Data de movimentação", Width:=125, Alignment:=2
    .ColumnHeaders.Add Text:="ID", Width:=61, Alignment:=2

End With

With ListView_Estoque
    .BorderStyle = ccFixedSingle
    .Gridlines = True
    .View = lvwReport
    .ColumnHeaders.Add Text:="Produtos", Width:=259
    .ColumnHeaders.Add Text:="Câmaras", Width:=79, Alignment:=2
End With

With TabelaDeInventario
    .BorderStyle = ccFixedSingle
    .Gridlines = True
    .View = lvwReport
    .ColumnHeaders.Add Text:="Item", Width:=259
    .ColumnHeaders.Add Text:="Lote", Width:=80, Alignment:=2
    .ColumnHeaders.Add Text:="Validade", Width:=80, Alignment:=2
    .ColumnHeaders.Add Text:="Prazo", Width:=80, Alignment:=2
    .ColumnHeaders.Add Text:="Câmara", Width:=80, Alignment:=2
    .ColumnHeaders.Add Text:="Quantidade (Kg)", Width:=77, Alignment:=2
End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ThisWorkbook.Save
    ThisWorkbook.Close
    Application.Quit
End Sub


Private Sub MultiPage1_Click(ByVal Index As Long)
    limpar2
End Sub


'============================================================================================================
'============================================================================================================
'=================================================BOTÕES=====================================================
'============================================================================================================
'============================================================================================================
Private Sub InventaPrazo1_Click()
Worksheets("Inventário").Select

    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=4, Criteria1:= _
        ">6", Operator:=xlAnd, Criteria2:="<=10"
        
Sheets("Inventário").Select
    
    TabelaDeInventario.ListItems.Clear
    
    
    lin = 2
    
    Do Until Sheets("Inventário").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
        Set li = TabelaDeInventario.ListItems.Add(Text:=Sheets("Inventário").Cells(lin, 2).Value)
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 1).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 3).Value
         li.ListSubItems.Add Text:=Format(Sheets("Inventário").Cells(lin, 4).Value, "0 Meses")
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 5).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 8).Value
    End If
        lin = lin + 1
    Loop
colorir_list
End Sub
Private Sub InventaPrazo2_Click()
Worksheets("Inventário").Select

    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=4, Criteria1:= _
        ">=4", Operator:=xlAnd, Criteria2:="<=6"
        
Sheets("Inventário").Select
    
    TabelaDeInventario.ListItems.Clear
    
    
    lin = 2
    
    Do Until Sheets("Inventário").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
        Set li = TabelaDeInventario.ListItems.Add(Text:=Sheets("Inventário").Cells(lin, 2).Value)
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 1).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 3).Value
         li.ListSubItems.Add Text:=Format(Sheets("Inventário").Cells(lin, 4).Value, "0 Meses")
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 5).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 8).Value
    End If
        lin = lin + 1
    Loop
colorir_list
End Sub

Private Sub InventaPrazo3_Click()
Worksheets("Inventário").Select

    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=4, Criteria1:= _
        ">0", Operator:=xlAnd, Criteria2:="<4"
        
Sheets("Inventário").Select
    
    TabelaDeInventario.ListItems.Clear
    
    
    lin = 2
    
    Do Until Sheets("Inventário").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
        Set li = TabelaDeInventario.ListItems.Add(Text:=Sheets("Inventário").Cells(lin, 2).Value)
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 1).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 3).Value
         li.ListSubItems.Add Text:=Format(Sheets("Inventário").Cells(lin, 4).Value, "0 Meses")
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 5).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 8).Value
    End If
        lin = lin + 1
    Loop
colorir_list
End Sub

Private Sub Classificar2_Click()
On Error GoTo en2
Worksheets("Inventário").Select
    Range("D2").Select
    ActiveWorkbook.Worksheets("Inventário").ListObjects("Tabela2").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Inventário").ListObjects("Tabela2").Sort.SortFields. _
        Add Key:=Range("D2"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Inventário").ListObjects("Tabela2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Inventário").Select
    
    TabelaDeInventario.ListItems.Clear
    
    
    lin = 2
    
    Do Until Sheets("Inventário").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
        Set li = TabelaDeInventario.ListItems.Add(Text:=Sheets("Inventário").Cells(lin, 2).Value)
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 1).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 3).Value
         li.ListSubItems.Add Text:=Format(Sheets("Inventário").Cells(lin, 4).Value, "0 Meses")
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 5).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 8).Value
    End If
        lin = lin + 1
    Loop
en2:
colorir_list
End Sub

Private Sub Classificar1_Click()

On Error GoTo en1

Sheets("Inventário").Select
    Range("D2").Select
    ActiveWorkbook.Worksheets("Inventário").ListObjects("Tabela2").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Inventário").ListObjects("Tabela2").Sort.SortFields. _
        Add Key:=Range("D2"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Inventário").ListObjects("Tabela2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    Sheets("Inventário").Select
    
    TabelaDeInventario.ListItems.Clear
    
    
    lin = 2
    
    Do Until Sheets("Inventário").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
        Set li = TabelaDeInventario.ListItems.Add(Text:=Sheets("Inventário").Cells(lin, 2).Value)
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 1).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 3).Value
         li.ListSubItems.Add Text:=Format(Sheets("Inventário").Cells(lin, 4).Value, "0 Meses")
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 5).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 8).Value
    End If
        lin = lin + 1
    Loop
en1:
colorir_list
End Sub

Private Sub InventaLimpar_Click()
Application.ScreenUpdating = False

Dim w As Worksheet

Set w = Sheets("Inventário")
w.Select


    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=2
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=3
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=4
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=5
    
InventaItem.Value = ""
InventaCamara.Value = ""
InventaLote.Value = ""
InventaPrazo.Value = ""

    Sheets("Inventário").Select
    
    TabelaDeInventario.ListItems.Clear
    
    
    lin = 2
    
    Do Until Sheets("Inventário").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
        Set li = TabelaDeInventario.ListItems.Add(Text:=Sheets("Inventário").Cells(lin, 2).Value)
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 1).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 3).Value
         li.ListSubItems.Add Text:=Format(Sheets("Inventário").Cells(lin, 4).Value, "0 Meses")
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 5).Value
         li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 8).Value
    End If
        lin = lin + 1
    Loop
colorir_list
End Sub

Private Sub InventaPesquisar_Click()
Application.ScreenUpdating = False

Dim w As Worksheet

Set w = Sheets("Inventário")
w.Select

    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=2
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=3
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=4
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=5
    
    
    
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=5, Criteria1:= _
            "=*" & InventaCamara.Value & "*", Operator:=xlAnd
        
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=2, Criteria1:= _
            "=*" & InventaItem.Value & "*", Operator:=xlAnd
    
         
    If InventaLote.Value <> "" Then
        ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=1, Criteria1:= _
            "=" & InventaLote.Value & "", Operator:=xlAnd
    ElseIf InventaLote.Value = "" Then
        GoTo ne
    End If
    
ne:
    ActiveSheet.ListObjects("Tabela2").Range.AutoFilter Field:=4, Criteria1:= _
            "=*" & InventaPrazo.Value & "*", Operator:=xlAnd

        
    Sheets("Inventário").Select
    
    TabelaDeInventario.ListItems.Clear
    
    
    lin = 2
    
    Do Until Sheets("Inventário").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
    Set li = TabelaDeInventario.ListItems.Add(Text:=Sheets("Inventário").Cells(lin, 2).Value)
    li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 1).Value
    li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 3).Value
    li.ListSubItems.Add Text:=Format(Sheets("Inventário").Cells(lin, 4).Value, "0 Meses")
    li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 5).Value
    li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 8).Value
    End If
        lin = lin + 1
    Loop
    colorir_list
    Application.ScreenUpdating = True
End Sub

Private Sub Image4_Click()
Application.ScreenUpdating = False

Dim w   As Worksheet

Set w = Sheets("Estoque")

w.Select
w.Range("b2").Select

If IsNumeric(TextBox1.Value) = True Then
    MsgBox "O nome do produto deve conter pelo menos uma letra!", , "Tente outra vez"
    Application.ScreenUpdating = True
    Exit Sub
End If

If TextBox1.Value = "" Then
    MsgBox "Você esqueceu de digitar o nome!", , "Tente outra vez"
    Application.ScreenUpdating = True
    Exit Sub
End If

Do While ActiveCell.Value <> ""
    If ActiveCell.Value = TextBox1.Value Then
        MsgBox "Este produto já existe!", , "Tente outra vez"
        EstoqueNomeItem.Value = ""
        EstoqueQuantInicial.Value = ""
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    
    ActiveCell.Offset(1, 0).Select
Loop

ActiveCell.Value = TextBox1.Value


TextBox1.Value = ""

atualizalista
atualizacombos
ThisWorkbook.Save

MsgBox "Novo produto cadastrado!", vbOKOnly, ""

atualizacombos

Application.ScreenUpdating = True
End Sub

Private Sub Image5_Click()
TextBox1.Value = ""
End Sub
Private Sub Image2_Click()
Application.ScreenUpdating = False

Dim w As Worksheet

Set w = Sheets("Estoque")
w.Select

ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=1
    

ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=1, Criteria1:= _
        "=*" & ComboBox2.Value & "*", Operator:=xlAnd

Sheets("Estoque").Select

ListView_Estoque.ListItems.Clear


lin = 2

Do Until Sheets("Estoque").Cells(lin, 1) = ""

    If Cells(lin, 1).Rows.Hidden = False Then
    
        Set li = ListView_Estoque.ListItems.Add(Text:=Sheets("Estoque").Cells(lin, 1).Value)
        li.ListSubItems.Add Text:=Sheets("Estoque").Cells(lin, 2).Value
        li.ListSubItems.Add Text:=Sheets("Estoque").Cells(lin, 3).Value
        li.ListSubItems.Add Text:=Sheets("Estoque").Cells(lin, 4).Value
    End If
    lin = lin + 1
Loop

End Sub

Private Sub Image3_Click()
Application.ScreenUpdating = False

Dim w As Worksheet

Set w = Sheets("Estoque")
w.Select

    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=2

ComboBox2.Value = ""


    Sheets("Estoque").Select
    
    ListView_Estoque.ListItems.Clear
    
    
    lin = 2
    
    Do Until Sheets("Estoque").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
        Set li = ListView_Estoque.ListItems.Add(Text:=Sheets("Estoque").Cells(lin, 1).Value)
        li.ListSubItems.Add Text:=Sheets("Estoque").Cells(lin, 2).Value
    End If
        lin = lin + 1
    Loop

Application.ScreenUpdating = True
End Sub

Private Sub CommandButton22_Click()
Application.ScreenUpdating = False

veri = InputBox("Digite a senha para acessar a planilha:", "")

If veri = "" Then
    Sheets("Estoque").Select
    UserForm1.Hide
    Application.Visible = True
    Application.ScreenUpdating = True
    Exit Sub
ElseIf veri <> "" And veri <> "" Then
    MsgBox "Senha incorreta!", , "Tente novamente"
    Application.ScreenUpdating = True
    Exit Sub
Else
    Application.ScreenUpdating = True
    Exit Sub
End If

Application.ScreenUpdating = True
End Sub

Private Sub CommandButton26_Click()
Application.ScreenUpdating = False

veri = InputBox("Digite a senha para acessar a planilha:", "")

If veri = "" Then
    Sheets("Movimentação").Select
    UserForm1.Hide
    Application.Visible = True
    Application.ScreenUpdating = True
    Exit Sub
ElseIf veri <> "" And veri <> "" Then
    MsgBox "Senha incorreta!", , "Tente novamente"
    Application.ScreenUpdating = True
    Exit Sub
Else
    Application.ScreenUpdating = True
    Exit Sub
End If

Application.ScreenUpdating = True
End Sub

Private Sub CommandButton25_Click()
Application.ScreenUpdating = False

Dim id  As Long
Dim w   As Worksheet
On Error GoTo erro

id = InputBox("Digite o ID do item: ", "")

Set w = Sheets("Movimentação")
w.Select
w.Range("F2").Select

Do While ActiveCell.Value <> ""
    If ActiveCell.Value = id Then
        ComboBox1.Value = ActiveCell.Offset(0, -4).Value & " | " & ActiveCell.Offset(0, -5).Value & " | " & ActiveCell.Offset(0, -3).Value & " Unid." & " | " & ActiveCell.Offset(0, -2).Value & " | " & ActiveCell.Value
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ActiveCell.Offset(1, 0).Select
Loop

MsgBox "ID não encontrado, tente novamente!", , ""

erro:

Application.ScreenUpdating = True
End Sub

Private Sub MovimentaLimpar_Click()
Application.ScreenUpdating = False

Dim w As Worksheet

Set w = Sheets("Movimentação")
w.Select

    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=2
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=7
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=4
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=8
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=9

MovimentaItem.Value = ""
MovimentaMovi.Value = ""
MovimentaQuant.Value = ""
MovimentaData.Value = ""
combocamara2.Value = ""

    Sheets("Movimentação").Select
    
    TabelaDeMovimentações.ListItems.Clear
    
    
    lin = 2
    
    Do Until Sheets("Movimentação").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
        Set li = TabelaDeMovimentações.ListItems.Add(Text:=Sheets("Movimentação").Cells(lin, 1).Value)
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 2).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 3).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 7).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 8).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 9).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 4).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 6).Value
    End If
        lin = lin + 1
    Loop

Application.ScreenUpdating = True
colorir_list
End Sub

Private Sub MovimentaPerquisar_Click()
Application.ScreenUpdating = False

Dim w As Worksheet

Set w = Sheets("Movimentação")
w.Select

    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=2
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=7
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=4
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=8
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=9
    
    
    
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=1, Criteria1:= _
            "=*" & MovimentaItem.Value & "*", Operator:=xlAnd
        
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=2, Criteria1:= _
            "=*" & MovimentaMovi.Value & "*", Operator:=xlAnd
         
    If MovimentaQuant.Value <> "" Then
        ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=7, Criteria1:= _
            "=" & MovimentaQuant.Value & "", Operator:=xlAnd
    ElseIf MovimentaQuant.Value = "" Then
        GoTo ne
    End If
    
ne:
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=4, Criteria1:= _
            "=*" & MovimentaData.Value & "*", Operator:=xlAnd
            
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=9, Criteria1:= _
            "=*" & combocamara2.Value & "*", Operator:=xlAnd
        
    Sheets("Movimentação").Select
    
    TabelaDeMovimentações.ListItems.Clear
    
    
    lin = 2
    
    Do Until Sheets("Movimentação").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
        Set li = TabelaDeMovimentações.ListItems.Add(Text:=Sheets("Movimentação").Cells(lin, 1).Value)
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 2).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 3).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 7).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 8).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 9).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 4).Value
        li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 6).Value
    End If
        lin = lin + 1
    Loop
        
    Application.ScreenUpdating = True
    colorir_list
    Exit Sub


End Sub

Private Sub EstoqueAlterar_Click()
Application.ScreenUpdating = False

Dim w   As Worksheet
Dim w2  As Worksheet

Set w = Sheets("Movimentação")
Set w2 = Sheets("Estoque")

w2.Select
w2.Range("a2").Select

If EstoqueBuscarItem.Value = "" Then
   MsgBox "Este item não existe para ser alterado!"
   Application.ScreenUpdating = True
   Exit Sub
End If


Do While ActiveCell.Value <> ""
    If ActiveCell.Value = EstoqueAlterarItem.Value Then
        MsgBox "Já existe um item no inventário com esse nome!"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ActiveCell.Offset(1, 0).Select
Loop

w2.Range("a2").Select

Do While ActiveCell.Value <> EstoqueBuscarItem.Value
    ActiveCell.Offset(1, 0).Select
Loop

ActiveCell.Value = EstoqueAlterarItem.Value

w.Select
w.Range("a2").Select

Do While ActiveCell.Value <> ""
    If ActiveCell.Value = EstoqueBuscarItem.Value Then
        ActiveCell.Value = EstoqueAlterarItem.Value
    End If
    ActiveCell.Offset(1, 0).Select
Loop

atualizacombos
atualizalista
ThisWorkbook.Save
MsgBox "Item alterado com sucesso!"

Application.ScreenUpdating = True
End Sub

Private Sub EstoqueDeletar_Click()
Application.ScreenUpdating = False

Dim w   As Worksheet

Set w = Sheets("Estoque")
w.Select
w.Range("A2").Select

If EstoqueBuscarItem.Value = "" Then
   MsgBox "Este item não existe para ser deletado!"
   Application.ScreenUpdating = True
   Exit Sub
End If

If ActiveCell.Value = EstoqueBuscarItem.Value And ActiveCell.Offset(1, 0).Value = "" Then
    ActiveCell.Value = ""
    atualizacombos
    atualizalista
    MsgBox "Produto apagado com sucesso!"
    EstoqueBuscarItem.Value = ""
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Exit Sub
ElseIf ActiveCell.Value = EstoqueBuscarItem.Value And ActiveCell.Offset(1, 0).Value <> "" Then
    ActiveCell.EntireRow.Delete
    atualizacombos
    EstoqueBuscarItem.Value = ""
    MsgBox "Item apagado com sucesso!"
    atualizalista
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Exit Sub
Else
    Do While ActiveCell.Value <> EstoqueBuscarItem.Value
        ActiveCell.Offset(1, 0).Select
    Loop
    ActiveCell.EntireRow.Delete
    atualizacombos
    EstoqueBuscarItem.Value = ""
    MsgBox "Item apagado com sucesso!"
    atualizalista
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Exit Sub
End If
End Sub

Private Sub EstoqueLimpar2_Click()
    EstoqueBuscarItem.Value = ""
End Sub

Private Sub EstoqueLimpar1_Click()
    EstoqueNomeItem.Value = ""
    EstoqueQuantInicial.Value = ""
End Sub

Private Sub EstoqueAdicionar_Click()
Application.ScreenUpdating = False

Dim w   As Worksheet

Set w = Sheets("Estoque")

w.Select
w.Range("A2").Select

If IsNumeric(EstoqueNomeItem.Value) = True Then
    MsgBox "O nome do produto deve conter pelo menos uma letra!", , "Tente outra vez"
    Application.ScreenUpdating = True
    Exit Sub
End If

If EstoqueNomeItem.Value = "" Then
    MsgBox "Você esqueceu de digitar o nome!", , "Tente outra vez"
    Application.ScreenUpdating = True
    Exit Sub
End If

Do While ActiveCell.Value <> ""
    If ActiveCell.Value = EstoqueNomeItem.Value Then
        MsgBox "Este produto já existe no inventário!", , "Tente outra vez"
        EstoqueNomeItem.Value = ""
        EstoqueQuantInicial.Value = ""
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    
    ActiveCell.Offset(1, 0).Select
Loop

ActiveCell.Value = EstoqueNomeItem.Value


EstoqueNomeItem.Value = ""

atualizalista
atualizacombos
ThisWorkbook.Save

MsgBox "Novo produto cadastrado!", vbOKOnly, ""

atualizacombos

Application.ScreenUpdating = True
End Sub

Private Sub btlimpar_Click()

limpar

End Sub

Private Sub btsalvar_Click() '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Application.ScreenUpdating = False


Dim w As Worksheet
Dim w2 As Worksheet

If txtquant.Value = 0 Then

    MsgBox "A quantidade não pode ser zero", vbOKOnly, "Atenção!"
    Application.ScreenUpdating = True
    Exit Sub
End If

'****************************Verifica a se há itens o suficiente em estoque**********************************

Dim g   As Worksheet

Set g = Sheets("Movimentação")
g.Select
g.Range("a2").Select



en = 0
sa = 0

Do While ActiveCell.Value <> ""
    If comboitem.Value = ActiveCell.Value And combocamara = ActiveCell.Offset(0, 8).Value And (Trim(Str(ActiveCell.Offset(0, 6).Value)) = Trim(Str(txtlote.Value))) = True Then
            If ActiveCell.Offset(0, 1).Value = "Entrada" Then
                en = en + ActiveCell.Offset(0, 2).Value
            ElseIf ActiveCell.Offset(0, 1).Value = "Saída" Then
                sa = sa + ActiveCell.Offset(0, 2).Value
            End If
    Else
    End If
    
    ActiveCell.Offset(1, 0).Select
Loop
sa = sa + txtquant.Value
If combomovimenta = "Saída" Then
    If en - sa < 0 Then
        MsgBox "Não há itens o suficiente para essa saída!", , ""
        Application.ScreenUpdating = True
        Exit Sub
    End If
End If
'*************************************************************************************************************


Set w = Sheets("Movimentação")

If comboitem.Value = "" Or combomovimenta.Value = "" Or txtquant.Value = "" Or txtvali.Value = "" Or txtlote.Value = "" Or combocamara.Value = "" Then

    MsgBox "Digite todos os dados!", vbOKOnly, "Atenção!"
    Application.ScreenUpdating = True
    Exit Sub
Else

    w.Select

    If Range("A2") = "" Then
    
        w.Range("a2").Select
        
        ActiveCell.Value = comboitem.Value
        ActiveCell.Offset(0, 1).Value = combomovimenta.Value
        ActiveCell.Offset(0, 2).Value = txtquant.Value
        If txtdata.Value = "" Then
            ActiveCell.Offset(0, 3).Value = Date & " - " & Time
        Else
            If Len(txtdata.Value) = 8 Then
                ActiveCell.Offset(0, 3).Value = Format(txtdata.Value, "00/00/0000") & " - " & Time
            Else
                 ActiveCell.Offset(0, 3).Value = txtdata.Value & " - " & Time
            End If
        End If
        
        ActiveCell.Offset(0, 4).Value = txtobs.Value
        ActiveCell.Offset(0, 5).Value = 1
        ActiveCell.Offset(0, 6).Value = txtlote.Value
        If Len(txtvali.Value) = 8 Then
                ActiveCell.Offset(0, 7).Value = Format(txtvali.Value, "00/00/0000")
            Else
                 ActiveCell.Offset(0, 7).Value = txtdata.Value
        End If
        ActiveCell.Offset(0, 8).Value = combocamara.Value
    Else
        w.Range("A1048576").Select
        ActiveCell.End(xlUp).Offset(1, 0).Select
        
        ActiveCell.Value = comboitem.Value
        ActiveCell.Offset(0, 1).Value = combomovimenta.Value
        ActiveCell.Offset(0, 2).Value = txtquant.Value
        If Len(txtdata.Value) > 10 Then
            
           ActiveCell.Offset(0, 3).Value = Mid(txtdata.Value, 1, 10) & " - " & Time
           
        ElseIf txtdata.Value = "" Then
            ActiveCell.Offset(0, 3).Value = Date & " - " & Time
        Else
            If Len(txtdata.Value) = 8 Then
                ActiveCell.Offset(0, 3).Value = Format(txtdata.Value, "00/00/0000") & " - " & Time
            Else
                 ActiveCell.Offset(0, 3).Value = txtdata.Value & " - " & Time
            End If
        End If
        
        ActiveCell.Offset(0, 4).Value = txtobs.Value
        ActiveCell.Offset(0, 5).Value = ActiveCell.Offset(-1, 5).Value + 1
        ActiveCell.Offset(0, 6).Value = txtlote.Value
           
        If Len(txtvali.Value) = 8 Then
                ActiveCell.Offset(0, 7).Value = Format(txtvali.Value, "00/00/0000")
        Else
                ActiveCell.Offset(0, 7).Value = txtvali.Value
        End If
        ActiveCell.Offset(0, 8).Value = combocamara.Value
        
    End If
    
    MsgBox "Processo concluído!", vbOKOnly, "Tudo certo"
End If
'==========================================================================================================
Set w2 = Sheets("Inventário")
w2.Select
w2.Range("a2").Select

Do While ActiveCell.Value <> ""
    If (Trim(Str(ActiveCell.Value)) = Trim(Str(txtlote.Value))) = True And ActiveCell.Offset(0, 1).Value = comboitem.Value And ActiveCell.Offset(0, 4).Value = combocamara.Value Then
    GoTo est
    End If
    ActiveCell.Offset(1, 0).Select
Loop
ActiveCell.Value = txtlote.Value
ActiveCell.Offset(0, 1).Value = comboitem.Value
ActiveCell.Offset(0, 2).Value = txtvali.Value
ActiveCell.Offset(0, 4).Value = combocamara.Value

'==========================================================================================================
est:
w2.Range("a2").Select

Do While ActiveCell.Value <> ""
If ActiveCell.Offset(0, 7).Value = 0 Then
    ActiveCell.EntireRow.Delete
End If
ActiveCell.Offset(1, 0).Select
Loop

limpar
    
atualizacombos

atualizalista

ThisWorkbook.Save

Application.ScreenUpdating = True

End Sub

'Botão Deletar
Private Sub CommandButton12_Click() '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Application.ScreenUpdating = False


Dim ident   As String
Dim w2      As Worksheet

If ComboBox1.Value = "" Then

    MsgBox "Sem dados para deletar!", vbOKOnly, "Atenção!"
    Application.ScreenUpdating = True
    Exit Sub
    
End If

ident = Mid(ComboBox1.Value, InStrRev(ComboBox1.Value, "|") + 2, 30)


Set w2 = Sheets("Movimentação")
w2.Select
w2.Range("A1028576").End(xlUp).Select


Do While ActiveCell.Value <> "Item"

    If ActiveCell.Offset(0, 5).Value = ident Then
    
        ActiveCell.EntireRow.Delete
    
    End If
    
    ActiveCell.Offset(-1, 0).Select
    
Loop

MsgBox "Registro apagado!", , ""
 
limpar

atualizacombos

atualizalista


ThisWorkbook.Save

Application.ScreenUpdating = True

End Sub

'Botão Alterar

Private Sub CommandButton16_Click() '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Application.ScreenUpdating = False

If ComboBox1.Value = "" Then

    MsgBox "Sem dados para alterar!", vbOKOnly, "Atenção!"
    Application.ScreenUpdating = True
    Exit Sub
    
End If

If txtquant.Value = 0 Then
    MsgBox "Digite uma quantidade válida", vbOKOnly, "Atenção!"
    txtquant.Value = ""
    Application.ScreenUpdating = True
    Exit Sub
End If

'****************************Verifica a se há itens o suficiente em estoque**********************************

Dim g   As Worksheet


Set g = Sheets("Estoque")
g.Select
g.Range("a2").Select

Do While ActiveCell.Value <> comboitem.Value

    ActiveCell.Offset(1, 0).Select
Loop

If Mid(ComboBox1.Value, 1, InStr(ComboBox1.Value, "|") - 2) = "Saída" Then

    If combomovimenta.Value = "Saída" Then
    
        If ActiveCell.Offset(0, 1).Value - ((ActiveCell.Offset(0, 2).Value - (Mid(ComboBox1, InStr(Mid(ComboBox1, 8, 30), "|") + 8, InStr(ComboBox1, "U") - 24))) + txtquant.Value) < 0 Then
            MsgBox "Você não tem " & comboitem.Value & " o suficiente para essa saída!", vbOKOnly, "Atenção"
            txtquant.Value = ""
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
    ElseIf combomovimenta.Value = "Entrada" Then
    
        If (ActiveCell.Offset(0, 1).Value + txtquant.Value) - (ActiveCell.Offset(0, 2).Value - txtquant.Value) < 0 Then
            MsgBox "Você não tem " & comboitem.Value & " o suficiente para essa saída!", vbOKOnly, "Atenção"
            txtquant.Value = ""
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
    End If
    
    
ElseIf Mid(ComboBox1.Value, 1, InStr(ComboBox1.Value, "|") - 2) = "Entrada" Then

    If combomovimenta.Value = "Saída" Then
    
        If (ActiveCell.Offset(0, 1).Value - Mid(ComboBox1, InStr(Mid(ComboBox1, 10, 30), "|") + 10, InStr(ComboBox1, "U") - 26)) - (ActiveCell.Offset(0, 2).Value + txtquant.Value) < 0 Then

            MsgBox "Você não tem " & comboitem.Value & " o suficiente para essa saída!", vbOKOnly, "Atenção"
            txtquant.Value = ""
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
        
    End If

End If

'*************************************************************************************************************

Dim ident   As String
Dim w2      As Worksheet

ident = Mid(ComboBox1.Value, InStrRev(ComboBox1.Value, "|") + 2, 30)

Set w2 = Sheets("Movimentação")
w2.Select
w2.Range("A1028576").End(xlUp).Select


Do While ActiveCell.Value <> "Item"

    If ActiveCell.Offset(0, 5).Value = ident Then
    
        ActiveCell.Value = comboitem.Value
        ActiveCell.Offset(0, 1).Value = combomovimenta.Value
        ActiveCell.Offset(0, 2).Value = txtquant.Value
        If Len(txtdata.Value) > 10 Then
            
           ActiveCell.Offset(0, 3).Value = Mid(ActiveCell.Offset(0, 3).Value, 1, 10) & " - " & Time
            
        Else
            ActiveCell.Offset(0, 3).Value = txtdata.Value & " - " & Time
        
        End If
        
        ActiveCell.Offset(0, 4).Value = txtobs.Value
        
    End If
    
    ActiveCell.Offset(-1, 0).Select
    
Loop
Sheets("Estoque").Select

limpar
    
atualizacombos

atualizalista

MsgBox "Registro alterado com sucesso", vbOKOnly, "Processo concluído!"


Application.ScreenUpdating = True

End Sub
'============================================================================================================
'============================================================================================================
'============================================================================================================

Private Sub EstoqueBuscarItem_Change()
 EstoqueAlterarItem.Value = EstoqueBuscarItem.Value
End Sub

Private Sub ComboBox1_Change() '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 
Application.ScreenUpdating = False

Dim ident   As String
Dim w2      As Worksheet

On Error GoTo erro

ident = Mid(ComboBox1.Value, InStrRev(ComboBox1.Value, "|") + 2, 30)



Set w2 = Sheets("Movimentação")
w2.Select
w2.Range("A1028576").End(xlUp).Select


Do While ActiveCell.Value <> "Item"

    If ActiveCell.Offset(0, 5).Value = ident Then
    
        comboitem.Value = ActiveCell.Value
        combomovimenta.Value = Mid(ActiveCell.Offset(0, 1), 1, 10)
        txtquant.Value = ActiveCell.Offset(0, 2).Value
        txtdata.Value = ActiveCell.Offset(0, 3).Value
        txtobs.Value = ActiveCell.Offset(0, 4).Value
        txtvali.Value = ActiveCell.Offset(0, 7).Value
        txtlote.Value = ActiveCell.Offset(0, 6).Value
        combocamara.Value = ActiveCell.Offset(0, 8).Value
    
    
    End If
    
    ActiveCell.Offset(-1, 0).Select
    
Loop

erro:

Sheets("Estoque").Select

Application.ScreenUpdating = True

End Sub

'============================================================================================================
'==========================================RESTRIÇÕES DE ENTRADA=============================================
'============================================================================================================

Private Sub txtdata_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    txtdata = Format(txtdata, "00/00/0000")
End Sub

Private Sub txtdata_Change()
    txtdata.MaxLength = 8
End Sub

Private Sub txtdata_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtquant_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub comboitem_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
     If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub combomovimenta_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub ComboBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub EstoqueQuantInicial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub


Private Sub MovimentaMovi_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub


Private Sub MovimentaData_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub MovimentaData_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    MovimentaData = Format(MovimentaData, "00/00/0000")
End Sub

Private Sub MovimentaData_Change()
    MovimentaData.MaxLength = 8
End Sub

Private Sub EstoqueBuscarItem_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub
Private Sub combocamara_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub txtvali_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
        txtvali = Format(txtvali, "00/00/0000")
End Sub
Private Sub txtvali_Change()
    txtvali.MaxLength = 8
End Sub
Private Sub txtvali_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub combocamara2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub
Private Sub InventaCamara_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub


'============================================================================================================
'============================================================================================================
'============================================================================================================
'=========================================SUBS REUTILIZÁVEIS=================================================
'============================================================================================================
'============================================================================================================
'============================================================================================================

Sub atualizacombos()

'Atualiza combobox

Application.ScreenUpdating = False

comboitem.Clear
ComboBox1.Clear
EstoqueBuscarItem.Clear
MovimentaItem.Clear
MovimentaMovi.Clear
combomovimenta.Clear
ComboBox2.Clear
combocamara.Clear
combocamara2.Clear
InventaItem.Clear
InventaCamara.Clear

MovimentaMovi.AddItem "Entrada"
MovimentaMovi.AddItem "Saída"
combomovimenta.AddItem "Entrada"
combomovimenta.AddItem "Saída"

Dim w       As Worksheet
Dim w2      As Worksheet
Dim a       As Integer

Set w = Sheets("Estoque")
w.Select
w.Range("A2").Select

Do While ActiveCell.Value <> ""

    comboitem.AddItem ActiveCell.Value
    EstoqueBuscarItem.AddItem ActiveCell.Value
    MovimentaItem.AddItem ActiveCell.Value
    ComboBox2.AddItem ActiveCell.Value
    InventaItem.AddItem ActiveCell.Value
    If ActiveCell.Offset(0, 1).Value <> "" Then
        combocamara.AddItem ActiveCell.Offset(0, 1).Value
        combocamara2.AddItem ActiveCell.Offset(0, 1).Value
        InventaCamara.AddItem ActiveCell.Offset(0, 1).Value
    End If
    
    ActiveCell.Offset(1, 0).Select
    
Loop

Set w2 = Sheets("Movimentação")
w2.Select
w2.Range("A1028576").End(xlUp).Select

Do While ActiveCell.Value <> "Item"
    
    If a > 19 Then
        Exit Do
        
    Else
        ComboBox1.AddItem ActiveCell.Offset(0, 1).Value & " | " & Mid(ActiveCell.Value, 1, InStr(ActiveCell.Value, " ")) & " | " & ActiveCell.Offset(0, 2).Value & " Unid." & " | " & ActiveCell.Offset(0, 3).Value & " | " & ActiveCell.Offset(0, 5)
        a = a + 1
        ActiveCell.Offset(-1, 0).Select
    End If
Loop

Application.ScreenUpdating = True


End Sub
Sub limpar() '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Application.ScreenUpdating = False

    comboitem.Value = ""
    combomovimenta.Value = ""
    txtquant.Value = ""
    txtdata.Value = ""
    txtobs.Value = ""
    txtvali.Value = ""
    txtlote.Value = ""
    combocamara.Value = ""
    
Application.ScreenUpdating = True
End Sub

Sub atualizalista()
Application.ScreenUpdating = False

Sheets("Movimentação").Select

TabelaDeMovimentações.ListItems.Clear


lin = 2

Do Until Sheets("Movimentação").Cells(lin, 1) = ""

    Set li = TabelaDeMovimentações.ListItems.Add(Text:=Sheets("Movimentação").Cells(lin, 1).Value)
    li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 2).Value
    li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 3).Value
    li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 7).Value
    li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 8).Value
    li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 9).Value
    li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 4).Value
    li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 6).Value

    lin = lin + 1
Loop

Sheets("Estoque").Select

ListView_Estoque.ListItems.Clear


lin = 2

Do Until Sheets("Estoque").Cells(lin, 1) = ""

    Set li = ListView_Estoque.ListItems.Add(Text:=Sheets("Estoque").Cells(lin, 1).Value)
    li.ListSubItems.Add Text:=Sheets("Estoque").Cells(lin, 2).Value
    li.ListSubItems.Add Text:=Sheets("Estoque").Cells(lin, 3).Value
    li.ListSubItems.Add Text:=Sheets("Estoque").Cells(lin, 4).Value
    lin = lin + 1
Loop

Sheets("Inventário").Select

TabelaDeInventario.ListItems.Clear


lin = 2

Do Until Sheets("Inventário").Cells(lin, 1) = ""

    Set li = TabelaDeInventario.ListItems.Add(Text:=Sheets("Inventário").Cells(lin, 2).Value)
    li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 1).Value
    li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 3).Value
    li.ListSubItems.Add Text:=Format(Sheets("Inventário").Cells(lin, 4).Value, "0 Meses")
    li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 5).Value
    li.ListSubItems.Add Text:=Sheets("Inventário").Cells(lin, 8).Value
    lin = lin + 1
Loop
colorir_list
Application.ScreenUpdating = True
End Sub

Sub limpar2()
Application.ScreenUpdating = False

Dim w As Worksheet

Set w = Sheets("Movimentação")
w.Select

    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=2
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=3
    ActiveSheet.ListObjects("Tabela3").Range.AutoFilter Field:=4


Dim w2 As Worksheet

Set w2 = Sheets("Estoque")
w2.Select

    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=1
colorir_list
Application.ScreenUpdating = True
End Sub

Sub colorir_list()
Dim i As Long


For i = 1 To TabelaDeInventario.ListItems.Count
    If TabelaDeInventario.ListItems.Item(i).ListSubItems(3).Text = "3 Meses" Or TabelaDeInventario.ListItems.Item(i).ListSubItems(3).Text = "0 Meses" Or TabelaDeInventario.ListItems.Item(i).ListSubItems(3).Text = "1 Meses" Or TabelaDeInventario.ListItems.Item(i).ListSubItems(3).Text = "2 Meses" Then
        TabelaDeInventario.ListItems.Item(i).ListSubItems(3).ForeColor = RGB(242, 184, 0)
        
    ElseIf TabelaDeInventario.ListItems.Item(i).ListSubItems(3).Text = "4 Meses" Or TabelaDeInventario.ListItems.Item(i).ListSubItems(3).Text = "5 Meses" Or TabelaDeInventario.ListItems.Item(i).ListSubItems(3).Text = "6 Meses" Then
        TabelaDeInventario.ListItems.Item(i).ListSubItems(3).ForeColor = RGB(0, 176, 80)
        
        
    ElseIf TabelaDeInventario.ListItems.Item(i).ListSubItems(3).Text = "Vencido" Then
        TabelaDeInventario.ListItems.Item(i).ListSubItems(3).ForeColor = RGB(225, 0, 0)
        
    Else
        TabelaDeInventario.ListItems.Item(i).ListSubItems(3).ForeColor = RGB(0, 112, 192)
    End If
    
    
Next

End Sub

