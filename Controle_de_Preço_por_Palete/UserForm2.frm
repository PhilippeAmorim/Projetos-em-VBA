VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13380
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Image2_Click()
Application.ScreenUpdating = False

veri = InputBox("Digite a senha para acessar a planilha:", "")

If veri = "" Then
    Sheets("Controle").Select
    Range("A1:XFD1048576").Select
    Selection.EntireRow.Hidden = False
    Selection.EntireColumn.Hidden = False
    Range("A2").Select
    UserForm2.Hide
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


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ThisWorkbook.Save
    ThisWorkbook.Close
    Application.Quit
End Sub


Private Sub UserForm_Initialize()

Label48.Visible = False
Label49.Visible = False
TextBox3.Visible = False
Label1.Caption = ""
Frame3.Visible = False
Valor.Visible = False
CommandButton1.Visible = False
Bt_Alterar.Visible = False



atualizar_combos
End Sub
'===================================================================================================
'===================================================================================================
'===================================================================================================
Private Sub MultiPage1_Change()
Range("A1:XFD1048576").Select
Selection.EntireRow.Hidden = False
Selection.EntireColumn.Hidden = False

End Sub

Private Sub Análise_Envio_Change()
Application.ScreenUpdating = False

Dim linha  As Long

ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear

Range("A1:XFD1048576").Select
Selection.EntireRow.Hidden = False
Selection.EntireColumn.Hidden = False


Range("A2").Select

Do While ActiveCell.Value <> ""
If Trim(Mid(ActiveCell.Value, 1, 8)) = Análise_Envio.Value Then
    Exit Do
End If

ActiveCell.Offset(1, 0).Select
Loop

ActiveCell.Offset(1, 0).Select

Va = ActiveCell.Row
B = 0
A = 0
C = 0

ActiveCell.Offset(1, 0).Select

Do While Mid(ActiveCell.Value, 1, 2) <> "En" And ActiveCell.Value <> ""
B = B + 1
ActiveCell.Offset(1, 0).Select
Loop

Range("A" & Va).Select

Análise_Preço.Caption = Mid(ActiveCell.Offset(-1, 0).Value, InStr(ActiveCell.Offset(-1, 0).Value, "R$"), 30)

ListView1.Refresh

'===================Cria o cabeçalho da listview=============================

With ListView1
    .BorderStyle = ccFixedSingle
    .Gridlines = True
    .View = lvwReport
    .ColumnHeaders.Add Text:=" ", Width:=155
    ActiveCell.Offset(0, 1).Select
        Do While A <= 22
        
        If ActiveCell.Value <> "" Then
            .ColumnHeaders.Add Text:=Combo_Produto.Column(0, ActiveCell.Column - 1), Width:=180, Alignment:=2
        End If
            ActiveCell.Offset(0, 1).Select
            A = A + 1
        Loop
End With

'=============================================================================
'=====================Coloca os dados na listview=============================
Range("A" & ActiveCell.Offset(-2, 0).Row & ":A1").Select
Selection.EntireRow.Hidden = True
Range("A" & Va).Select
Range("A" & ActiveCell.Offset(B + 1, 0).Row & ":A1048576").Select
Selection.EntireRow.Hidden = True
Range("A" & Va).Select
A = 0

Do While A <= 22
    If ActiveCell.Value = "" Then
        ActiveCell.EntireColumn.Hidden = True
    End If
ActiveCell.Offset(0, 1).Select
A = A + 1
Loop



Range("A" & Va).Select

Sheets("Controle").Select

ListView1.ListItems.Clear


lin = 2

Do Until Sheets("Controle").Cells(lin, 1) = ""

    If Cells(lin, 1).Rows.Hidden = False Then
    
        Set li = ListView1.ListItems.Add(Text:=Sheets("Controle").Cells(lin, 1).Value)
        
        A = 1
        cont = 1
        Do While A <= 22
            If ActiveCell.Offset(0, A).Value <> "" Then
                cont = cont + 1
                li.ListSubItems.Add Text:=Sheets("Controle").Cells(lin, A + 1).Value
            End If
        A = A + 1
        Loop
        li.ListSubItems.Add Text:=Sheets("Controle").Cells(lin, 1).Value

    End If
    lin = lin + 1
Loop

'=============================================================================
ListView1.ListItems.Remove (1)

colorir_list
Application.ScreenUpdating = True
End Sub


Private Sub Combo_INVOICE_Change()
Application.ScreenUpdating = False

Range("A2").Select

If Combo_Produto.Value = "" Then
  Text_Quant.Value = ""
  GoTo pro
End If



Do While Mid(ActiveCell.Value, 1, 8) <> Mid(Label1.Caption, 1, 8)
ActiveCell.Offset(1, 0).Select
Loop
ActiveCell.Offset(1, 0).Select


Text_Quant.Value = Range("a" & Valor).Offset(Combo_INVOICE.ListIndex + 1, Combo_Produto.ListIndex + 1).Value

pro:

Range("A" & Valor).Select

Application.ScreenUpdating = True

End Sub

Private Sub Combo_Produto_Change()
Application.ScreenUpdating = False

Range("A2").Select

Do While Mid(ActiveCell.Value, 1, 7) <> Mid(Label1.Caption, 1, 7)
ActiveCell.Offset(1, 0).Select
Loop

ActiveCell.Offset(1, 0).Select

 
TextBox2.Value = Mid(Range("a" & Valor).Offset(Combo_INVOICE.ListIndex + 1, Combo_Produto.ListIndex + 1).Value, InStr(Range("a" & Valor).Offset(Combo_INVOICE.ListIndex + 1, Combo_Produto.ListIndex + 1).Value, "/") + 1, 30)

Range("A" & Valor).Select

Application.ScreenUpdating = True
End Sub


Private Sub Comboultimo_Change()
Application.ScreenUpdating = False

Combo_INVOICE.Clear

Label1.Visible = True

If Comboultimo.Value <> "" Then
    CommandButton1.Visible = False
    Bt_Alterar.Visible = True
    Text_in1.Visible = False
    Text_in2.Visible = False
    Label10.Visible = False
    Label8.Visible = False
Else
    CommandButton1.Visible = True
End If


Range("A2").Select

Do While ActiveCell.Value <> ""
If Trim(Mid(ActiveCell.Value, 1, 8)) = Comboultimo.Value Then
    Exit Do
End If
ActiveCell.Offset(1, 0).Select
Loop

Label1.Caption = Comboultimo.Value & " - " & Mid(ActiveCell.Value, InStr(ActiveCell.Value, "/") + 1, 4)
If ActiveCell.Value <> "" Then
    Text_Preco.Value = Mid(ActiveCell, InStr(ActiveCell.Value, "R$"), 30)
End If
Valor.Value = ActiveCell.Offset(1, 0).Row
Frame3.Visible = True

Range("A" & Valor).Select


Application.ScreenUpdating = True
End Sub

'===================================================================================================
'=============================================Botões================================================
'===================================================================================================
Private Sub Bt_Deletar_Click()
Application.ScreenUpdating = False

If Comboultimo.Value = "" Then
    MsgBox "Selecione um envio para apagar.", , ""
    Application.ScreenUpdating = True
    Exit Sub
End If

With ActiveCell
    For k = 1 To Combo_INVOICE.ListCount
    .Offset(1, 0).EntireRow.Delete
    Next
    .Offset(-1, 0).EntireRow.Delete
    .EntireRow.Delete
End With

MsgBox Comboultimo & " deletado com sucesso!", , ""

limpar_tudo

atualizar_combos

Application.ScreenUpdating = True
End Sub

Private Sub Bt_Alterar_Click()
If Text_Preco.Value = "" Then
    MsgBox "Coloque um preço!", , ""
    Exit Sub
End If
ActiveCell.Offset(-1, 0).Value = Mid(ActiveCell.Offset(-1, 0).Value, 1, InStr(ActiveCell.Offset(-1, 0).Value, ":") + 1) & Format(Text_Preco, "R$ #.00")


Label49.Visible = True
TempoEspera = Now() + TimeValue("00:00:01")
While Now() < TempoEspera
    DoEvents
Wend
Label49.Visible = False

atualizar_combos


End Sub

Private Sub Bt_Salvar_Click()
If Frame3.Visible = False Then
    salvar
End If

If Text_Preco.Value <> "" Or Text_in1.Value <> "" Or Text_in2.Value <> "" Then
    MsgBox "Envio Salvo!", , ""
End If
limpar_tudo
End Sub

Private Sub Bt_Limpar_Click()
limpar_tudo
End Sub


Private Sub CommandButton2_Click()
If Combo_INVOICE <> "" And Text_Quant <> "" And TextBox2 <> "" And Combo_Produto <> "" Then
ActiveCell.Offset(Combo_INVOICE.ListIndex + 1, Combo_Produto.ListIndex + 1).Value = Text_Quant.Value
Else
MsgBox "Preecha todos os dados!", , ""
End If


TextBox3.Visible = True
TempoEspera = Now() + TimeValue("00:00:01")
While Now() < TempoEspera
    DoEvents
Wend
TextBox3.Visible = False

End Sub

Private Sub CommandButton3_Click()
With ActiveCell.Offset(0, Combo_Produto.ListIndex + 1)
    If TextBox2 <> "" And Combo_Produto <> "" Then
        .FormulaR1C1 = "=SOMAESP(R[1]C:R[" & Combo_INVOICE.ListCount & "]C," & TextBox2 & ")"
        .Font.Size = 12
        .Font.Bold = True
    Else
        MsgBox "Preecha todos os dados!", , ""
    End If
End With


Label48.Visible = True
TempoEspera = Now() + TimeValue("00:00:01")
While Now() < TempoEspera
    DoEvents
Wend
Label48.Visible = False


End Sub

Private Sub CommandButton1_Click()

salvar

CommandButton1.Visible = False

End Sub

'===================================================================================================
'============================================RESTRIÇÕES=============================================
'===================================================================================================
Private Sub Text_Preco_Enter()
If Comboultimo.Value <> "" Then
    Text_Preco.Value = ""
End If
End Sub

Private Sub Text_Preco_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

Text_Preco.Value = Format(Text_Preco, "R$ #.00")

Application.ScreenUpdating = False
Dim w As Worksheet

Label1.Visible = True

Set w = Sheets("Controle")
w.Select
w.Range("A2").Select

cont = 0

If Range("A2").Value = "" Then
Label1.Caption = "Envio 1" & " - " & Mid(Date, 7, 4)
ElseIf Range("A2").Value <> "" And Comboultimo.Value = "" Then
    Do While ActiveCell.Value <> ""
    If Mid(ActiveCell.Value, 1, 2) = "En" Then
        cont = cont + 1
    End If
    ActiveCell.Offset(1, 0).Select
    Loop
    
    Do While ActiveCell <> Range("A1")
        If Mid(ActiveCell.Value, 9, 4) <> Mid(Date, 7, 4) Then
            Label1.Caption = "Envio 1" & " - " & Mid(Date, 7, 4)
            Application.ScreenUpdating = True
            CommandButton1.Visible = True
            Exit Sub
        End If
   CommandButton1.Visible = True
    ActiveCell.Offset(-1, 0).Select
    Loop
Label1.Caption = "Envio " & cont + 1 & " - " & Mid(Date, 7, 4)
End If
If Comboultimo.Value = "" Then
    CommandButton1.Visible = True
End If

If Comboultimo.Value <> "" Then
    Range("A" & Valor).Select
End If


If Comboultimo.Value = "" Then
    CommandButton1.Visible = True
End If

Application.ScreenUpdating = True

End Sub

Private Sub Text_in1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text_in2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text_Quant_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Text_Preco_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57, 8, 44
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Combo_Produto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub Combo_INVOICE_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub Análise_Envio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

'===================================================================================================
'===================================================================================================
'==========================================RESUTILIZÁVEIS===========================================
'===================================================================================================
'===================================================================================================

Sub atualizar_combos()

Application.ScreenUpdating = False

Combo_Produto.Clear
Comboultimo.Clear

Dim w As Worksheet

Set w = Sheets("Controle")
w.Select
w.Range("B1").Select

Do While ActiveCell.Value <> ""
    Combo_Produto.AddItem ActiveCell.Value

ActiveCell.Offset(0, 1).Select
Loop

Range("A2").Select


Do While ActiveCell.Value <> ""

    If Mid(ActiveCell.Value, 1, 2) = "En" Then
       Comboultimo.AddItem Trim(Mid(ActiveCell.Value, 1, 8))
       Análise_Envio.AddItem Trim(Mid(ActiveCell.Value, 1, 8))
    End If

ActiveCell.Offset(1, 0).Select
Loop

Application.ScreenUpdating = True

End Sub

Private Sub Valor_Change()
Application.ScreenUpdating = False

Dim w As Worksheet

Set w = Sheets("Controle")

If Valor.Value = "" Then
Application.ScreenUpdating = True
Exit Sub

Else
    Range("a" & Valor).Select
    Do While ActiveCell.Value <> ""
        If Mid(ActiveCell.Value, 1, 2) = "De" Then
        GoTo pro
        End If
        If Mid(ActiveCell.Value, 1, 2) = "En" Then
            Application.ScreenUpdating = True
            Exit Sub
        Else
        Combo_INVOICE.AddItem ActiveCell.Value
        
        End If
pro:
        ActiveCell.Offset(1, 0).Select
    Loop
End If
Application.ScreenUpdating = True
End Sub

Sub limpar_tudo()
Combo_INVOICE.Value = ""
Text_Quant.Value = ""
Combo_Produto.Value = ""
TextBox2.Value = ""
Text_in2.Value = ""
Text_Preco.Value = ""
Text_in1.Value = ""
Comboultimo.Value = ""
Frame3.Visible = False
Bt_Alterar.Visible = True
    Bt_Alterar.Visible = False
    Text_in1.Visible = True
    Text_in2.Visible = True
    Label10.Visible = True
    Label8.Visible = True
    Label1.Visible = False
    CommandButton1.Visible = False
End Sub


Sub salvar()
Application.ScreenUpdating = False


If Text_Preco.Value = "" Or Text_in1.Value = "" Or Text_in2.Value = "" Then
    MsgBox "Por favor, preencha o preço e o intervalo do INVOICE!", , ""
    Application.ScreenUpdating = True
    Exit Sub
End If

If Text_in1.Value > Text_in2.Value Then
    MsgBox "Intervalo de INVOICE incorreto!", , ""
    Application.ScreenUpdating = True
    Exit Sub
End If
Frame3.Visible = True

Dim w As Worksheet

Set w = Sheets("Controle")
w.Select
w.Range("A2").Select

If ActiveCell.Value = "" Then
    ActiveCell.Value = "Envio 1" & " /" & Mid(Date, 7, 4) & " - Preço: " & Text_Preco
    Range("A" & ActiveCell.Row & ":Y" & ActiveCell.Row).Borders(xlEdgeTop).LineStyle = xlContinuous
    Range("A" & ActiveCell.Row & ":Y" & ActiveCell.Row).Borders(xlEdgeTop).Weight = xlThick
Else
    Do While ActiveCell.Value <> ""
    
    ActiveCell.Offset(1, 0).Select
    
    Loop
    ActiveCell.Value = Mid(Label1.Caption, 1, InStr(Label1.Caption, "-") - 2) & " /" & Mid(Date, 7, 4) & " - Preço: " & Text_Preco
    Range("A" & ActiveCell.Row & ":Y" & ActiveCell.Row).Borders(xlEdgeTop).LineStyle = xlContinuous
    Range("A" & ActiveCell.Row & ":Y" & ActiveCell.Row).Borders(xlEdgeTop).Weight = xlThick
End If


ActiveCell.Offset(1, 0).Value = "Demanda:"
ActiveCell.Offset(2, 0).Select


MF = Text_in1.Value


For A = MF To Text_in2.Value

    If MF < 10 Then
        ActiveCell.Value = "MF" & "0" & MF & Mid(Date, 9, 2)
    Else
        ActiveCell.Value = "MF" & MF & Mid(Date, 9, 2)
    End If
    MF = MF + 1

    ActiveCell.Offset(1, 0).Select
Next


Valor.Value = ""

Valor.Value = ActiveCell.Offset(-((Text_in2.Value - Text_in1.Value) + 2), 0).Row

atualizar_combos


Application.ScreenUpdating = True
End Sub



Sub colorir_list()
Dim i As Long


For i = 1 To ListView1.ListItems(2).ListSubItems.Count
        ListView1.ListItems.Item(2).ListSubItems(i).Bold = True
Next

End Sub



