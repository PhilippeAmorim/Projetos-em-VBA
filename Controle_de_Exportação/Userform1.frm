VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Userform1 
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13380
   OleObjectBlob   =   "Userform1.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Userform1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

Text_PrazoDraft.Visible = False
Text_PrazoCarga.Visible = False
Prazo_Carga.Visible = False
Prazo_Draft.Visible = False
ID_For.Visible = False
Prazo.Visible = False
VG_Mudar_Status.Visible = False
Atualizar_Combos
atualizalista

With List_Pesquisar
    .BorderStyle = ccFixedSingle
    .Gridlines = True
    .View = lvwReport
    .ColumnHeaders.Add Text:="Status", Width:=79
    .ColumnHeaders.Add Text:="Marca", Width:=79, Alignment:=2
    .ColumnHeaders.Add Text:="Booking", Width:=108, Alignment:=2
    .ColumnHeaders.Add Text:="Draft", Width:=74, Alignment:=2
    .ColumnHeaders.Add Text:="Hora", Width:=74, Alignment:=2
    .ColumnHeaders.Add Text:="Carga", Width:=74, Alignment:=2
    .ColumnHeaders.Add Text:="Hora", Width:=74, Alignment:=2
    .ColumnHeaders.Add Text:="INVOICE", Width:=69, Alignment:=2
End With


End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ThisWorkbook.Save
    ThisWorkbook.Close
    Application.Quit
End Sub

'===================================================================================================
'==============================================BOTÕES===============================================
'===================================================================================================
Private Sub VG_Mudar_Status_Click()
Application.ScreenUpdating = False

Dim w       As Worksheet

Set w = Sheets("Movimentação")
w.Select
w.Range("I1048576").End(xlUp).Select

Do While ActiveCell.Value <> "INVOICE"

If ActiveCell.Value = Combo_Previsão.Value Then

    If ActiveCell.Offset(0, -1).Value = "Aguardando" Then
        ActiveCell.Offset(0, -1).Value = "Enviado"
        atualizalista2
        Application.ScreenUpdating = True
        Exit Sub
    End If
    If ActiveCell.Offset(0, -1).Value = "Enviado" Then
        ActiveCell.Offset(0, -1).Value = "Aguardando"
        atualizalista2
        Application.ScreenUpdating = True
        Exit Sub
    End If
End If
ActiveCell.Offset(-1, 0).Select
Loop

Application.ScreenUpdating = True

End Sub

Private Sub Adicionar_bt_Adicionar_Click()
Application.ScreenUpdating = False

Dim w   As Worksheet

Set w = Sheets("Bancos de dados")
w.Select


If ComboBox1.Value = "Marca" Then
w.Range("A2").Select
    Do While ActiveCell.Value <> ""
        If ActiveCell.Value = Adicionar_Nome.Value Then
            MsgBox "Já existe uma marca com esse nome!", , ""
            Exit Sub
        End If
        ActiveCell.Offset(1, 0).Select
    Loop
    ActiveCell.Value = Adicionar_Nome.Value
    MsgBox "Marca adicionada!", , ""
    ThisWorkbook.Save
ElseIf ComboBox1.Value = "Agente" Then
w.Range("B2").Select
Do While ActiveCell.Value <> ""
        If ActiveCell.Value = Adicionar_Nome.Value Then
            MsgBox "Já existe um agente com esse nome!", , ""
        End If
        ActiveCell.Offset(1, 0).Select
    Loop
    ActiveCell.Value = Adicionar_Nome.Value
    MsgBox "Agente adicionado!", , ""
    ThisWorkbook.Save
ElseIf ComboBox1.Value = "Armador" Then
w.Range("C2").Select
Do While ActiveCell.Value <> ""
        If ActiveCell.Value = Adicionar_Nome.Value Then
            MsgBox "Já existe um armador com esse nome!", , ""
        End If
        ActiveCell.Offset(1, 0).Select
    Loop
    ActiveCell.Value = Adicionar_Nome.Value
    MsgBox "Armador adicionado!", , ""
    ThisWorkbook.Save
Else
MsgBox "Prencha o campo de escolha!", , ""
Application.ScreenUpdating = True
Exit Sub
End If


Atualizar_Combos

Application.ScreenUpdating = True
End Sub


Private Sub CommandButton1_Click()
Application.ScreenUpdating = False

veri = InputBox("Digite a senha para acessar a planilha:", "")

If veri = "" Then
    Sheets("Bancos de dados").Select
    Userform1.Hide
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

Private Sub Adicionar_bt_Alterar_Click()
Application.ScreenUpdating = False

Dim w       As Worksheet

Set w = Sheets("Bancos de dados")
w.Select

If Adicionar_Alterar_Marca.Value <> "" Then
w.Range("A2").Select
    Do While ActiveCell.Value <> ""
        If ActiveCell.Value = Adicionar_Alterar_Marca.Value Then
            If Adicionar_Alterar_ProdutoTxt.Value <> "" Then
                ActiveCell.Value = Adicionar_Alterar_ProdutoTxt.Value
                GoTo proximo1
            Else
               MsgBox "O novo nome não pode ser vazio!", , ""
               Application.ScreenUpdating = True
               Exit Sub
            End If
        End If
        ActiveCell.Offset(1, 0).Select
    Loop
End If
proximo1:
If Adicionar_Alterar_AgenteCombo.Value <> "" Then
w.Range("B2").Select
    Do While ActiveCell.Value <> ""
        If ActiveCell.Value = Adicionar_Alterar_AgenteCombo.Value Then
            If Adicionar_Alterar_AgenteTxt.Value = "" Then
                MsgBox "O novo agente não pode ser vazio!", , ""
                Application.ScreenUpdating = True
                Exit Sub
            Else
                ActiveCell.Value = Adicionar_Alterar_AgenteTxt.Value
                GoTo proximo2
            End If
        End If
        ActiveCell.Offset(1, 0).Select
    Loop
End If
proximo2:
If Adicionar_Alterar_ArmadorCombo.Value <> "" Then
w.Range("c2").Select
    Do While ActiveCell.Value <> ""
        If ActiveCell.Value = Adicionar_Alterar_ArmadorCombo.Value Then
            If Adicionar_Alterar_ArmadorTxt.Value = "" Then
                MsgBox "O novo armador não pode ser vazio!", , ""
                Application.ScreenUpdating = True
                Exit Sub
            Else
                ActiveCell.Value = Adicionar_Alterar_ArmadorTxt.Value
                GoTo fim
            End If
        End If
        ActiveCell.Offset(1, 0).Select
    Loop
End If
fim:

MsgBox "Registo(s) alterado(s)!", , ""

Adicionar_Alterar_ArmadorTxt.Value = ""
Adicionar_Alterar_ArmadorCombo.Value = ""
Adicionar_Alterar_AgenteTxt.Value = ""
Adicionar_Alterar_AgenteCombo.Value = ""
Adicionar_Alterar_ProdutoTxt.Value = ""
Adicionar_Alterar_Marca.Value = ""

ThisWorkbook.Save

Atualizar_Combos

Application.ScreenUpdating = False
End Sub


Private Sub Adicionar_bt_Limpar1_Click()
Adicionar_Nome.Value = ""
ComboBox1.Value = ""
End Sub

Private Sub Adicionar_bt_Limpar2_Click()

Adicionar_Alterar_ArmadorTxt.Value = ""
Adicionar_Alterar_ArmadorCombo.Value = ""
Adicionar_Alterar_AgenteTxt.Value = ""
Adicionar_Alterar_AgenteCombo.Value = ""
Adicionar_Alterar_ProdutoTxt.Value = ""
Adicionar_Alterar_Marca.Value = ""

End Sub


Private Sub Pesquisar_bt_Limpar_Click()
Combo_Previsão.Clear
Pesquisar_Marca.Value = ""
TextBox1.Value = ""
Pesquisar_Booking.Value = ""
Pesquisar_Status.Value = ""
Pesquisar_Draft.Value = "24/09/2021"
Pesquisar_Carga.Value = "24/09/2021"

atualizalista
End Sub

Private Sub Pesquisar_bt_pesquisar_Click()
Application.ScreenUpdating = False

Dim w As Worksheet

Combo_Previsão.Clear

Set w = Sheets("Movimentação")
w.Select

   
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=2
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=3
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=4
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=5
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=6
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=7
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=8
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=9
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=10
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=11
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=12
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=13
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=14
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=15
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=16
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=17
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=18
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=19
    

    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=1, Criteria1:= _
            "=*" & Pesquisar_Marca.Value & "*", Operator:=xlAnd
            
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=9, Criteria1:= _
            "=*" & TextBox1.Value & "*", Operator:=xlAnd
    
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=4, Criteria1:= _
            "=*" & Pesquisar_Booking.Value & "*", Operator:=xlAnd
            
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=8, Criteria1:= _
            "=" & Pesquisar_Status.Value & "*", Operator:=xlAnd
                   

If Pesquisar_Draft.Value <> "24/09/2021" Or Pesquisar_Carga.Value <> "24/09/2021" Then
    If Pesquisar_Draft.Value <> "24/09/2021" And Pesquisar_Carga.Value = "24/09/2021" Then
    
        ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=14, Criteria1:= _
            "=" & Pesquisar_Draft.Value & "", Operator:=xlAnd
    ElseIf Pesquisar_Carga.Value <> "24/09/2021" And Pesquisar_Draft.Value = "24/09/2021" Then
            
        ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=17, Criteria1:= _
            "=" & Pesquisar_Carga.Value & "", Operator:=xlAnd
    ElseIf Pesquisar_Draft.Value <> "24/09/2021" And Pesquisar_Carga.Value <> "24/09/2021" Then
        ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=14, Criteria1:= _
            "=" & Pesquisar_Draft.Value & "", Operator:=xlAnd
        ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=17, Criteria1:= _
            "=" & Pesquisar_Carga.Value & "", Operator:=xlAnd
    End If
    
End If
            
            
    List_Pesquisar.ListItems.Clear
    
    
    lin = 2
    
    Do Until Sheets("Movimentação").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
    
        Set li = List_Pesquisar.ListItems.Add(Text:=Sheets("Movimentação").Cells(lin, 8).Value)
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 1).Value
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 4).Value
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 12).Value
         li.ListSubItems.Add Text:=Format(Sheets("Movimentação").Cells(lin, 13).Value, "hh:mm:ss")
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 15).Value
         li.ListSubItems.Add Text:=Format(Sheets("Movimentação").Cells(lin, 16).Value, "hh:mm:ss")
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 9).Value

        Combo_Previsão.AddItem Sheets("Movimentação").Cells(lin, 9).Value
    End If
        lin = lin + 1
    Loop
    colorir_list
    Application.ScreenUpdating = True
    Exit Sub


End Sub

Private Sub Bt_Alterar_Click()

Dim w   As Worksheet
Dim l   As Long

Set w = Sheets("Movimentação")
w.Select
w.Range("a2").Select

If ID_For.Value = "" Then
MsgBox "Selecione um registro para alterar", , ""
Application.ScreenUpdating = True
Exit Sub
End If

Do While ActiveCell.Value <> ""
c = ActiveCell.Row

    If ActiveCell.Offset(0, 8).Value = ID_For.Value Then
        If Range("A" & c).Value <> Combo_Marca.Value Then
        w.Range("I1048576").End(xlUp).Select
        
        
        Do While ActiveCell.Value <> "INVOICE"
        
        
            If Mid(ActiveCell.Value, 1, 2) = Combo_Marca Then
            
                If Mid(ActiveCell.Value, 5, 2) <> Mid(Date, 9, 2) Then
                    w.Range("A" & c).Offset(0, 8).Value = Combo_Marca & "01" & Mid(Date, 9, 2)
                    GoTo proci
                End If
                
                If Mid(ActiveCell.Value, 3, 2) <= 8 Then
                     w.Range("A" & c).Offset(0, 8).Value = Combo_Marca.Value & "0" & (Mid(ActiveCell.Value, 3, 2) + 1) & Mid(Date, 9, 2)
                Else
                    w.Range("A" & c).Offset(0, 8).Value = Combo_Marca.Value & (Mid(ActiveCell.Value, 3, 2) + 1) & Mid(Date, 9, 2)
                End If
            GoTo proci
            End If
        
        ActiveCell.Offset(-1, 0).Select
        Loop
        
         w.Range("A" & c).Offset(0, 8).Value = Combo_Marca & "01" & Mid(Date, 9, 2)
         End If
proci:
            Range("A" & c).Select
            ActiveCell.Value = Combo_Marca.Value
            ActiveCell.Offset(0, 1) = Combo_Agente
            ActiveCell.Offset(0, 2) = Combo_Armador
            ActiveCell.Offset(0, 3) = Text_Booking
            ActiveCell.Offset(0, 4) = Text_Destino
            ActiveCell.Offset(0, 5) = Text_Transportador
            ActiveCell.Offset(0, 6) = Text_Depot
            ActiveCell.Offset(0, 7) = Combo_Status
            
                    ActiveCell.Offset(0, 10) = Tx_DataDeFabricacao.Value
                    ActiveCell.Offset(0, 11) = DT_Draft.Value
                    ActiveCell.Offset(0, 12) = DT_Hora.Value
                    If OptionButton1.Value = True Then
                        ActiveCell.Offset(0, 13) = "Enviado"
                    Else
                        ActiveCell.Offset(0, 13) = "Não enviado"
                    End If
                    ActiveCell.Offset(0, 14) = DT_Carga.Value
                    ActiveCell.Offset(0, 15) = DT_Hora_Carga.Value
                    If OptionButton3.Value = True Then
                        ActiveCell.Offset(0, 16) = "Enviado"
                    Else
                        ActiveCell.Offset(0, 16) = "Não enviado"
                    End If
                    
                    ActiveCell.Offset(0, 17) = DT_ETD.Value
                    ActiveCell.Offset(0, 18) = DT_ETA.Value
         
         Atualizar_Combos
         limpar_campos
         atualizalista
         MsgBox "Registro alterado com sucesso!", , ""
         Application.ScreenUpdating = True
        Exit Sub
    End If

    ActiveCell.Offset(1, 0).Select
Loop

End Sub

Private Sub Bt_Deletar_Click()
Application.ScreenUpdating = False

Dim w   As Worksheet

Set w = Sheets("Movimentação")
w.Select
w.Range("I2").Select

If ID_For.Value = "" Then
MsgBox "Selecione um registro para deletar", , ""
         Application.ScreenUpdating = True
        Exit Sub
End If

Do While ActiveCell.Value <> ""
    If ActiveCell.Value = ID_For.Value Then
         ActiveCell.EntireRow.Delete
         Atualizar_Combos
         limpar_campos
         atualizalista
         MsgBox "Registro deletado com sucesso!", , ""
         Application.ScreenUpdating = True
        Exit Sub
    End If

    ActiveCell.Offset(1, 0).Select
Loop

End Sub

Private Sub Bt_Limpar_Click()
    limpar_campos
End Sub

Private Sub Bt_Busca_Click()
Application.ScreenUpdating = False

Dim ID  As String
On Error GoTo erro

ID = InputBox("Digite o ID do item: ", "")
ID_For.Value = UCase(ID)

erro:

Application.ScreenUpdating = True
End Sub

Private Sub Bt_Adicionar_Click()
Application.ScreenUpdating = False

If Combo_Marca = "" Or Combo_Agente = "" Or Combo_Armador = "" Or Text_Booking = "" Or Text_Destino = "" Or Text_Transportador = "" Or Text_Depot = "" Or Combo_Status = "" Then
    MsgBox "Preencha todos os dados!", vbOKOnly, ""
    Application.ScreenUpdating = True
    Exit Sub
End If

Dim w   As Worksheet

Set w = Sheets("Movimentação")
w.Select
w.Range("A2").Select


Do While ActiveCell.Value <> ""
ActiveCell.Offset(1, 0).Select
Loop

ActiveCell.Value = Combo_Marca
ActiveCell.Offset(0, 1) = Combo_Agente
ActiveCell.Offset(0, 2) = Combo_Armador
ActiveCell.Offset(0, 3) = Text_Booking
ActiveCell.Offset(0, 4) = Text_Destino
ActiveCell.Offset(0, 5) = Text_Transportador
ActiveCell.Offset(0, 6) = Text_Depot
ActiveCell.Offset(0, 7) = Combo_Status

'============================================================
w.Range("I1048576").End(xlUp).Select


Do While ActiveCell.Value <> "INVOICE"


    If Mid(ActiveCell.Value, 1, 2) = Combo_Marca Then
    
        If Mid(ActiveCell.Value, 5, 2) <> Mid(Date, 9, 2) Then
            w.Range("I1048576").End(xlUp).Value = Combo_Marca & "01" & Mid(Date, 9, 2)
            GoTo proci
        End If
        
        If Mid(ActiveCell.Value, 3, 2) <= 8 Then
            w.Range("I1048576").End(xlUp).Value = Combo_Marca.Value & "0" & (Mid(ActiveCell.Value, 3, 2) + 1) & Mid(Date, 9, 2)
        Else
            w.Range("I1048576").End(xlUp).Value = Combo_Marca.Value & (Mid(ActiveCell.Value, 3, 2) + 1) & Mid(Date, 9, 2)
        End If
    GoTo proci
    End If

ActiveCell.Offset(-1, 0).Select
Loop

w.Range("I1048576").End(xlUp).Value = Combo_Marca & "01" & Mid(Date, 9, 2)
'============================================================
proci:

w.Range("A2").Select


Do While ActiveCell.Value <> ""
ActiveCell.Offset(1, 0).Select
Loop
ActiveCell.Offset(-1, 0).Select

        ActiveCell.Offset(0, 10) = Tx_DataDeFabricacao.Value
        ActiveCell.Offset(0, 11) = DT_Draft.Value
        ActiveCell.Offset(0, 12) = DT_Hora.Value
        If OptionButton1.Value = True Then
            ActiveCell.Offset(0, 13) = "Enviado"
        Else
            ActiveCell.Offset(0, 13) = "Não enviado"
        End If
        ActiveCell.Offset(0, 14) = DT_Carga.Value
        ActiveCell.Offset(0, 15) = DT_Hora_Carga.Value
        If OptionButton3.Value = True Then
            ActiveCell.Offset(0, 16) = "Enviado"
        Else
            ActiveCell.Offset(0, 16) = "Não enviado"
        End If
        
        ActiveCell.Offset(0, 17) = DT_ETD.Value
        ActiveCell.Offset(0, 18) = DT_ETA.Value

MsgBox "Processo concluído!", vbOKOnly, ""

Atualizar_Combos

limpar_campos

atualizalista

Application.ScreenUpdating = True
End Sub




'===================================================================================================
'==============================================LAYOUT===============================================
'===================================================================================================
Private Sub Combo_Previsão_Change()
If Combo_Previsão.Value <> "" Then
Application.ScreenUpdating = False
Prazo_Carga.Visible = True
Prazo_Draft.Visible = True
VG_Mudar_Status.Visible = True

Dim w       As Worksheet

Set w = Sheets("Movimentação")
w.Select
w.Range("I1048576").End(xlUp).Select

Do While ActiveCell.Value <> ""

If ActiveCell.Value = Combo_Previsão.Value Then

    If ActiveCell.Offset(0, 5).Value = "" Then
        Prazo_Draft.Value = "Draft não cadastrado!"
        GoTo tem1
    End If
    
    If ActiveCell.Offset(0, 5).Value = "Enviado" Then
        Prazo_Draft.Value = "Draft já enviado!"
        GoTo tem1
    End If

    If ActiveCell.Offset(0, 3).Value - Date > 0 Then
                Prazo_Draft.Value = ActiveCell.Offset(0, 3).Value - Date & " dia(s)" & ", " _
                & Format(ActiveCell.Offset(0, 4).Value - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(ActiveCell.Offset(0, 4).Value - Time, "hh:mm"), 4, 2) & " minuto(s)"
                
    ElseIf ActiveCell.Offset(0, 3).Value - Date = 0 Then
               If ActiveCell.Offset(0, 4).Value > Time Then
                    Prazo_Draft.Value = "Vence em " & _
                    Format(ActiveCell.Offset(0, 4).Value - Time, "hh") & " hora(s)" & " e " & _
                    Mid(Format(ActiveCell.Offset(0, 4).Value - Time, "hh:mm"), 4, 2) & " minuto(s)"
                Else
                    Prazo_Draft.Value = "Vencido a " & _
                    Format(ActiveCell.Offset(0, 4).Value - Time, "hh") & " hora(s)" & " e " & _
                    Mid(Format(ActiveCell.Offset(0, 4).Value - Time, "hh:mm"), 4, 2) & " minuto(s)"
                
                End If
    Else
                Prazo_Draft.Value = "Vencido a " & (ActiveCell.Offset(0, 3).Value - Date) * -1 & " dia(s)"
    End If
            
tem1:
            
If ActiveCell.Offset(0, 8).Value = "" Then
        Prazo_Carga.Value = "Carga não cadastrado!"
        Exit Sub
End If

If ActiveCell.Offset(0, 8).Value = "Enviado" Then
        Prazo_Carga.Value = "Draft já enviado!"
        Exit Sub
End If

        
If ActiveCell.Offset(0, 6).Value - Date > 0 Then
            Prazo_Carga.Value = ActiveCell.Offset(0, 6).Value - Date & " dia(s)" & ", " _
            & Format(ActiveCell.Offset(0, 7).Value - Time, "hh") & " hora(s)" & " e " & _
            Mid(Format(ActiveCell.Offset(0, 7).Value - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
        ElseIf ActiveCell.Offset(0, 6).Value - Date = 0 Then
            If ActiveCell.Offset(0, 7).Value > Time Then
                Prazo_Carga.Value = "Vence em " & _
                Format(ActiveCell.Offset(0, 7).Value - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(ActiveCell.Offset(0, 7).Value - Time, "hh:mm"), 4, 2) & " minuto(s)"
            Else
                Prazo_Carga.Value = "Vencido a " & _
                Format(ActiveCell.Offset(0, 7).Value - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(ActiveCell.Offset(0, 7).Value - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
            End If
        Else
            Prazo_Carga.Value = "Vencido a " & (ActiveCell.Offset(0, 6).Value - Date) * -1 & " dia(s)"
        End If
        

Application.ScreenUpdating = True
Exit Sub
End If

ActiveCell.Offset(-1, 0).Select
Loop


Else
Prazo_Carga.Visible = False
Prazo_Draft.Visible = False
End If
End Sub

Private Sub Adicionar_Alterar_AgenteCombo_Change()
Adicionar_Alterar_AgenteTxt.Value = Adicionar_Alterar_AgenteCombo.Value
End Sub

Private Sub Adicionar_Alterar_ArmadorCombo_Change()
Adicionar_Alterar_ArmadorTxt.Value = Adicionar_Alterar_ArmadorCombo.Value
End Sub


Private Sub DT_Hora_Change()
    If OptionButton2.Value = True Then
        If DT_Draft.Value - Date > 0 Then
            Text_PrazoDraft.Value = DT_Draft.Value - Date & " dia(s)" & ", " _
            & Format(TimeValue(DT_Hora.Value) - Time, "hh") & " hora(s)" & " e " & _
            Mid(Format(TimeValue(DT_Hora.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
        ElseIf DT_Draft.Value - Date = 0 Then
            If TimeValue(DT_Hora.Value) > Time Then
                Text_PrazoDraft.Value = "Vence em " & _
                Format(TimeValue(DT_Hora.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            Else
                Text_PrazoDraft.Value = "Vencido a " & _
                Format(TimeValue(DT_Hora.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
            End If
        Else
            Text_PrazoDraft.Value = "Vencido a " & (DT_Draft.Value - Date) * -1 & " dia(s)"
        End If
    End If
End Sub

Private Sub DT_Hora_Carga_Change()
If OptionButton4.Value = True Then
        If DT_Carga.Value - Date > 0 Then
            Text_PrazoCarga.Value = DT_Carga.Value - Date & " dia(s)" & ", " _
            & Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh") & " hora(s)" & " e " & _
            Mid(Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
        ElseIf DT_Carga.Value - Date = 0 Then
            If TimeValue(DT_Hora_Carga.Value) > Time Then
                Text_PrazoCarga.Value = "Vence em " & _
                Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            Else
                Text_PrazoCarga.Value = "Vencido a " & _
                Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
            End If
        Else
            Text_PrazoCarga.Value = "Vencido a " & (DT_Carga.Value - Date) * -1 & " dia(s)"
        End If
    End If
End Sub

Private Sub DT_Draft_Change()
    If OptionButton2.Value = True Then
        If DT_Draft.Value - Date > 0 Then
            Text_PrazoDraft.Value = DT_Draft.Value - Date & " dia(s)" & ", " _
            & Format(TimeValue(DT_Hora.Value) - Time, "hh") & " hora(s)" & " e " & _
            Mid(Format(TimeValue(DT_Hora.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
        ElseIf DT_Draft.Value - Date = 0 Then
            If TimeValue(DT_Hora.Value) > Time Then
                Text_PrazoDraft.Value = "Vence em " & _
                Format(TimeValue(DT_Hora.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            Else
                Text_PrazoDraft.Value = "Vencido a " & _
                Format(TimeValue(DT_Hora.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
            End If
        Else
            Text_PrazoDraft.Value = "Vencido a " & (DT_Draft.Value - Date) * -1 & " dia(s)"
        End If
    End If
End Sub

Private Sub DT_Carga_Change()
    If OptionButton4.Value = True Then
        If DT_Carga.Value - Date > 0 Then
            Text_PrazoCarga.Value = DT_Carga.Value - Date & " dia(s)" & ", " _
            & Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh") & " hora(s)" & " e " & _
            Mid(Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
        ElseIf DT_Carga.Value - Date = 0 Then
            If TimeValue(DT_Hora_Carga.Value) > Time Then
                Text_PrazoCarga.Value = "Vence em " & _
                Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            Else
                Text_PrazoCarga.Value = "Vencido a " & _
                Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
            End If
        Else
            Text_PrazoCarga.Value = "Vencido a " & (DT_Carga.Value - Date) * -1 & " dia(s)"
        End If
    End If
End Sub

Private Sub OptionButton2_Click()
Prazo.Visible = True
Text_PrazoDraft.Visible = True
    If OptionButton2.Value = True Then
        If DT_Draft.Value - Date > 0 Then
            Text_PrazoDraft.Value = DT_Draft.Value - Date & " dia(s)" & ", " _
            & Format(TimeValue(DT_Hora.Value) - Time, "hh") & " hora(s)" & " e " & _
            Mid(Format(TimeValue(DT_Hora.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
        ElseIf DT_Draft.Value - Date = 0 Then
           If TimeValue(DT_Hora.Value) > Time Then
                Text_PrazoDraft.Value = "Vence em " & _
                Format(TimeValue(DT_Hora.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            Else
                Text_PrazoDraft.Value = "Vencido a " & _
                Format(TimeValue(DT_Hora.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
            End If
        Else
            Text_PrazoDraft.Value = "Vencido a " & (DT_Draft.Value - Date) * -1 & " dia(s)"
        End If
    End If
End Sub

Private Sub OptionButton1_Click()
     If OptionButton1.Value = True Then
        Prazo.Visible = False
        Text_PrazoDraft.Visible = False
    End If

End Sub

Private Sub OptionButton4_Click()
Prazo2.Visible = True
Text_PrazoCarga.Visible = True
    If OptionButton4.Value = True Then
        If DT_Carga.Value - Date > 0 Then
            Text_PrazoCarga.Value = DT_Carga.Value - Date & " dia(s)" & ", " _
            & Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh") & " hora(s)" & " e " & _
            Mid(Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
        ElseIf DT_Carga.Value - Date = 0 Then
            If TimeValue(DT_Hora_Carga.Value) > Time Then
                Text_PrazoCarga.Value = "Vence em " & _
                Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            Else
                Text_PrazoCarga.Value = "Vencido a " & _
                Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh") & " hora(s)" & " e " & _
                Mid(Format(TimeValue(DT_Hora_Carga.Value) - Time, "hh:mm"), 4, 2) & " minuto(s)"
            
            End If
        Else
            Text_PrazoCarga.Value = "Vencido a " & (DT_Carga.Value - Date) * -1 & " dia(s)"
        End If
    End If
End Sub

Private Sub OptionButton3_Click()
     If OptionButton3.Value = True Then
        Prazo2.Visible = False
        Text_PrazoCarga.Visible = False
    End If

End Sub

'===================================================================================================
'============================================RESTRIÇÕES=============================================
'===================================================================================================

Private Sub Combo_Agente_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub Combo_Armador_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub Combo_Status_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub Text_PrazoDraft_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub Text_PrazoCarga_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub Pesquisar_Status_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub ComboBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub Adicionar_Alterar_AgenteCombo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub Adicionar_Alterar_ArmadorCombo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

Private Sub Adicionar_Nome_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ComboBox1.Value = "Marca" Then
        Adicionar_Nome.MaxLength = 2
    End If
End Sub
Private Sub Combo_Previsão_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii <> 1000 Then KeyAscii = 0
End Sub

'===================================================================================================
'=========================================SUBS REUTILIZÁVEIS========================================
'===================================================================================================

Sub Atualizar_Combos()
Application.ScreenUpdating = False

Combo_Status.Clear
Combo_Agente.Clear
Combo_Armador.Clear
Pesquisar_Status.Clear
Combo_Previsão.Clear

Combo_Marca.Clear
Adicionar_Alterar_AgenteCombo.Clear
Adicionar_Alterar_ArmadorCombo.Clear
ComboBox1.Clear
Pesquisar_Marca.Clear

ComboBox1.AddItem "Marca"
ComboBox1.AddItem "Agente"
ComboBox1.AddItem "Armador"
Combo_Status.AddItem "Aguardando"
Combo_Status.AddItem "Enviado"
Pesquisar_Status.AddItem "Aguardando"
Pesquisar_Status.AddItem "Enviado"

Dim w   As Worksheet

Set w = Sheets("Bancos de dados")
w.Select
w.Range("A2").Select

Do While ActiveCell.Value <> ""

    Combo_Marca.AddItem ActiveCell.Value
    Pesquisar_Marca.AddItem ActiveCell.Value
    Adicionar_Alterar_Marca.AddItem ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
Loop
w.Range("b2").Select
Do While ActiveCell.Value <> ""
Combo_Agente.AddItem ActiveCell.Value
Adicionar_Alterar_AgenteCombo.AddItem ActiveCell.Value
ActiveCell.Offset(1, 0).Select
Loop
w.Range("c2").Select
Do While ActiveCell.Value <> ""
Combo_Armador.AddItem ActiveCell.Value
Adicionar_Alterar_ArmadorCombo.AddItem ActiveCell.Value
ActiveCell.Offset(1, 0).Select
Loop

Application.ScreenUpdating = True
End Sub

Private Sub ID_For_Change()
Application.ScreenUpdating = False
    
Dim w As Worksheet

Set w = Sheets("Movimentação")
w.Select
w.Range("A1048576").End(xlUp).Select


Do While ActiveCell.Value <> "Nome do produto"
On Error GoTo erro

    If ActiveCell.Offset(0, 8).Value = ID_For.Value Then
         Combo_Marca.Value = ActiveCell.Value
         Combo_Agente.Value = ActiveCell.Offset(0, 1).Value
         Combo_Armador.Value = ActiveCell.Offset(0, 2).Value
         Text_Booking.Value = ActiveCell.Offset(0, 3).Value
         Text_Destino.Value = ActiveCell.Offset(0, 4).Value
         Text_Transportador.Value = ActiveCell.Offset(0, 5).Value
         Text_Depot.Value = ActiveCell.Offset(0, 6).Value
         Combo_Status.Value = ActiveCell.Offset(0, 7).Value
         
         If Combo_Status.Value = "Enviado" Then
            Tx_DataDeFabricacao.Value = ActiveCell.Offset(0, 10).Value
            DT_Draft.Value = ActiveCell.Offset(0, 11).Value
            DT_Hora.Value = ActiveCell.Offset(0, 12).Value
            If ActiveCell.Offset(0, 13).Value = "Enviado" Then
                OptionButton1.Value = True
            Else
                OptionButton2.Value = True
            End If
            DT_Carga.Value = ActiveCell.Offset(0, 14).Value
            DT_Hora_Carga.Value = ActiveCell.Offset(0, 15).Value
            If ActiveCell.Offset(0, 16).Value = "Enviado" Then
                OptionButton3.Value = True
            Else
                OptionButton4.Value = True
            End If
            DT_ETD.Value = ActiveCell.Offset(0, 17).Value
            DT_ETA.Value = ActiveCell.Offset(0, 18).Value
         End If
    
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ActiveCell.Offset(-1, 0).Select
Loop


erro:
MsgBox "ID não encontrado, tente novamente!", , ""

Application.ScreenUpdating = True
End Sub

Sub limpar_campos()

Combo_Marca.Value = ""
Combo_Agente.Value = ""
Combo_Agente.Value = ""
Combo_Armador.Value = ""
Text_Booking.Value = ""
Text_Destino.Value = ""
Text_Transportador.Value = ""
Text_Depot = ""
Combo_Status = ""
OptionButton1.Value = False
OptionButton2.Value = False
OptionButton3.Value = False
OptionButton4.Value = False
ID_For.Value = ""



End Sub

Sub atualizalista()
Application.ScreenUpdating = False

Sheets("Movimentação").Select

    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=8
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=1
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=4
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=12
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=15
    ActiveSheet.ListObjects("Tabela1").Range.AutoFilter Field:=9

List_Pesquisar.ListItems.Clear


lin = 2

Do Until Sheets("Movimentação").Cells(lin, 1) = ""

    Set li = List_Pesquisar.ListItems.Add(Text:=Sheets("Movimentação").Cells(lin, 8).Value)
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 1).Value
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 4).Value
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 12).Value
         li.ListSubItems.Add Text:=Format(Sheets("Movimentação").Cells(lin, 13).Value, "hh:mm:ss")
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 15).Value
         li.ListSubItems.Add Text:=Format(Sheets("Movimentação").Cells(lin, 16).Value, "hh:mm:ss")
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 9).Value


    lin = lin + 1
Loop

colorir_list

Application.ScreenUpdating = True

End Sub
Sub atualizalista2()
Application.ScreenUpdating = False

Sheets("Movimentação").Select
List_Pesquisar.ListItems.Clear


lin = 2

Do Until Sheets("Movimentação").Cells(lin, 1) = ""
    
    If Cells(lin, 1).Rows.Hidden = False Then
    
    
        Set li = List_Pesquisar.ListItems.Add(Text:=Sheets("Movimentação").Cells(lin, 8).Value)
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 1).Value
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 4).Value
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 12).Value
         li.ListSubItems.Add Text:=Format(Sheets("Movimentação").Cells(lin, 13).Value, "hh:mm:ss")
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 15).Value
         li.ListSubItems.Add Text:=Format(Sheets("Movimentação").Cells(lin, 16).Value, "hh:mm:ss")
         li.ListSubItems.Add Text:=Sheets("Movimentação").Cells(lin, 9).Value
    End If
        lin = lin + 1
    Loop
    colorir_list
    Application.ScreenUpdating = True
    Exit Sub

End Sub

Sub colorir_list()
Dim i As Long


For i = 1 To List_Pesquisar.ListItems.Count
    If List_Pesquisar.ListItems.Item(i).Text = "Aguardando" Then
        List_Pesquisar.ListItems.Item(i).ForeColor = RGB(225, 0, 0)
    ElseIf List_Pesquisar.ListItems.Item(i).Text = "Enviado" Then
        List_Pesquisar.ListItems.Item(i).ForeColor = RGB(0, 176, 80)
    End If
Next

End Sub
'===================================================================================================
'===================================================================================================
'===================================================================================================


