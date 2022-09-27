Attribute VB_Name = "Main2"
Public tabela(2, 2) As String
Public Vertical As String
Public horizontal As String
Public diagonal1 As String
Public diagonal2 As String
Public cont As Integer
Public velha As Boolean
Public jogador1 As String
Public jogador2 As String
Public placarJog1 As Integer
Public placarJog2 As Integer
Public rodadas As Integer
Public placarVelha As Integer
Public ganhando As Integer
Public x_turn As Boolean
Public O_turn As Boolean
Public novoJogo As Boolean
Sub play()

victory_o = False
victory_x = False
placarJog1 = 0
placarJog2 = 0
placarVelha = 0
rodadas = 0


jogador1 = InputBox("Jogador 1, digite seu nome")
jogador2 = InputBox("Jogador 2, digite seu nome")

If jogador1 = "" Or jogador2 = "" Then
    Unload UserForm1
    Exit Sub
End If

UserForm1.Show
limparMatriz

Unload UserForm1
    
End Sub
Sub Tic_tac_toe()


For I = 0 To 2
    'vertical
    If tabela(0, I) = "X" And tabela(1, I) = "X" And tabela(2, I) = "X" Then
        victory_x = True
        placarJog1 = placarJog1 + 1
        UserForm1.Label4.Caption = placarJog1
        MsgBox (jogador1 & " Ganhou") '"X Wins"
        Vertical = "ok"
        Exit For
    ElseIf tabela(0, I) = "O" And tabela(1, I) = "O" And tabela(2, I) = "O" Then
        victory_o = True
        placarJog2 = placarJog2 + 1
        UserForm1.Label5.Caption = placarJog2
        MsgBox (jogador2 & " Ganhou") '"O Wins"
        Vertical = "ok"
        Exit For
    End If
    
    'horizontal
    If tabela(I, 0) = "X" And tabela(I, 1) = "X" And tabela(I, 2) = "X" Then
        horizontal = "ok"
        victory_x = True
        placarJog1 = placarJog1 + 1
        UserForm1.Label4.Caption = placarJog1
        MsgBox (jogador1 & " Ganhou") '"X Wins"
        Exit For
    ElseIf tabela(I, 0) = "O" And tabela(I, 1) = "O" And tabela(I, 2) = "O" Then
        horizontal = "ok"
        victory_o = True
        placarJog2 = placarJog2 + 1
        UserForm1.Label5.Caption = placarJog2
        MsgBox (jogador2 & " Ganhou") '"O Wins"
        Exit For
    End If
    
    'diagonal1
    If tabela(0, 0) = "X" And tabela(1, 1) = "X" And tabela(2, 2) = "X" Then
        diagonal1 = "ok"
        victory_x = True
        placarJog1 = placarJog1 + 1
        UserForm1.Label4.Caption = placarJog1
        MsgBox (jogador1 & " Ganhou") '"X Wins"
        Exit For
    ElseIf tabela(0, 0) = "O" And tabela(1, 1) = "O" And tabela(2, 2) = "O" Then
        diagonal1 = "ok"
        victory_o = True
        placarJog2 = placarJog2 + 1
        UserForm1.Label5.Caption = placarJog2
        MsgBox (jogador2 & " Ganhou") '"O Wins"
        Exit For
    End If
    
    'diagonal2
    If tabela(0, 2) = "X" And tabela(1, 1) = "X" And tabela(2, 0) = "X" Then
        diagonal2 = "ok"
        victory_x = True
        placarJog1 = placarJog1 + 1
        UserForm1.Label4.Caption = placarJog1
        MsgBox (jogador1 & "Ganhou") '"X Wins"
        Exit For
    ElseIf tabela(0, 2) = "O" And tabela(1, 1) = "O" And tabela(2, 0) = "O" Then
        diagonal2 = "ok"
        victory_o = True
        placarJog2 = placarJog2 + 1
        UserForm1.Label5.Caption = placarJog2
        MsgBox (jogador2 & "Ganhou") '"O Wins"
        Exit For
    End If
    
Next


If Vertical = "ok" Then

    Select Case (I)
        Case 0
            UserForm1.Label7.BackStyle = fmBackStyleOpaque
            UserForm1.Label10.BackStyle = fmBackStyleOpaque
            UserForm1.Label13.BackStyle = fmBackStyleOpaque
            UserForm1.Label7.BackColor = RGB(252, 163, 17)
            UserForm1.Label10.BackColor = RGB(252, 163, 17)
            UserForm1.Label13.BackColor = RGB(252, 163, 17)
        Case 1
        
            UserForm1.Label8.BackStyle = fmBackStyleOpaque
            UserForm1.Label11.BackStyle = fmBackStyleOpaque
            UserForm1.Label14.BackStyle = fmBackStyleOpaque
            UserForm1.Label8.BackColor = RGB(252, 163, 17)
            UserForm1.Label11.BackColor = RGB(252, 163, 17)
            UserForm1.Label14.BackColor = RGB(252, 163, 17)
        Case 2
            UserForm1.Label9.BackStyle = fmBackStyleOpaque
            UserForm1.Label12.BackStyle = fmBackStyleOpaque
            UserForm1.Label15.BackStyle = fmBackStyleOpaque
            UserForm1.Label9.BackColor = RGB(252, 163, 17)
            UserForm1.Label12.BackColor = RGB(252, 163, 17)
            UserForm1.Label15.BackColor = RGB(252, 163, 17)
    End Select

End If

If horizontal = "ok" Then
        Select Case (I)
        Case 0
            UserForm1.Label7.BackStyle = fmBackStyleOpaque
            UserForm1.Label8.BackStyle = fmBackStyleOpaque
            UserForm1.Label9.BackStyle = fmBackStyleOpaque
            UserForm1.Label7.BackColor = RGB(252, 163, 17)
            UserForm1.Label8.BackColor = RGB(252, 163, 17)
            UserForm1.Label9.BackColor = RGB(252, 163, 17)
            
        Case 1
            UserForm1.Label10.BackStyle = fmBackStyleOpaque
            UserForm1.Label11.BackStyle = fmBackStyleOpaque
            UserForm1.Label12.BackStyle = fmBackStyleOpaque
            UserForm1.Label10.BackColor = RGB(252, 163, 17)
            UserForm1.Label11.BackColor = RGB(252, 163, 17)
            UserForm1.Label12.BackColor = RGB(252, 163, 17)
            
        Case 2
            UserForm1.Label12.BackStyle = fmBackStyleOpaque
            UserForm1.Label14.BackStyle = fmBackStyleOpaque
            UserForm1.Label15.BackStyle = fmBackStyleOpaque
            UserForm1.Label13.BackColor = RGB(252, 163, 17)
            UserForm1.Label14.BackColor = RGB(252, 163, 17)
            UserForm1.Label15.BackColor = RGB(252, 163, 17)
            
        End Select
End If

If diagonal1 = "ok" Then
    UserForm1.Label7.BackStyle = fmBackStyleOpaque
    UserForm1.Label11.BackStyle = fmBackStyleOpaque
    UserForm1.Label15.BackStyle = fmBackStyleOpaque
    UserForm1.Label7.BackColor = RGB(252, 163, 17)
    UserForm1.Label11.BackColor = RGB(252, 163, 17)
    UserForm1.Label15.BackColor = RGB(252, 163, 17)
    
ElseIf diagonal2 = "ok" Then
    UserForm1.Label9.BackStyle = fmBackStyleOpaque
    UserForm1.Label11.BackStyle = fmBackStyleOpaque
    UserForm1.Label13.BackStyle = fmBackStyleOpaque
    UserForm1.Label9.BackColor = RGB(252, 163, 17)
    UserForm1.Label11.BackColor = RGB(252, 163, 17)
    UserForm1.Label13.BackColor = RGB(252, 163, 17)
    
End If

  
'velha
cont = 0
For I = 0 To 2
    For j = 0 To 2
        If tabela(I, j) <> "" Then
        cont = cont + 1
        End If
    Next
Next

If cont = 9 And victory_o = False And victory_x = False Then
    velha = True
    
    UserForm1.Label7.BackStyle = fmBackStyleOpaque
    UserForm1.Label8.BackStyle = fmBackStyleOpaque
    UserForm1.Label9.BackStyle = fmBackStyleOpaque
    UserForm1.Label10.BackStyle = fmBackStyleOpaque
    UserForm1.Label11.BackStyle = fmBackStyleOpaque
    UserForm1.Label12.BackStyle = fmBackStyleOpaque
    UserForm1.Label13.BackStyle = fmBackStyleOpaque
    UserForm1.Label14.BackStyle = fmBackStyleOpaque
    UserForm1.Label15.BackStyle = fmBackStyleOpaque
    UserForm1.Label7.BackColor = RGB(148, 28, 47)
    UserForm1.Label8.BackColor = RGB(148, 28, 47)
    UserForm1.Label9.BackColor = RGB(148, 28, 47)
    UserForm1.Label10.BackColor = RGB(148, 28, 47)
    UserForm1.Label11.BackColor = RGB(148, 28, 47)
    UserForm1.Label12.BackColor = RGB(148, 28, 47)
    UserForm1.Label13.BackColor = RGB(148, 28, 47)
    UserForm1.Label14.BackColor = RGB(148, 28, 47)
    UserForm1.Label15.BackColor = RGB(148, 28, 47)
    placarVelha = placarVelha + 1
    MsgBox "Deu velha"
    
End If

rodadas = placarJog1 + placarJog2 + placarVelha
UserForm1.Label6.Caption = rodadas

If placarJog1 > placarJog2 Then
    UserForm1.Label17.Caption = jogador1 & " Está vencendo"
ElseIf placarJog1 < placarJog2 Then
    UserForm1.Label17.Caption = jogador2 & " Está vencendo"
ElseIf placarJog1 = placarJog2 Then
    UserForm1.Label17.Caption = "Jogo empatado"
End If

'If victory_o = True Or victory_x = True Or velha = True Then
    'NewGame = MsgBox("Deseja jogar novamente", vbYesNo)
    'If NewGame = vbYes Then
        'zerar_label
    'Else
        'placarJog1 = 0
        'placarJog2 = 0
        'Unload UserForm1
    'End If
'End If

End Sub
