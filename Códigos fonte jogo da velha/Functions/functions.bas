Attribute VB_Name = "functions"
Sub X()

If O_turn = True Then
    MsgBox "Vez do jogador O"
    Exit Sub
End If
    
    If ActiveCell.Value = "" Then
        ActiveCell.Value = "X"
    Else
        Exit Sub
    End If
    
    x_turn = False
    victory
    
    If x_turn = False Then
        Range("E4").Value = "VEZ DO JOGADOR O"
        O_turn = True
    End If
    
End Sub

Sub O()

If x_turn = True Then
    MsgBox "Vez do jogador X"
    Exit Sub
End If
    If ActiveCell.Value = "" Then
        ActiveCell.Value = "O"
    Else
        Exit Sub
    End If
    
    O_turn = False
    victory
    
    If O_turn = False Then
        Range("E4").Value = "VEZ DO JOGADOR DO X"
        x_turn = True
    End If
    
'    Range("E1:G3").ClearContents
'    Range("E1:G3").Interior.Pattern = xlNone
End Sub
Sub Button_change_color(bt As CommandButton)
    bt.BackColor = &H80000012
End Sub
Sub labe_change_color(lab)
    lab.BackStyle = fmBackStyleOpaque
    lab.BackColor = RGB(252, 163, 17)
End Sub

Function limparMatriz()
If victory_o = True Or victory_x = True Or velha = True Or novoJogo = True Then
For I = 0 To 2
    For j = 0 To 2
        If tabela(I, j) <> "" Then
        tabela(I, j) = ""
        End If
    Next
Next
End If
velha = False
novoJogo = False
End Function

Function colorir()

If Vertical = "ok" Then
    Select Case (I)
        Case 0
            UserForm1.Label7.BackColor = RGB(0, 255, 0)
            UserForm1.Label10.BackColor = RGB(0, 255, 0)
            UserForm1.Label13.BackColor = RGB(0, 255, 0)
            
        Case 1
            UserForm1.Label8.BackColor = RGB(0, 255, 0)
            UserForm1.Label11.BackColor = RGB(0, 255, 0)
            UserForm1.Label14.BackColor = RGB(0, 255, 0)
            
        Case 2
            UserForm1.Label9.BackColor = RGB(0, 255, 0)
            UserForm1.Label12.BackColor = RGB(0, 255, 0)
            UserForm1.Label15.BackColor = RGB(0, 255, 0)
            
    End Select
End If

If horizontal = "ok" Then
    Select Case (I)
    Case 0
        UserForm1.Label7.BackColor = RGB(0, 255, 0)
        UserForm1.Label8.BackColor = RGB(0, 255, 0)
        UserForm1.Label9.BackColor = RGB(0, 255, 0)
    Case 1
        UserForm1.Label10.BackColor = RGB(0, 255, 0)
        UserForm1.Label11.BackColor = RGB(0, 255, 0)
        UserForm1.Label12.BackColor = RGB(0, 255, 0)
    Case 2
        UserForm1.Label13.BackColor = RGB(0, 255, 0)
        UserForm1.Label14.BackColor = RGB(0, 255, 0)
        UserForm1.Label15.BackColor = RGB(0, 255, 0)
    End Select
End If

If diagonal1 = "ok" Then
    UserForm1.Label7.BackColor = RGB(0, 255, 0)
    UserForm1.Label11.BackColor = RGB(0, 255, 0)
    UserForm1.Label15.BackColor = RGB(0, 255, 0)
ElseIf diagonal2 = "ok" Then
    UserForm1.Label9.BackColor = RGB(0, 255, 0)
    UserForm1.Label11.BackColor = RGB(0, 255, 0)
    UserForm1.Label13.BackColor = RGB(0, 255, 0)
End If

End Function

Function zerar_label()
cont = 0
Vertical = ""
horizontal = ""
diagonal1 = ""
diagonal2 = ""
limparMatriz
victory_o = False
victory_x = False
novoJogo = True

UserForm1.Label7.BackStyle = fmBackStyleTransparent
UserForm1.Label8.BackStyle = fmBackStyleTransparent
UserForm1.Label9.BackStyle = fmBackStyleTransparent
UserForm1.Label10.BackStyle = fmBackStyleTransparent
UserForm1.Label11.BackStyle = fmBackStyleTransparent
UserForm1.Label12.BackStyle = fmBackStyleTransparent
UserForm1.Label13.BackStyle = fmBackStyleTransparent
UserForm1.Label14.BackStyle = fmBackStyleTransparent
UserForm1.Label15.BackStyle = fmBackStyleTransparent

UserForm1.Label7.Caption = ""
UserForm1.Label8.Caption = ""
UserForm1.Label9.Caption = ""
UserForm1.Label10.Caption = ""
UserForm1.Label11.Caption = ""
UserForm1.Label12.Caption = ""
UserForm1.Label13.Caption = ""
UserForm1.Label14.Caption = ""
UserForm1.Label15.Caption = ""

'UserForm1.Label7.BackColor = RGB(255, 255, 255)
'UserForm1.Label8.BackColor = RGB(255, 255, 255)
'UserForm1.Label9.BackColor = RGB(255, 255, 255)
'UserForm1.Label10.BackColor = RGB(255, 255, 255)
'UserForm1.Label11.BackColor = RGB(255, 255, 255)
'UserForm1.Label12.BackColor = RGB(255, 255, 255)
'UserForm1.Label13.BackColor = RGB(255, 255, 255)
'UserForm1.Label14.BackColor = RGB(255, 255, 255)
'UserForm1.Label15.BackColor = RGB(255, 255, 255)

End Function

