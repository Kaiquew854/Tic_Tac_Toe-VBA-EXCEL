VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    zerar_label
End Sub


Private Sub UserForm_Initialize()
    x_turn = False
    O_turn = False
    Label1.Caption = jogador1
    Label2.Caption = jogador2
    
    
    Label11.BackStyle = fmBackStyleTransparent
    Label12.BackStyle = fmBackStyleTransparent
    Label13.BackStyle = fmBackStyleTransparent
    Label14.BackStyle = fmBackStyleTransparent
    Label15.BackStyle = fmBackStyleTransparent
    Label16.BackStyle = fmBackStyleTransparent
    
    Label7.BackStyle = fmBackStyleTransparent
    Label8.BackStyle = fmBackStyleTransparent
    Label9.BackStyle = fmBackStyleTransparent
    Label10.BackStyle = fmBackStyleTransparent
    Label11.BackStyle = fmBackStyleTransparent
    Label12.BackStyle = fmBackStyleTransparent
    Label13.BackStyle = fmBackStyleTransparent
    Label14.BackStyle = fmBackStyleTransparent
    Label15.BackStyle = fmBackStyleTransparent
    
    If x_turn = False And O_turn = False Then
        x_turn = True
        Label16.Caption = "Vez do jogador ""X"" " & jogador1
    End If
End Sub

Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call labe_change_color(Label7)
    
    Label8.BackStyle = fmBackStyleTransparent
    Label9.BackStyle = fmBackStyleTransparent
    Label10.BackStyle = fmBackStyleTransparent
    Label11.BackStyle = fmBackStyleTransparent
    Label12.BackStyle = fmBackStyleTransparent
    Label13.BackStyle = fmBackStyleTransparent
    Label14.BackStyle = fmBackStyleTransparent
    Label15.BackStyle = fmBackStyleTransparent
    
End Sub
Private Sub Label8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call labe_change_color(Label8)
    
    Label7.BackStyle = fmBackStyleTransparent
    Label9.BackStyle = fmBackStyleTransparent
    Label10.BackStyle = fmBackStyleTransparent
    Label11.BackStyle = fmBackStyleTransparent
    Label12.BackStyle = fmBackStyleTransparent
    Label13.BackStyle = fmBackStyleTransparent
    Label14.BackStyle = fmBackStyleTransparent
    Label15.BackStyle = fmBackStyleTransparent
    
End Sub
Private Sub Label9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call labe_change_color(Label9)
    
    Label7.BackStyle = fmBackStyleTransparent
    Label8.BackStyle = fmBackStyleTransparent
    Label10.BackStyle = fmBackStyleTransparent
    Label11.BackStyle = fmBackStyleTransparent
    Label12.BackStyle = fmBackStyleTransparent
    Label13.BackStyle = fmBackStyleTransparent
    Label14.BackStyle = fmBackStyleTransparent
    Label15.BackStyle = fmBackStyleTransparent
    
End Sub
Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call labe_change_color(Label10)
    
    Label7.BackStyle = fmBackStyleTransparent
    Label8.BackStyle = fmBackStyleTransparent
    Label9.BackStyle = fmBackStyleTransparent
    Label11.BackStyle = fmBackStyleTransparent
    Label12.BackStyle = fmBackStyleTransparent
    Label13.BackStyle = fmBackStyleTransparent
    Label14.BackStyle = fmBackStyleTransparent
    Label15.BackStyle = fmBackStyleTransparent
    
End Sub
Private Sub Label11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call labe_change_color(Label11)
    
    Label7.BackStyle = fmBackStyleTransparent
    Label8.BackStyle = fmBackStyleTransparent
    Label9.BackStyle = fmBackStyleTransparent
    Label10.BackStyle = fmBackStyleTransparent
    Label12.BackStyle = fmBackStyleTransparent
    Label13.BackStyle = fmBackStyleTransparent
    Label14.BackStyle = fmBackStyleTransparent
    Label15.BackStyle = fmBackStyleTransparent
    
End Sub
Private Sub Label12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call labe_change_color(Label12)
    
    Label7.BackStyle = fmBackStyleTransparent
    Label8.BackStyle = fmBackStyleTransparent
    Label9.BackStyle = fmBackStyleTransparent
    Label10.BackStyle = fmBackStyleTransparent
    Label11.BackStyle = fmBackStyleTransparent
    Label13.BackStyle = fmBackStyleTransparent
    Label14.BackStyle = fmBackStyleTransparent
    Label15.BackStyle = fmBackStyleTransparent

End Sub
Private Sub Label13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call labe_change_color(Label13)
    
    Label7.BackStyle = fmBackStyleTransparent
    Label8.BackStyle = fmBackStyleTransparent
    Label9.BackStyle = fmBackStyleTransparent
    Label10.BackStyle = fmBackStyleTransparent
    Label11.BackStyle = fmBackStyleTransparent
    Label12.BackStyle = fmBackStyleTransparent
    Label14.BackStyle = fmBackStyleTransparent
    Label15.BackStyle = fmBackStyleTransparent

End Sub
Private Sub Label14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call labe_change_color(Label14)
    
    Label7.BackStyle = fmBackStyleTransparent
    Label8.BackStyle = fmBackStyleTransparent
    Label9.BackStyle = fmBackStyleTransparent
    Label10.BackStyle = fmBackStyleTransparent
    Label11.BackStyle = fmBackStyleTransparent
    Label12.BackStyle = fmBackStyleTransparent
    Label13.BackStyle = fmBackStyleTransparent
    Label15.BackStyle = fmBackStyleTransparent
    
End Sub
Private Sub Label15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call labe_change_color(Label15)
    
    Label7.BackStyle = fmBackStyleTransparent
    Label8.BackStyle = fmBackStyleTransparent
    Label9.BackStyle = fmBackStyleTransparent
    Label10.BackStyle = fmBackStyleTransparent
    Label11.BackStyle = fmBackStyleTransparent
    Label12.BackStyle = fmBackStyleTransparent
    Label13.BackStyle = fmBackStyleTransparent
    Label14.BackStyle = fmBackStyleTransparent
    
End Sub
Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Call Button_change_color(CommandButton1)
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim bt As Object
    
    For Each bt In UserForm1.Controls
        If TypeName(bt) = "CommandButton" Then
            bt.BackColor = &H3D2114
        End If
    Next
    
    For Each lab In UserForm1.Controls
        If TypeName(lab) = "Label" Then
            lab.BackStyle = fmBackStyleTransparent = &H3D2114
        End If
    Next
    
End Sub



Private Sub Label7_Click()

    If x_turn = True And Label7.Caption = "" Then
        Label7.Caption = "X"
        Label7.ForeColor = RGB(51, 153, 137)
        x_turn = False
    End If
    
    If O_turn = True And Label7.Caption = "" Then
        Label7.Caption = "O"
        Label7.ForeColor = RGB(43, 44, 40)
        x_turn = True
        O_turn = False
    End If

    If x_turn = False Then
        O_turn = True
        Label16.Caption = "Vez do jogador ""O"" " & jogador2
     Else
        Label16.Caption = "Vez do jogador ""X"" " & jogador1
    End If
    tabela(0, 0) = UserForm1.Label7.Caption
    Tic_tac_toe


End Sub

Private Sub Label8_Click()

    If x_turn = True And Label8.Caption = "" Then
        Label8.Caption = "X"
        Label8.ForeColor = RGB(51, 153, 137)
        x_turn = False
    End If
    
    If O_turn = True And Label8.Caption = "" Then
        Label8.Caption = "O"
        Label8.ForeColor = RGB(43, 44, 40)
        x_turn = True
        O_turn = False
    End If

    If x_turn = False Then
        O_turn = True
        Label16.Caption = "Vez do jogador ""O"" " & jogador2
     Else
        Label16.Caption = "Vez do jogador ""X"" " & jogador1
    End If
    tabela(0, 1) = UserForm1.Label8.Caption
    Tic_tac_toe

End Sub

Private Sub Label9_Click()


    If x_turn = True And Label9.Caption = "" Then
        Label9.Caption = "X"
        Label9.ForeColor = RGB(51, 153, 137)
        x_turn = False
    End If
    
    If O_turn = True And Label9.Caption = "" Then
        Label9.Caption = "O"
        Label9.ForeColor = RGB(43, 44, 40)
        x_turn = True
        O_turn = False
    End If

    If x_turn = False Then
        O_turn = True
        Label16.Caption = "Vez do jogador ""O"" " & jogador2
     Else
        Label16.Caption = "Vez do jogador ""X"" " & jogador1
    End If
    tabela(0, 2) = UserForm1.Label9.Caption
    Tic_tac_toe

End Sub

Private Sub Label10_Click()

    If x_turn = True And Label10.Caption = "" Then
        Label10.Caption = "X"
        Label10.ForeColor = RGB(51, 153, 137)
        x_turn = False
    End If
    
    If O_turn = True And Label10.Caption = "" Then
        Label10.Caption = "O"
        Label10.ForeColor = RGB(43, 44, 40)
        x_turn = True
        O_turn = False
    End If

    If x_turn = False Then
        O_turn = True
        Label16.Caption = "Vez do jogador ""O"" " & jogador2
     Else
        Label16.Caption = "Vez do jogador ""X"" " & jogador1
    End If
    tabela(1, 0) = UserForm1.Label10.Caption
    Tic_tac_toe
        
End Sub

Private Sub Label11_Click()


    If x_turn = True And Label11.Caption = "" Then
        Label11.Caption = "X"
        Label11.ForeColor = RGB(51, 153, 137)
        x_turn = False
    End If
    
    If O_turn = True And Label11.Caption = "" Then
        Label11.Caption = "O"
        Label11.ForeColor = RGB(43, 44, 40)
        x_turn = True
        O_turn = False
    End If

    If x_turn = False Then
        O_turn = True
        Label16.Caption = "Vez do jogador ""O"" " & jogador2
     Else
        Label16.Caption = "Vez do jogador ""X"" " & jogador1
    End If
    tabela(1, 1) = UserForm1.Label11.Caption
    Tic_tac_toe

End Sub

Private Sub Label12_Click()

    If x_turn = True And Label12.Caption = "" Then
        Label12.Caption = "X"
        Label12.ForeColor = RGB(51, 153, 137)
        x_turn = False
    End If
    
    If O_turn = True And Label12.Caption = "" Then
        Label12.Caption = "O"
        Label12.ForeColor = RGB(43, 44, 40)
        x_turn = True
        O_turn = False
    End If

    If x_turn = False Then
        O_turn = True
        Label16.Caption = "Vez do jogador ""O"" " & jogador2
     Else
        Label16.Caption = "Vez do jogador ""X"" " & jogador1
    End If
    tabela(1, 2) = UserForm1.Label12.Caption
    Tic_tac_toe

End Sub

Private Sub Label13_Click()

    If x_turn = True And Label13.Caption = "" Then
        Label13.Caption = "X"
        Label13.ForeColor = RGB(51, 153, 137)
        x_turn = False
    End If
    
    If O_turn = True And Label13.Caption = "" Then
        Label13.Caption = "O"
        Label13.ForeColor = RGB(43, 44, 40)
        x_turn = True
        O_turn = False
    End If

    If x_turn = False Then
        O_turn = True
        Label16.Caption = "Vez do jogador ""O"" " & jogador2
     Else
        Label16.Caption = "Vez do jogador ""X"" " & jogador1
    End If
    tabela(2, 0) = UserForm1.Label13.Caption
    Tic_tac_toe

End Sub

Private Sub Label14_Click()

    If x_turn = True And Label14.Caption = "" Then
        Label14.Caption = "X"
        Label14.ForeColor = RGB(51, 153, 137)
        x_turn = False
    End If
    
    If O_turn = True And Label14.Caption = "" Then
        Label14.Caption = "O"
        Label14.ForeColor = RGB(43, 44, 40)
        x_turn = True
        O_turn = False
    End If
    
    If x_turn = False Then
        O_turn = True
        Label16.Caption = "Vez do jogador ""O"" " & jogador2
     Else
        Label16.Caption = "Vez do jogador ""X"" " & jogador1
    End If
    tabela(2, 1) = UserForm1.Label14.Caption
    Tic_tac_toe

End Sub

Private Sub Label15_Click()

    If x_turn = True And Label15.Caption = "" Then
        Label15.Caption = "X"
        Label15.ForeColor = RGB(51, 153, 137)
        x_turn = False
    End If
    
    If O_turn = True And Label15.Caption = "" Then
        Label15.Caption = "O"
        Label15.ForeColor = RGB(43, 44, 40)
        x_turn = True
        O_turn = False
    End If

    If x_turn = False Then
        O_turn = True
        Label16.Caption = "Vez do jogador ""O"" " & jogador2
     Else
        Label16.Caption = "Vez do jogador ""X"" " & jogador1
    End If
    tabela(2, 2) = UserForm1.Label15.Caption
    Tic_tac_toe

End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        zerar_label
        Unload UserForm1
    End If
End Sub
