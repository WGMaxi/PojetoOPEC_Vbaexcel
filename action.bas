Attribute VB_Name = "action"
Sub Redimensiona_menor()
    Dim Message, Title, Default, MyValue
    Message = "MOSTRAR NA TELA 1 OU 2?"    ' Set prompt.
    Title = "ESCOLHA A TELA"    ' Set title.
    Default = "2"    ' Set default.

    ' Display dialog box at position 7500, 4500.
    MyValue = InputBox(Message, Title, Default, 7500, 4500)

    If MyValue = 2 Then
        With Application
            .WindowState = xlNormal
            .Top = 200
            .Left = 1920
            .Height = 150.5
            .Width = 10
        End With
        UserForm_copy.Top = -50
    End If

    If MyValue = 1 Then
        With Application
            .WindowState = xlNormal
            .Top = 200
            .Left = 950
            .Height = 150.5
            .Width = 10
        End With
        UserForm_copy.Top = -50
    End If
End Sub

Sub Redimensiona_maior()
    'UserForm_copy.StartUpPosition = 0
    With Application
        .WindowState = xlMaximized
        .Top = 0
        .Left = 0
        .Height = 550
        .Width = 1020
    End With
End Sub

Sub Redimensiona_maiorcomp()
    'UserForm_copy.StartUpPosition = 0
    With Application
        .WindowState = xlMaximized
    End With
End Sub

