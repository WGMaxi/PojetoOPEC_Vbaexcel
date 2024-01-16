Attribute VB_Name = "inserir"
Sub copy()

Call Redimensiona_menor
UserForm_copy.Show

End Sub
Sub INSERIR_info()
    Dim NovaLinha As Range
    Dim Tbmaio As Object
      'Inserir uma linha no final da tabela.
      ' ActiveSheet.ListObjects("maio").ListRows.Add AlwaysInsert:=True
        tabela = UserForm_copy.TextB_m.Value
        Set Tbmaio = ActiveSheet.ListObjects(tabela)
        Set NovaLinha = Tbmaio.ListRows.Add.Range
        Tbmaio.ListColumns("cliente").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.TextBoxCLI.Text)
        Tbmaio.ListColumns("negociação").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.TextBoxNEG.Text)
        Tbmaio.ListColumns("pi").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.TextBoxPI.Text)
        Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.TextBoxCONTR.Text)
        Tbmaio.ListColumns("término").Range.Rows(NovaLinha.Row) = UCase(Format(UserForm_copy.TextBoxTER.Value))
        Tbmaio.ListColumns("valor bruto").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.TextBoxBRT.Text)
        Tbmaio.ListColumns("EXECUTIVO").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.TextBoxCONT.Text)
        Tbmaio.ListColumns("vencimento").Range.Rows(NovaLinha.Row) = UCase(Format(UserForm_copy.TextBoxVENC.Text))
        Tbmaio.ListColumns("obs").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.TextBoxOBS.Text)
        Tbmaio.ListColumns("V.liq.").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.TextBoxLQT.Text)
        Tbmaio.ListColumns("cnpj").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.TextBoxCNPJ.Text)
        
        ' limpa campos
        
        UserForm_copy.TextBoxCLI.Text = ""
        UserForm_copy.TextBoxNEG.Text = ""
        UserForm_copy.TextBoxPI.Text = ""
        UserForm_copy.TextBoxCONTR.Text = ""
        UserForm_copy.TextBoxTER.Text = ""
        UserForm_copy.TextBoxBRT.Text = ""
        UserForm_copy.TextBoxCONT.Text = ""
        UserForm_copy.TextBoxOBS.Text = ""
        UserForm_copy.TextBoxVENC.Text = ""
        UserForm_copy.TextBoxCN.Text = ""
        UserForm_copy.TextBoxLQT.Text = ""
        UserForm_copy.TextBoxCNPJ.Text = ""
        
End Sub
Sub INSERIR_po()
    Dim NovaLinha As Range
    Dim Tbmaio As Object
    
    If UserForm_copy.neg.Text = "15765" Then
    ab = "N"
    End If
    
    If UserForm_copy.neg.Text = "15766" Then
    ab = "E"
    End If
    
   'tabela = UserForm_copy.TextB_m.Value
   
    Set Tbmaio = ActiveSheet.ListObjects("POLITICOS")
    
        If UserForm_copy.d1.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("01" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d1.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        If UserForm_copy.d2.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("02" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d2.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
     
        If UserForm_copy.d3.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("03" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d3.Value
             Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        If UserForm_copy.d4.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("04" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d4.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
        If UserForm_copy.d5.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("05" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d2.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        If UserForm_copy.d6.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("06" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d6.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
        If UserForm_copy.d7.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("07" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d7.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        If UserForm_copy.d8.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("08" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d8.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
        If UserForm_copy.d9.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("09" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d9.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
       
        
        If UserForm_copy.d10.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("10" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d10.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
       If UserForm_copy.d11.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("11" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d11.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
       If UserForm_copy.d12.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("12" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d12.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
       If UserForm_copy.d13.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("13" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d13.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
       If UserForm_copy.d14.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("14" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d14.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
        If UserForm_copy.d15.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("15" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d15.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
        If UserForm_copy.d16.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("16" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d16.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
        If UserForm_copy.d17.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("17" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d17.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
        If UserForm_copy.d18.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("18" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d18.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
       If UserForm_copy.d19.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("19" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d19.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
       If UserForm_copy.d20.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("20" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d20.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
       If UserForm_copy.d21.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("21" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d21.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
        If UserForm_copy.d22.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("22" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d22.Value
             Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
        If UserForm_copy.d23.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("23" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d23.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
        If UserForm_copy.d24.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("24" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d24.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
       If UserForm_copy.d25.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("25" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d25.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        
        If UserForm_copy.d26.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("26" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d26.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
       If UserForm_copy.d27.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("27" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d27.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        If UserForm_copy.d28.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("28" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d28.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        If UserForm_copy.d29.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("29" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d29.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        If UserForm_copy.d30.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("30" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d30.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If
        
        If UserForm_copy.d31.Value <> "" Then
            Set NovaLinha = Tbmaio.ListRows.Add.Range
            Tbmaio.ListColumns("partido").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.partido.Text)
            Tbmaio.ListColumns("material").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.MATERIAL.Text)
            Tbmaio.ListColumns("n/e").Range.Rows(NovaLinha.Row) = ab
            Tbmaio.ListColumns("data").Range.Rows(NovaLinha.Row) = Format("31" & "/" & UserForm_copy.t_mes.Text & "/" & UserForm_copy.t_ano.Text, "DD/MM/yy")
            Tbmaio.ListColumns("contrato").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.contrato.Text)
            Tbmaio.ListColumns("neg").Range.Rows(NovaLinha.Row) = UCase(UserForm_copy.neg.Text)
            Tbmaio.ListColumns("qt").Range.Rows(NovaLinha.Row) = UserForm_copy.d31.Value
            Tbmaio.ListColumns("sec").Range.Rows(NovaLinha.Row) = UserForm_copy.sec.Value
        End If

        ' limpa campos

        UserForm_copy.d1.Text = ""
        UserForm_copy.d2.Text = ""
        UserForm_copy.d3.Text = ""
        UserForm_copy.d4.Text = ""
        UserForm_copy.d5.Text = ""
        UserForm_copy.d6.Text = ""
        UserForm_copy.d7.Text = ""
        UserForm_copy.d8.Text = ""
        UserForm_copy.partido.Text = ""
        UserForm_copy.MATERIAL.Text = ""
        UserForm_copy.ne.Text = ""
        UserForm_copy.data.Text = ""
    
End Sub
Sub testClipB()
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
 MsgBox CB.GetText
End Sub
Sub add_cli()
 Dim Msg, Style, Title, Help, Ctxt, Response, MyString
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
 MsgBox CB.GetText
 'cola = CB.GetText
'CLIENTE
    If UserForm_copy.TextBoxCLI.Text = "" Then
            UserForm_copy.TextBoxCLI.Value = UCase(CB.GetText)
     Else
        Msg = "DESEJA SUBSTITUIR O CLIENTE?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "TROCAR CLIENTE?"    ' Define title.
        Help = "DEMO.HLP"    ' Define Help file.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
            UserForm_copy.TextBoxCLI.Text = UCase(CB.GetText)
           ' MsgBox CB.GetText
        Else    ' User chose No.
            MyString = "No"    ' Perform some action.
        End If
     End If
  
End Sub
Sub add_neg()
 Dim Msg, Style, Title, Help, Ctxt, Response, MyString
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
MsgBox CB.GetText


'NEGOCIAÇÃO
    
    If UserForm_copy.TextBoxNEG.Text = "" Then
 
            UserForm_copy.TextBoxNEG.Text = CB.GetText

     Else
        Msg = "DESEJA SUBSTITUIR O NEGOCIAÇÃO?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "TROCAR NEGOCIAÇÃO?"    ' Define title.
        Help = "DEMO.HLP"    ' Define Help file.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
            UserForm_copy.TextBoxNEG.Text = UCase(CB.GetText)
           ' MsgBox CB.GetText
        Else    ' User chose No.
            MyString = "No"    ' Perform some action.
        End If
     End If
  
End Sub
Sub add_pi()
 Dim Msg, Style, Title, Help, Ctxt, Response, MyString
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
 MsgBox CB.GetText

'PI
    
    If UserForm_copy.TextBoxPI.Text = "" Then

            UserForm_copy.TextBoxPI.Text = UCase(CB.GetText)
        Else
     
        Msg = "DESEJA SUBSTITUIR O PI?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "TROCAR PI?"    ' Define title.
        Help = "DEMO.HLP"    ' Define Help file.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
            UserForm_copy.TextBoxPI.Text = UCase(CB.GetText)
           ' MsgBox CB.GetText
            Else    ' User chose No.
                MyString = "No"    ' Perform some action.
        End If
    
    End If
  
End Sub
Sub add_contr()
 Dim Msg, Style, Title, Help, Ctxt, Response, MyString
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
 MsgBox CB.GetText

 
'CONTRATO

    If UserForm_copy.TextBoxCONTR.Text = "" Then

            UserForm_copy.TextBoxCONTR.Text = UCase(CB.GetText)
 
     Else
        Msg = "DESEJA SUBSTITUIR O CONTRATO?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "TROCAR CONTRATO?"    ' Define title.
        Help = "DEMO.HLP"    ' Define Help file.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
            UserForm_copy.TextBoxCONTR.Text = UCase(CB.GetText)
           ' MsgBox CB.GetText
        Else    ' User chose No.
            MyString = "No"    ' Perform some action.
        End If
     End If
  
End Sub
Sub add_TERM()
 Dim Msg, Style, Title, Help, Ctxt, Response, MyString
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
 MsgBox CB.GetText


'TÉRMINO
    
    If UserForm_copy.TextBoxTER.Text = "" Then
 
            UserForm_copy.TextBoxTER.Text = UCase(CB.GetText)

     Else
        Msg = "DESEJA SUBSTITUIR A DATA FINAL?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "TROCAR DATA FINAL?"    ' Define title.
        Help = "DEMO.HLP"    ' Define Help file.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
            UserForm_copy.TextBoxTER.Text = UCase(CB.GetText)
           ' MsgBox CB.GetText
        Else    ' User chose No.
            MyString = "No"    ' Perform some action.
        End If
     End If
  
End Sub
Sub add_VB()
 Dim Msg, Style, Title, Help, Ctxt, Response, MyString
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
 MsgBox CB.GetText


'VALOR BRUTO

    If UserForm_copy.TextBoxBRT.Text = "" Then

            UserForm_copy.TextBoxBRT.Text = UCase(CB.GetText)

     Else
        Msg = "DESEJA SUBSTITUIR O VALOR BRUTO?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "TROCAR VALOR BRUTO?"    ' Define title.
        Help = "DEMO.HLP"    ' Define Help file.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
            UserForm_copy.TextBoxBRT.Text = UCase(CB.GetText)
           ' MsgBox CB.GetText
        Else    ' User chose No.
            MyString = "No"    ' Perform some action.
        End If
     End If
  
End Sub
Sub add_Vc()
 Dim Msg, Style, Title, Help, Ctxt, Response, MyString
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
 MsgBox CB.GetText


'VENCIMENTO

    If UserForm_copy.TextBoxVENC.Text = "" Then

            UserForm_copy.TextBoxVENC.Text = UCase(CB.GetText)

     Else
        Msg = "DESEJA SUBSTITUIR O VALOR BRUTO?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "TROCAR VENCIMENTO?"    ' Define title.
        Help = "DEMO.HLP"    ' Define Help file.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
            UserForm_copy.TextBoxVENC.Text = UCase(CB.GetText)
           ' MsgBox CB.GetText
        Else    ' User chose No.
            MyString = "No"    ' Perform some action.
        End If
     End If
  
End Sub
Sub add_CONT()
 Dim Msg, Style, Title, Help, Ctxt, Response, MyString
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
 MsgBox CB.GetText


'CONTATO
    
    If UserForm_copy.TextBoxCONT.Text = "" Then

            UserForm_copy.TextBoxCONT.Text = UCase(CB.GetText)
     Else
        Msg = "DESEJA SUBSTITUIR O CONTATO?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "TROCAR CONTATO?"    ' Define title.
        Help = "DEMO.HLP"    ' Define Help file.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
            UserForm_copy.TextBoxCONT.Text = UCase(CB.GetText)
           ' MsgBox CB.GetText
        Else    ' User chose No.
            MyString = "No"    ' Perform some action.
        End If
     End If
  
End Sub
Sub add_OBS()
 Dim Msg, Style, Title, Help, Ctxt, Response, MyString
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
 MsgBox CB.GetText


'OBSERVAÇÃO
    
    If UserForm_copy.TextBoxOBS.Text = "" Then

            UserForm_copy.TextBoxOBS.Text = UCase(CB.GetText)

     Else
        Msg = "DESEJA SUBSTITUIR A OBSERVAÇÃO?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "TROCAR OBSERVAÇÃO?"    ' Define title.
        Help = "DEMO.HLP"    ' Define Help file.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
            UserForm_copy.TextBoxOBS.Text = UCase(CB.GetText)
           ' MsgBox CB.GetText
        Else    ' User chose No.
            MyString = "No"    ' Perform some action.
        End If
     End If
  
End Sub
Sub add_CNPJ()
 Dim Msg, Style, Title, Help, Ctxt, Response, MyString
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
 MsgBox CB.GetText


'CONTATO
    
    If UserForm_copy.TextBoxCNPJ.Text = "" Then

            UserForm_copy.TextBoxCNPJ.Text = UCase(CB.GetText)
     Else
        Msg = "DESEJA SUBSTITUIR O CNPJ?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "TROCAR CNPJ?"    ' Define title.
        Help = "DEMO.HLP"    ' Define Help file.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
            UserForm_copy.TextBoxCNPJ.Text = UCase(CB.GetText)
           ' MsgBox CB.GetText
        Else    ' User chose No.
            MyString = "No"    ' Perform some action.
        End If
     End If
  
End Sub
Sub add_liq()
 Dim Msg, Style, Title, Help, Ctxt, Response, MyString
 Dim CB As Object
 Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 CB.GetFromClipboard
 MsgBox CB.GetText


'CONTATO
    
    If UserForm_copy.TextBoxLQT.Text = "" Then

            UserForm_copy.TextBoxLQT.Text = UCase(CB.GetText)
     Else
        Msg = "DESEJA SUBSTITUIR O VALOR LIQUIDO?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "TROCAR LIQUIDO?"    ' Define title.
        Help = "DEMO.HLP"    ' Define Help file.
        Ctxt = 1000    ' Define topic context.
                ' Display message.
        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
            UserForm_copy.TextBoxLQT.Text = UCase(CB.GetText)
           ' MsgBox CB.GetText
        Else    ' User chose No.
            MyString = "No"    ' Perform some action.
        End If
     End If
  
End Sub
