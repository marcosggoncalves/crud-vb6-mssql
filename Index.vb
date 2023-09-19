Dim Con As New ADODB.Connection
Dim Rec As New ADODB.Recordset


Private Sub Clear()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    
    Deletar.Visible = True
    
    ListView1.ListItems.Clear
End Sub

Private Sub Listagem()
    Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=gestfro;uid=sa;pwd=Lopes9951;"
        
    Rec.Open "Select * from veiculo ", Con, adOpenStatic, adLockReadOnly
    
    Do Until Rec.EOF
       Set lastItems = ListView1.ListItems.Add(, , Rec!id)
       lastItems.SubItems(1) = Rec!placa
       lastItems.SubItems(2) = Rec!descricao
       lastItems.SubItems(3) = Rec!modelo
       lastItems.SubItems(4) = Rec!marca
       lastItems.SubItems(5) = Rec!quantidade_lugares
       Rec.MoveNext
    Loop
    
    Rec.Close
    Con.Close
End Sub


Private Sub Form_Load()
    Call Listagem
End Sub

Private Sub Pesquisar_Click()
    If Text1.Text = "" Then
         MsgBox "Digite a Placa do Veiculo para realizar pesquisar!", vbOKCancel, "Erro!"
    Else
        Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=gestfro;uid=sa;pwd=Lopes9951;"
        
        Rec.Open "Select * from veiculo where placa='" & Text1.Text & "'", Con, adOpenStatic, adLockReadOnly
        
        If Rec.RecordCount > 0 Then
            Text6.Text = Rec.Fields!id
            Text2.Text = Rec.Fields!modelo
            Text3.Text = Rec.Fields!marca
            Text4.Text = Rec.Fields!quantidade_lugares
            Text5.Text = Rec.Fields!descricao
            
            Deletar.Visible = True
        Else
            MsgBox "Cadastro não encontrado!", vbOKCancel, "Erro!"
            
            If Text2.Text <> "" Then
                Call Clear
                Call Listagem
            End If
        End If
            
        Rec.Close
        Con.Close
    End If
End Sub

Private Sub Deletar_Click()
    If Text1.Text = "" Then
         MsgBox "Selecione um veiculo para realizar a exclusão!", vbOKCancel, "Erro!"
    Else
        Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=gestfro;uid=sa;pwd=Lopes9951;"
     
        Rec.Open "SELECT * FROM veiculo WHERE id = " & Text6.Text & "", Con, adOpenKeyset, adLockOptimistic
        
        If Not Rec.EOF Then
            MsgBox "Cadastro excluido com  sucesso!", vbOKCancel, "Realizado!"
            
            Rec.Delete
            Rec.Close
            Con.Close
            
            Call Clear
            Call Listagem
        End If
    End If
End Sub

Private Sub Gravar_Click()
    Dim strsql As String

    If Text1.Text = "" Then
        MsgBox "Digite a placa do Veiculo!", vbOKCancel, "Erro!"
    ElseIf Text2.Text = "" Then
        MsgBox "Digite o modelo do veiculo!", vbOKCancel, "Erro!"
    ElseIf Text3.Text = "" Then
        MsgBox "Digite a marca do veiculo!", vbOKCancel, "Erro!"
    ElseIf Text4.Text = "" Then
        MsgBox "Digite a quantidade de lugares do veiculo!", vbOKCancel, "Erro!"
    ElseIf Text5.Text = "" Then
        MsgBox "Digite uma descrição do veiculo!", vbOKCancel, "Erro!"
    Else
        
        Con.Open "Driver=SQL Server;Server=.\SQLEXPRESS;Database=gestfro;uid=sa;pwd=Lopes9951;"
        
        If Text6 <> "" Then
            strsql = "Update dbo.veiculo set placa = '" & Text1.Text & "', modelo = '" & Text2.Text & "', marca = '" & Text3.Text & "', descricao = '" & Text4.Text & "', quantidade_lugares = '" & Text5.Text & "' where id = '" & Text6.Text & "'"
        
            MsgBox "Cadastro alterado com  sucesso!", vbOKCancel, "Realizado!"
        Else
            strsql = "INSERT INTO dbo.veiculo(placa,modelo,marca,quantidade_lugares, descricao)VALUES('" & Text1.Text & "', '" & Text2.Text & "','" & Text3.Text & "', '" & Text4.Text & "', '" & Text5.Text & "')"
        
            MsgBox "Cadastro registrado com  sucesso!", vbOKCancel, "Realizado!"
        End If
        
        
        Con.BeginTrans
        Con.Execute strsql
        Con.CommitTrans
        Con.Close
        
        Call Clear
        Call Listagem
    End If
End Sub
