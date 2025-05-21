VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Del_Ed_Emp 
   Caption         =   "Editar/Exluir itens livros emprestados"
   ClientHeight    =   7125
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   11850
   OleObjectBlob   =   "frm_Del_Ed_Emp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Del_Ed_Emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
Call PreencherComboBox
End Sub
Private Sub btnHome_Click()
Unload frm_Del_Ed_Emp
frm_Menu.Show
End Sub
Private Sub cbLivros_Change()
Dim ws As Worksheet
Dim i As Long
Dim livroSel As String

livroSel = cbLivros.Value
Set ws = ThisWorkbook.Sheets("Cadastro_Emprestimos")

For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

If ws.Cells(i, 1).Value = livroSel Then
    txtLivro.Value = ws.Cells(i, 1).Value
    txtLeitor.Value = ws.Cells(i, 2).Value
    txtDtEmp.Value = ws.Cells(i, 3).Value
    txtDtDevo.Value = ws.Cells(i, 4).Value
    txtStatusEmp.Value = ws.Cells(i, 5).Value
    txtNotes.Value = ws.Cells(i, 6).Value
    
    Exit For
End If

Next i

End Sub
Private Sub lblEditar_Click()
    Dim ws As Worksheet
    Dim resposta As VbMsgBoxResult
    Dim i As Long
    Dim idx As Long

    Set ws = ThisWorkbook.Sheets("Cadastro_Emprestimos")

    resposta = MsgBox("Deseja realmente editar o registro?", vbYesNo + vbQuestion, "Edição de livro")
    
    If cbLivros.Value = "" Then
        MsgBox "Nenhum livro selecionado para edição. Selecione o livro no combobox!", vbCritical, "Error"
        Exit Sub
    End If
    
    If resposta = vbYes Then
        For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If ws.Cells(i, 1).Value = cbLivros.Value Then
                ' Atualiza os dados da linha
                ws.Cells(i, 1).Value = txtLivro.Value
                ws.Cells(i, 2).Value = txtLeitor.Value
                ws.Cells(i, 3).Value = txtDtEmp.Value
                ws.Cells(i, 4).Value = txtDtDevo.Value
                ws.Cells(i, 5).Value = txtStatusEmp.Value
                ws.Cells(i, 6).Value = txtNotes.Value
               
            
                MsgBox "Informações atualizadas com sucesso!"
                
                'Atualização do nome do livro (caso houver) no cb de seleção
                If idx >= 0 Then
                    cbLivros.RemoveItem idx
                    cbLivros.AddItem txtLivro.Value, idx ' Usa o valor atualizado do TextBox
                    cbLivros.ListIndex = idx ' Reposiciona no mesmo item
                End If
            
                'Limpa os campos
                cbLivros.Value = ""
                txtLivro.Value = ""
                txtLeitor.Value = ""
                txtDtEmp.Value = ""
                txtDtDevo.Value = ""
                txtStatusEmp.Value = ""
                txtNotes.Value = ""
                
                Exit For
            End If
        Next i
    End If
    
    ThisWorkbook.Save

End Sub

Private Sub lblExcluir_Click()
    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Cadastro_Emprestimos")
    
    If cbLivros.Value = "" Then
        MsgBox "Nenhum livro selecionado para exclusão. Selecione o livro no combobox!", vbCritical, "Erro"
        Exit Sub
    End If

    If MsgBox("Deseja realmente excluir este registro?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then
        Exit Sub
    End If

  
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).Value = cbLivros.Value Then
            ws.Rows(i).Delete
            
            'Remove o item do cb de seleção
            cbLivros.RemoveItem cbLivros.ListIndex

            MsgBox "Livro excluído com sucesso!"
            
            ' Limpa os campos
            cbLivros.Value = ""
            txtLivro.Value = ""
            txtLeitor.Value = ""
            txtDtEmp.Value = ""
            txtDtDevo.Value = ""
            txtStatusEmp.Value = ""
            txtNotes.Value = ""
        
            Me.Hide
            Me.Show ' Recarrega o formulário
            Exit For
        End If
    Next i
    
ThisWorkbook.Save

End Sub
Sub PreencherComboBox()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Cadastro_Emprestimos")
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Limpa o ComboBox
    cbLivros.Clear
    
    ' Preenche o ComboBox com os nomes dos livros
    For i = 2 To ultimaLinha
        cbLivros.AddItem ws.Cells(i, 1).Value
    Next i
End Sub
