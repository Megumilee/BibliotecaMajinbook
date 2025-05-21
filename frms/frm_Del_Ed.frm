VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Del_Ed 
   Caption         =   "EDITAR OU EXCLUIR LIVROS"
   ClientHeight    =   7380
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   13440
   OleObjectBlob   =   "frm_Del_Ed.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Del_Ed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnHome_Click()
Unload Me
frm_Menu.Show
End Sub
Private Sub cbLivros_Change()
Dim ws As Worksheet
Dim i As Long
Dim livroSel As String

livroSel = cbLivros.Value
Set ws = ThisWorkbook.Sheets("Cadastro_Livros")


For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

If ws.Cells(i, 1).Value = livroSel Then
    txtLivro.Value = ws.Cells(i, 1).Value
    txtAutor.Value = ws.Cells(i, 2).Value
    txtEditora.Value = ws.Cells(i, 3).Value
    txtGenero.Value = ws.Cells(i, 4).Value
    txtVolume.Value = ws.Cells(i, 5).Value
    txtLivraria.Value = ws.Cells(i, 6).Value
    txtPrat.Value = ws.Cells(i, 7).Value
    txtStatus.Value = ws.Cells(i, 8).Value
    txtNotes.Value = ws.Cells(i, 9).Value
    
    Exit For
End If

Next i

End Sub
Private Sub lblEditar_Click()
    Dim ws As Worksheet
    Dim resposta As VbMsgBoxResult
    Dim i As Long
    Dim idx As Long

    Set ws = ThisWorkbook.Sheets("Cadastro_Livros")

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
                ws.Cells(i, 2).Value = txtAutor.Value
                ws.Cells(i, 3).Value = txtEditora.Value
                ws.Cells(i, 4).Value = txtGenero.Value
                ws.Cells(i, 5).Value = txtVolume.Value
                ws.Cells(i, 6).Value = txtLivraria.Value
                ws.Cells(i, 7).Value = txtPrat.Value
                ws.Cells(i, 8).Value = txtStatus.Value
                ws.Cells(i, 9).Value = txtNotes.Value
            
                
                MsgBox "Livro atualizado com sucesso!"
                
                'Atualização do nome do livro (caso houver) no cb de seleção
                If idx >= 0 Then
                    cbLivros.RemoveItem idx
                    cbLivros.AddItem txtLivro.Value, idx ' Usa o valor atualizado do TextBox
                    cbLivros.ListIndex = idx ' Reposiciona no mesmo item
                End If
            
                'Limpa os campos
                cbLivros.Value = ""
                txtLivro.Value = ""
                txtAutor.Value = ""
                txtEditora.Value = ""
                txtGenero.Value = ""
                txtVolume.Value = ""
                txtLivraria.Value = ""
                txtPrat.Value = ""
                txtStatus.Value = ""
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

    Set ws = ThisWorkbook.Sheets("Cadastro_Livros")
    
    If cbLivros.Value = "" Then
        MsgBox "Nenhum livro selecionado para exclusão. Selecione o livro no combobox!", vbCritical, "Erro"
        Exit Sub
    End If

    If MsgBox("Deseja realmente excluir este livro?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then
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
            txtAutor.Value = ""
            txtEditora.Value = ""
            txtGenero.Value = ""
            txtVolume.Value = ""
            txtLivraria.Value = ""
            txtPrat.Value = ""
            txtStatus.Value = ""
            txtNotes.Value = ""
        
            Me.Hide
            Me.Show ' Recarrega o formulário
            Exit For
        End If
    Next i
    
    ThisWorkbook.Save
    
End Sub

Private Sub UserForm_Initialize()
Call PreencherComboBox
End Sub
Sub PreencherComboBox()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Cadastro_Livros")
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Limpa o ComboBox
    cbLivros.Clear
    
    ' Preenche o ComboBox com os nomes dos livros
    For i = 2 To ultimaLinha
        cbLivros.AddItem ws.Cells(i, 1).Value
    Next i
End Sub







