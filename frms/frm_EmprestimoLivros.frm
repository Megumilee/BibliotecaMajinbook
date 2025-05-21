VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_EmprestimoLivros 
   Caption         =   "Cadastro de emprestimo de livros"
   ClientHeight    =   7110
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   10755
   OleObjectBlob   =   "frm_EmprestimoLivros.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_EmprestimoLivros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCadastrarEmp_Click()
    Dim plan As Worksheet
    Dim lastRow As Long
    Dim msgCad As String

    Set plan = ThisWorkbook.Sheets("Cadastro_Emprestimos") '
    
    ' Encontrar a última linha preenchida na coluna A
    lastRow = plan.Cells(plan.Rows.Count, 1).End(xlUp).Row

    ' Adicionar uma nova linha após a última linha preenchida
    plan.Cells(lastRow + 1, 1).Value = Me.cbLivroEmp.Value
    plan.Cells(lastRow + 1, 2).Value = Me.txtSolicitante.Value
    plan.Cells(lastRow + 1, 3).Value = Me.txtDtEmp.Value
    plan.Cells(lastRow + 1, 4).Value = Me.txtDtDevo.Value

    
    ' Preenche de acordo com a opcao selecionada na caixa de opcoes
    If lastRow < 1 Then lastRow = 1
    If OpDevo.Value = True Then
        Status = "Livro devolvido"
    ElseIf OpCLeitor = True Then
        Status = "Livro em posse do leitor solicitante"
    End If
    plan.Cells(lastRow + 1, 5).Value = Status
    plan.Cells(lastRow + 1, 6).Value = Me.txtNotes.Value
    

    msgCad = "LIVRO " & Me.cbLivroEmp.Value & " CADASTRADO COM SUCESSO NO CONTROLE DE EMPRÉSTIMOS!"
    MsgBox msgCad, vbOKOnly, "LIVRO CADASTRADO!"
    
    ' Limpar os campos de entrada
    Me.txtSolicitante.Value = ""
    Me.txtDtEmp.Value = ""
    Me.txtDtDevo.Value = ""
    Me.txtNotes.Value = ""
    Me.cbLivroEmp.Value = ""
    Me.OpDevo.Value = False
    Me.OpCLeitor.Value = False
    
    ThisWorkbook.Save


End Sub

Private Sub btnHome_Click()
Unload frm_EmprestimoLivros
frm_Menu.Show
End Sub


Private Sub UserForm_Initialize()
Call PreencherComboBox
End Sub
Sub PreencherComboBox()

    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim cbLivroEmp As MSForms.ComboBox
    Dim i As Integer


    Set ws = ThisWorkbook.Sheets("Cadastro_Livros") '
    Set cbLivroEmp = frm_EmprestimoLivros.cbLivroEmp '

    ' Defina o intervalo de dados
    Set rng = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

    ' Limpar a ComboBox existente
    cbLivroEmp.Clear

    ' Preencher ComboBox com nomes dos livros
    For Each cell In rng
        cbLivroEmp.AddItem cell.Value
    Next cell

End Sub
Private Sub txtDtEmp_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    txtDtEmp.MaxLength = 10 ' dd/mm/aaaa

    Select Case KeyAscii
        Case 8, 48 To 57 ' backspace e numéricos

            If Len(txtDtEmp.Text) = 2 Then
                txtDtEmp.Text = txtDtEmp.Text & "/"
            ElseIf Len(txtDtEmp.Text) = 5 Then
                txtDtEmp.Text = txtDtEmp.Text & "/"
            End If

        Case Else
            ' não faz nada
            KeyAscii = 0
    End Select

End Sub
Private Sub txtDtDevo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    txtDtDevo.MaxLength = 10 ' dd/mm/aaaa

    Select Case KeyAscii
        Case 8, 48 To 57 ' backspace e numéricos

            If Len(txtDtDevo.Text) = 2 Then
                txtDtDevo.Text = txtDtDevo.Text & "/"
            ElseIf Len(txtDtDevo.Text) = 5 Then
                txtDtDevo.Text = txtDtDevo.Text & "/"
            End If

        Case Else
            ' não faz nada
            KeyAscii = 0
    End Select

End Sub
