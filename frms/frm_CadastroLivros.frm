VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_CadastroLivros 
   Caption         =   "CADASTRO DE LIVROS"
   ClientHeight    =   7995
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   11475
   OleObjectBlob   =   "frm_CadastroLivros.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_CadastroLivros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCadastrar_Click()
    Dim plan As Worksheet
    Dim lastRow As Long
    Dim msgCad As String

    Set plan = ThisWorkbook.Sheets("Cadastro_Livros") '
    
    ' Encontrar a última linha preenchida na coluna A
    lastRow = plan.Cells(plan.Rows.Count, 1).End(xlUp).Row

    ' Adicionar uma nova linha após a última linha preenchida
    plan.Cells(lastRow + 1, 1).Value = Me.txtLivro.Value
    plan.Cells(lastRow + 1, 2).Value = Me.txtAutor.Value
    plan.Cells(lastRow + 1, 3).Value = Me.txtEditora.Value
    plan.Cells(lastRow + 1, 4).Value = Me.cbGenero.Value
    
    If checkUnico.Value = True Then
        Liv = "Livro único"
        plan.Cells(lastRow + 1, 5).Value = Liv
    Else
        plan.Cells(lastRow + 1, 5).Value = Me.cbVolume.Value
    End If
    
    plan.Cells(lastRow + 1, 6).Value = Me.cbLivraria.Value
    
    If checkNotApp.Value = True Then
        notAp = "Não aplicável/Livro digital"
        plan.Cells(lastRow + 1, 7).Value = notAp
    Else
        plan.Cells(lastRow + 1, 7).Value = Me.cbPrat.Value
    End If
    
    ' Preenche de acordo com a opcao selecionada na caixa de opcoes
    If lastRow < 1 Then lastRow = 1
    If OpNaoInic.Value = True Then
        Status = "Leitura não iniciada"
    ElseIf OpAnda = True Then
        Status = "Leitura em andamento"
    ElseIf OpConc = True Then
        Status = "Leitura concluída!"
    End If
    plan.Cells(lastRow + 1, 8).Value = Status
    plan.Cells(lastRow + 1, 9).Value = Me.txtNotes.Value
    

    msgCad = "LIVRO " & Me.txtLivro.Value & " CADASTRADO COM SUCESSO!"
    MsgBox msgCad, vbOKOnly + vbInformation, "LIVRO CADASTRADO!"
    
    ' Limpar os campos de entrada
    Me.txtLivro.Value = ""
    Me.txtAutor.Value = ""
    Me.txtEditora.Value = ""
    Me.txtNotes.Value = ""
    Me.cbGenero.Value = ""
    Me.cbLivraria.Value = ""
    Me.cbPrat.Value = ""
    Me.cbVolume.Value = ""
    Me.OpNaoInic.Value = False
    Me.OpAnda.Value = False
    Me.OpConc.Value = False
    
        
    ThisWorkbook.Save
    
End Sub

Private Sub btnHome_Click()
Unload frm_CadastroLivros
frm_Menu.Show
End Sub

Private Sub UserForm_Initialize()
    Dim Generos As Variant
    Dim Gen As Variant
    Dim Volumes As Variant
    Dim Vol As Variant
    Dim Prateleira As Variant
    Dim Prat As Variant
    Dim Livrarias As Variant
    Dim Liv As Variant
    
    ' Definindo a lista de gêneros
    Generos = Array("Ação", "Aventura", "Biografia", "True Crime", _
                    "Drama", "Fantasia", "Ficção Científica", _
                   "Horror", "Mistério", "Romance", _
                    "Suspense", "Terror", "Thriller", _
                    "Policial/Investigação", "Literatura Clássica", _
                    "Romance", "Contos", _
                    "Humor", "Científico", "Quadrinhos", "Mangás")
    cbGenero.Clear
    For Each Gen In Generos
        cbGenero.AddItem Gen
    Next Gen
    
    ' Definindo a lista de volumes
    Volumes = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10")
    
    cbVolume.Clear
    For Each Vol In Volumes
        cbVolume.AddItem Vol
    Next Vol
    
    
    ' Definindo a lista de prateleiras
    Prateleira = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10")
    
    cbPrat.Clear
    For Each Prat In Prateleira
        cbPrat.AddItem Prat
    Next Prat
    
    
     ' Definindo a lista de livrarias
    Livrarias = Array("JC SEBO", "AMAZON", "LIVRARIA (SS)", "COMIX", "Sebo Liberdade")
    
    cbLivraria.Clear
    For Each Liv In Livrarias
        cbLivraria.AddItem Liv
    Next Liv
    
End Sub

