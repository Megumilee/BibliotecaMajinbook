VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Pesquisar 
   Caption         =   "Consultar livros"
   ClientHeight    =   5775
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   9855.001
   OleObjectBlob   =   "frm_Pesquisar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Pesquisar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnConsultar_Click()
Call FiltraLivros
End Sub

Private Sub btnHome_Click()
Unload frm_Pesquisar
frm_Menu.Show
End Sub

Private Sub UserForm_Initialize()
Call CriaCabecalho
End Sub
Sub CriaCabecalho()

    With lvLivros
        .View = lvwReport
        .FullRowSelect = True
        
        With .ColumnHeaders
            .Clear
            .Add , , "Titulo do livro", 60
            .Add , , "Autor(es)", 50, 2
            .Add , , "Editora", 90, 2
            .Add , , "Genero", 90, 2
            .Add , , "Numero do volume", 80, 2
            .Add , , "Local de aquisicao", 80, 2
            .Add , , "Numero da prateleira", 80, 2
            .Add , , "Status da leitura", 80, 2
            .Add , , "Anotacoes adicionais", 200, 2
        End With
        
    End With
End Sub
Sub FiltraLivros()

    Dim lvItem      As MSComctlLib.ListItem
    Dim Wplan       As Worksheet
    Dim lin         As Integer
    Dim lvLivros    As MSComctlLib.ListView
    Dim filtroAutor As String
    Dim filtroTitulo As String
    Dim valorAutor As String
    Dim valorTitulo As String

    Set Wplan = ThisWorkbook.Sheets("Cadastro_Livros")
    Set lvLivros = frm_Pesquisar.lvLivros

    filtroAutor = LCase(Trim(frm_Pesquisar.txtAutorPesq.Text))
    filtroTitulo = LCase(Trim(frm_Pesquisar.txtLivroPesq.Text))
    filtroEditora = LCase(Trim(frm_Pesquisar.txtEditora.Text))

    lin = 2

    lvLivros.ListItems.Clear

    With Wplan
        While .Cells(lin, 1).Value <> ""
        
   valorTitulo = LCase(Trim(CStr(.Cells(lin, "A").Value))) ' Título
   valorAutor = LCase(Trim(CStr(.Cells(lin, "B").Value)))  ' Autor
   valorEditora = LCase(Trim(CStr(.Cells(lin, "C").Value)))  ' Editora

If (filtroAutor <> "" And InStr(valorAutor, filtroAutor) > 0) Or _
   (filtroAutor = "" And _
       ((filtroTitulo = "" Or InStr(valorTitulo, filtroTitulo) > 0) And _
        (filtroEditora = "" Or InStr(valorEditora, filtroEditora) > 0))) Then


    Set lvItem = lvLivros.ListItems.Add(, , Format(.Cells(lin, "A").Value, "0000"))
    
    lvItem.ListSubItems.Add , , .Cells(lin, "B").Value
    lvItem.ListSubItems.Add , , .Cells(lin, "C").Value
    lvItem.ListSubItems.Add , , .Cells(lin, "D").Value
    lvItem.ListSubItems.Add , , .Cells(lin, "E").Value
    lvItem.ListSubItems.Add , , .Cells(lin, "F").Value
    lvItem.ListSubItems.Add , , .Cells(lin, "G").Value
    lvItem.ListSubItems.Add , , .Cells(lin, "H").Value
    lvItem.ListSubItems.Add , , .Cells(lin, "I").Value
    lvItem.ListSubItems.Add , , .Cells(lin, "J").Value
End If

        lin = lin + 1
        Wend
    End With

End Sub

