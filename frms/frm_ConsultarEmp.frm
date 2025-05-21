VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ConsultarEmp 
   Caption         =   "Consultar lista de livros emprestados"
   ClientHeight    =   6090
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   9930.001
   OleObjectBlob   =   "frm_ConsultarEmp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_ConsultarEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnConsultar_Click()
Call FiltraLivros
End Sub

Private Sub btnHome_Click()
Unload frm_ConsultarEmp
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
            .Add , , "Nome do leitor", 50, 2
            .Add , , "Data do emprestimo", 90, 2
            .Add , , "Data da devolucao", 90, 2
            .Add , , "Status do emprestimo", 80, 2
            .Add , , "Observacoes", 80, 2
        End With
        
    End With
End Sub
Sub FiltraLivros()

    Dim lvItem      As MSComctlLib.ListItem
    Dim Wplan       As Worksheet
    Dim lin         As Integer
    Dim lvLivros    As MSComctlLib.ListView
    'Dim filtroAutor As String
    Dim filtroTitulo As String
    Dim filtroLeitor As String
    'Dim valorAutor As String
    Dim valorTitulo As String
    Dim valorLeitor As String

    Set Wplan = ThisWorkbook.Sheets("Cadastro_Emprestimos")
    Set lvLivros = frm_ConsultarEmp.lvLivros

    'filtroAutor = LCase(Trim(frm_ConsultarEmp.txtAutorPesq.Text))
    filtroTitulo = LCase(Trim(frm_ConsultarEmp.txtLivroPesq.Text))
    filtroLeitor = LCase(Trim(frm_ConsultarEmp.txtLeitor.Text))

    lin = 2

    lvLivros.ListItems.Clear

    With Wplan
        While .Cells(lin, 1).Value <> ""
        
   valorTitulo = LCase(Trim(CStr(.Cells(lin, "A").Value))) ' Título
   valorLeitor = LCase(Trim(CStr(.Cells(lin, "B").Value)))  ' Leitor
   

If (filtroTitulo <> "" And InStr(valorTitulo, filtroTitulo) > 0) Or _
   (filtroTitulo = "" And _
       ((filtroLeitor = "" Or InStr(valorLeitor, filtroLeitor) > 0))) Then


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


