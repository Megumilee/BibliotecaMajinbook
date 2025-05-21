VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Menu 
   Caption         =   "Menu - Biblioteca"
   ClientHeight    =   8490.001
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11475
   OleObjectBlob   =   "frm_Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl_cadastrar_Click()
frm_Menu.Hide
frm_CadastroLivros.Show
End Sub

Private Sub lbl_consultar_Click()

 Dim resposta As VbMsgBoxResult

    resposta = MsgBox("Qual lista deseja consultar?" & vbCrLf & vbCrLf & _
                      "Clique 'Sim' para Livros Gerais." & vbCrLf & _
                      "Clique 'Não' para Livros Emprestados." & vbCrLf & _
                      "Clique 'Cancelar' para sair.", _
                     vbYesNoCancel + vbQuestion, "Consulta de listas")

    Select Case resposta
        Case vbYes
            frm_Menu.Hide
            frm_Pesquisar.Show ' Livros Gerais
        Case vbNo
            frm_Menu.Hide
            frm_ConsultarEmp.Show ' Livros Emprestados
        Case vbCancel
            MsgBox "Consulta cancelada", vbInformation
    End Select
End Sub

Private Sub lbl_editar_Click()
Dim resposta As VbMsgBoxResult

    resposta = MsgBox("Qual lista deseja acessar?" & vbCrLf & vbCrLf & _
                      "Clique 'Sim' para Livros Gerais." & vbCrLf & _
                      "Clique 'Não' para Livros Emprestados." & vbCrLf & _
                      "Clique 'Cancelar' para sair.", _
                      vbYesNoCancel + vbQuestion, "Exportar Lista")

Select Case resposta
    Case vbYes
        frm_Menu.Hide
        frm_Del_Ed.Show
    Case vbNo
        frm_Menu.Hide
        frm_Del_Ed_Emp.Show
    Case vbCancel
        MsgBox "Operação cancelada!", vbInformation
End Select

End Sub

Private Sub lbl_emprestimo_Click()
frm_Menu.Hide
frm_EmprestimoLivros.Show
End Sub
Private Sub lbl_pdf_Click()

    Dim resposta As VbMsgBoxResult

    resposta = MsgBox("Qual lista deseja exportar?" & vbCrLf & vbCrLf & _
                      "Clique 'Sim' para Livros Gerais." & vbCrLf & _
                      "Clique 'Não' para Livros Emprestados." & vbCrLf & _
                      "Clique 'Cancelar' para sair.", _
                      vbYesNoCancel + vbQuestion, "Exportar Lista")

    Select Case resposta
        Case vbYes
            Call ExportarLivrosPDF ' Livros Gerais
        Case vbNo
            Call ExportarLivrosEmpPDF ' Livros Emprestados
        Case vbCancel
            MsgBox "Exportação cancelada.", vbInformation
    End Select

End Sub
Sub ExportarLivrosPDF()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim printRange As Range
    Dim colLetter As String
    Dim caminhoSalvar As String
    Dim nomeArquivo As String
    
    Set ws = ThisWorkbook.Sheets("Cadastro_Livros")
    
    ' Encontrar última linha e última coluna preenchida
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Converter número da coluna para letra
    colLetter = Split(ws.Cells(1, lastCol).Address(True, False), "$")(0)
    
    ' Definir área de impressão para células preenchidas
    Set printRange = ws.Range("A1:" & colLetter & lastRow)
    ws.PageSetup.PrintArea = printRange.Address
    
    ' Configurar página para paisagem e ajustar para caber largura
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    ' Caminho e nome do arquivo PDF
    caminhoSalvar = ThisWorkbook.Path & "\"
    nomeArquivo = "Cadastro_Livros_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    
    ' Exportar PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=caminhoSalvar & nomeArquivo, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        
    MsgBox "Planilha de lista de livros exportada em PDF com sucesso!"
End Sub

Sub ExportarLivrosEmpPDF()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim printRange As Range
    Dim colLetter As String
    Dim caminhoSalvar As String
    Dim nomeArquivo As String
    
    Set ws = ThisWorkbook.Sheets("Cadastro_Emprestimos")
    
    ' Encontrar última linha e última coluna preenchida
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Converter número da coluna para letra
    colLetter = Split(ws.Cells(1, lastCol).Address(True, False), "$")(0)
    
    ' Definir área de impressão para células preenchidas
    Set printRange = ws.Range("A1:" & colLetter & lastRow)
    ws.PageSetup.PrintArea = printRange.Address
    
    ' Configurar página para paisagem e ajustar para caber largura
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    ' Caminho e nome do arquivo PDF
    caminhoSalvar = ThisWorkbook.Path & "\"
    nomeArquivo = "Cadastro_Emprestimos_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    
    ' Exportar PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=caminhoSalvar & nomeArquivo, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        
    MsgBox "Planilha de emprestimos exportada em PDF com sucesso!"
End Sub
