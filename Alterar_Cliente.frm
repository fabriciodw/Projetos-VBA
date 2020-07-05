VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Alterar_Cliente 
   Caption         =   "Alterar cliente"
   ClientHeight    =   6930
   ClientLeft      =   75
   ClientTop       =   390
   ClientWidth     =   11265
   OleObjectBlob   =   "Alterar_Cliente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Alterar_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim indice, linha, existe As Integer
'Variaveis globais do auto preencher
Private Const r As String = "C2:C100"
Private entrada As String
Dim parar As Boolean
Dim ativo, nome As String
Dim ID As Boolean

'Verificar se já está cadastrado
Private Sub verificar_nome(nome)
        ID = False
        L = 3
        While Plan7.Cells(L, 1) <> "" 'Faça enquanto a linha coluna 01 for diferente de vazio
            If Plan7.Cells(L, 3) = nome Then 'Logica para verificar se o nome existe
                If Plan7.Cells(L, 1) = existe Then ' logica para verificar se o nome que existe tem ID diferente do novo
                    ID = False
                    Else
                      ID = True
                    End If
            End If
            L = L + 1
        Wend
End Sub

' Alterar um cliente
Private Sub bt_alterar_Click()
    Dim retorno As Integer
    retorno = Text_ID
    linha = retorno + 1
    existe = retorno
    nome = UCase(Text_nome.Text)
    Call verificar_nome(nome)
    If Text_nome = "" Then 'Verifica se a box nome está vazia
        MsgBox "Campo nome é obrigatório!!", vbCritical, "Alterar Cliente"
        ElseIf ID = True Then
            MsgBox "Nome do cliente já existente!!", vbCritical, "Alterar Cliente"
            Else
                Worksheets("Cliente").Select ' Seleciona a plan cliente
                ativo = ""
                If Option_sim = True Then ' condição para colocar ativo ou inativo
                    ativo = "Ativo"
                    Else
                    ativo = "Inativo"
                End If
                'Recebe os valores nas text
                Cells(linha, 1) = Text_ID.Value
                Cells(linha, 2) = Text_data
                Cells(linha, 3) = UCase(Text_nome.Text)
                Cells(linha, 4) = Text_CNPJ.Value
                Cells(linha, 5) = Text_CPF.Value
                Cells(linha, 6) = Text_celular.Value
                Cells(linha, 7) = Text_fixo.Value
                Cells(linha, 8) = Text_endereco.Text
                Cells(linha, 9) = Text_numero.Value
                Cells(linha, 10) = Text_CEP.Value
                Cells(linha, 11) = Combo_estado.Text
                Cells(linha, 12) = Text_bairro.Text
                Cells(linha, 13) = Text_complemento.Text
                Cells(linha, 14) = Text_cidade.Text
                Cells(linha, 15) = ativo
                Cells(linha, 16) = Date
                
                'Chama a função limpar
                Call bt_limpar_Click
        End If
End Sub

'limpar todos os campos
Private Sub bt_limpar_Click()
    Text_ID = ""
    Text_data = ""
    Text_nome = ""
    Text_nome = ""
    Text_CNPJ = ""
    Text_CPF = ""
    Text_celular = ""
    Text_fixo = ""
    Text_endereco = ""
    Text_numero = ""
    Text_CEP = ""
    Combo_estado = ""
    Text_bairro = ""
    Text_complemento = ""
    Text_cidade = ""
    Option_sim = True
    Option_nao = False
    Frame2.Enabled = True
    alterar.Enabled = False
    Call btn_CancelaPesq_Click
End Sub

'Sair do form
Private Sub bt_voltar_Click()
    Unload Me
End Sub

'Cancela a pequisa
Private Sub btn_CancelaPesq_Click()
        TextPesquisa = ""
        TextPesquisa.SetFocus
        btnListar.Locked = False
        btnPesquisar.Locked = True
End Sub

'botão listar do form
Private Sub btnListar_Click()
        alterar.Enabled = True
        Pesquisa_Cliente.Show
End Sub

'sai do form
Private Sub btnSair2_Click()
        Unload Me
End Sub

Private Sub Text_CNPJ_Change()
        If Not IsNumeric(Text_CNPJ.Text) Then Text_CNPJ.Text = Empty
End Sub

'Inicio do codigo para fazer pesquisa dentro do text box, ele retorno o nome se existir
Private Sub textPesquisa_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) '
    If (KeyCode = vbKeyBack) Or (KeyCode = vbKeyDelete) Then
        parar = True
    Else
        parar = False
    End If
End Sub

Private Sub btnPesquisar_Click()
    Plan7.Activate
    With Plan7.Range("A:C")
        Set c = .Find(TextPesquisa.Value, LookIn:=xlValues, LOOKAT:=xlWhole)
        If Not c Is Nothing Then
            indice = c.Offset(0, -2)
            Call Carregar
            Frame2.Enabled = False
            alterar.Enabled = True
            
        End If
        If c Is Nothing Then
            MsgBox ("Nome não encontrado!!!"), vbOKOnly, ("Seu Aplicativo Pesquisando Dados")
            Call btn_CancelaPesq_Click
        End If
    End With
           
End Sub

Private Sub textPesquisa_Change()
    Dim sWord As String
    If parar Then
        parar = False
    Else
        entrada = Left(Me.TextPesquisa, Me.TextPesquisa.SelStart)
        sWord = GetFirstCloserWord(entrada)
        If sWord & "" <> "" Then
            parar = True
            Me.TextPesquisa.Text = sWord
            Me.TextPesquisa.SelStart = Len(entrada)
            Me.TextPesquisa.SelLength = 999999 'Tamanho do campo
        End If
    End If
    btnListar.Locked = True
    btnPesquisar.Locked = False
End Sub
 
Private Function GetFirstCloserWord(ByVal Word As String) As String
    Plan7.Select
    Plan7.Activate
    Dim c As Range
    For Each c In ActiveSheet.Range(r).Cells
    If LCase(c.Value) Like LCase(Word & "*") Then
            GetFirstCloserWord = c.Value
            Exit Function
        End If
    Next c
    Set c = Nothing
 
End Function

''MUDAR ESSA LOGICA JÀ TEM O INDICE PARA USAR EM VEZ DO FOR
'Carrega o form pelo modelo text box
Private Sub Carregar()
        Plan7.Activate
        linha = 2
        For coluna = 3 To 3
         While Plan7.Cells(linha, 3) <> ""
            xCel = Plan7.Cells(linha, 1)
            If xCel = indice Then
                 
                  Text_nome = Plan7.Cells(linha, 3)
                  Text_data = Date
                  Text_ID = Plan7.Cells(linha, 1)
                  Text_CNPJ = Plan7.Cells(linha, 4)
                  Text_CPF = Plan7.Cells(linha, 5)
                  Text_celular = Plan7.Cells(linha, 6)
                  Text_fixo = Plan7.Cells(linha, 7)
                  Text_endereco = Plan7.Cells(linha, 8)
                  Text_numero = Plan7.Cells(linha, 9)
                  Text_CEP = Plan7.Cells(linha, 10)
                  Text_bairro = Plan7.Cells(linha, 12)
                  Text_complemento = Plan7.Cells(linha, 13)
                  Text_cidade = Plan7.Cells(linha, 14)
                  Combo_estado = Plan7.Cells(linha, 11)
                                  
                  If Plan7.Cells(linha, 15) = "Ativo" Then
                        Option_sim = True
                    Else
                        Option_nao = True
                  End If
                                    
                  Exit For
            End If
              linha = linha + 1
        Wend
    Next coluna
    Text_ID.Locked = True
    Text_data.Locked = True

End Sub

'iniciando o for
Private Sub UserForm_Initialize()
    btnPesquisar.Locked = True
    alterar.Enabled = False
    Me.Combo_estado.RowSource = "Planilha1!A2:A20"
End Sub

' Posição do forme na tela
Private Sub UserForm_Layout()
   Me.Move 360, 92
End Sub
