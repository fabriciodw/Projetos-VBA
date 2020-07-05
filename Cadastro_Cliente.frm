VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cadastro_Cliente 
   Caption         =   "Cadastro de cliente"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11265
   OleObjectBlob   =   "Cadastro_Cliente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cadastro_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variaveis globais
Dim ativo, nome As String
Dim linha As Integer
Dim ID As Boolean

'Verificar se já está cadastrado
Private Sub verificar_nome(nome)
        ID = False
        linha = 3
        While Plan7.Cells(linha, 1) <> "" 'Faça enquanto a linha coluna 01 for diferente de vazio
            If Plan7.Cells(linha, 3) = nome Then
                    ID = True
            End If
        linha = linha + 1
    Wend
End Sub
'limpar todos os campos
Private Sub limpar()
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
    Text_nome.SetFocus
End Sub

'Pegar a ultima linha em branco
Private Sub Carregar()
    Text_nome.SetFocus 'Seta o foco na caixa de texto
    Worksheets("Cliente").Select 'Seleciona a plan cliente
    linha = 1
    While Plan7.Cells(linha, 1) <> "" 'Faça enquanto a linha coluna 01 for diferente de vazio
        linha = linha + 1
    Wend
    Text_data.Locked = True
    Text_ID.Locked = True
    Text_ID.Value = linha - 1
    Text_data.Value = Date
    Option_sim.Value = True
End Sub


' Codigo para adicionar um cliente na planilha
Private Sub bt_salvar_Click()
    Plan7.Activate
    Plan7.Select
    nome = UCase(Text_nome.Text)
    Call verificar_nome(nome)
    If Text_nome = "" Then
        MsgBox "Campo nome é obrigatório!!", vbCritical, "Cadastro de Cliente"
        ElseIf ID = True Then
                MsgBox "Cliente ja cadastrado!!", vbCritical, "Cadastro de Cliente"
                Else
                    If Option_sim = True Then
                        ativo = "Ativo"
                        Else
                        ativo = "Inativo"
                    End If
                    
                    registro = linha 'Pegando a linha que irá entrar os dados
                    
                    Cells(registro, 1) = Text_ID.Value
                    Cells(registro, 2) = Text_data
                    Cells(registro, 3) = UCase(Text_nome.Text)
                    Cells(registro, 4) = Text_CNPJ.Value
                    Cells(registro, 5) = Text_CPF.Value
                    Cells(registro, 6) = Text_celular.Value
                    Cells(registro, 7) = Text_fixo.Value
                    Cells(registro, 8) = Text_endereco.Text
                    Cells(registro, 9) = Text_numero.Value
                    Cells(registro, 10) = Text_CEP.Value
                    Cells(registro, 11) = Combo_estado.Text
                    Cells(registro, 12) = Text_bairro.Text
                    Cells(registro, 13) = Text_complemento.Text
                    Cells(registro, 14) = Text_cidade.Text
                    Cells(registro, 15) = ativo
                    Cells(registro, 1).Select
                    
                    Call limpar
                    Call Carregar
        End If
End Sub

'Fecha o forme
Private Sub bt_voltar_Click()
    Unload Me
End Sub

'Clicar no botão editar no forme abre a lista de clientes
Private Sub btnEditarCliente_Click()
    Alterar_Cliente.Show
End Sub

'Ao inicializar o form
Private Sub UserForm_Initialize()
    'Preenchedo o combo_estado
    Me.Combo_estado.RowSource = "Planilha1!A2:A20"
    Call Carregar
End Sub

'Definir um local fixo do form na tela
Private Sub UserForm_Layout() ' Posição do forme na tela
   Me.Move 360, 90
End Sub
