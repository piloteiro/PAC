Attribute VB_Name = "PAC"
'Declara a função que lê números inteiros do arquivo ini do PAC.
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'Declara a função que escreve strings no arquivo ini do PAC.
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Declara a função que lê strings do arquivo ini do PAC.
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


'Avisa aos outros procedimentos que o form está sendo carregado.
Public EstouCarregando As Integer

'Identifica qual é o sexo do Ego.
Public SexoEgo As Integer

'Identifica qual é o estado civil do Ego.
Public CivilEgo As Integer

'Recebe o valor lido do PAC.ini, onde 0=Português e 2300=Inglês.
Public Lingua As Integer

'Indica que o DBcp_Ego foi atualizado. 0=Não atualizado e 1=Atualizado
Public EgoAtualizado As Integer

'Indica se a tabela EGO do DBCP está vazia. 0=vazia e 1=tem algum dados.
Public BancoCP_EGO As Integer

'Indica se a tabela CASAIS do DBCP está vazia. 0=vazia e 1=tem algum dados.
Public BancoCP_CASAIS As Integer

Public CorrenteEgo As Integer

'Esta variável tem por finalidade guarda a ligação entre o ego inicial da pesquisa _
 e o ego pai dos filhos que porventura são encontrados. _
 Tree_filhos(Ego_Inicial, Ego_Pai, Quantidade, camada) Quando a quantidade=0, _
 indicará quantos TB estão incluidos da lista de Ego_Inicial
Public Tree_Filhos() As String

'Idicará qual a camada, ou geração, está sendo considerada na busca pelos TB=Termo Básico
Public Camada As Integer

'Arvore(Índice do ego_selecionado, quantidade de filho quando o indice for zero _
 exemplo= arvore(41,0,0)=2 indica que o ego 41 tem dois TB na lista como segue: _
 arvore(41,1,0)=34 e arvore(41,2,0)=25, o último índice define se o dado é _
 o ID ou nome do  ==> arvore(41,2,1)="Mané" ou arvore(41,2,0)=25)
Public Arvore() As String




Public Sub Idioma(ConstIdioma As Integer) 'ConstIdioma recebe o valor numérico que reflete o idioma da interface.
'Esta rotina ajusta o idioma usado na interface.

'Declara as variaveis para este Sub.
    Dim ID_RES As Variant 'Receberá os valores de um array que reflete os índices do Pac1.res.
    Dim ID_Control As Integer 'Receberá os valores de índice dos controles em array.

'Forms
    PAC1.Caption = LoadResString(10 + ConstIdioma)
    
'Labels
    ID_RES = Array(72, 42, 33, 61, 36, 57, 37, 51, 54, 55, 53, 45, 50, 38, 56, 42, 33, 61, 36, 57, 58, 52, 46, 47, 37, 51, 64, 54, 55, 72, 50, 38, 67, 66, 41, 39, 0, 40, 0, 66, 41, 43, 69, 42, 33, 61, 36, 57, 69, 44, 63, 59, 30, 60, 69, 68, 0, 69, 71, 42, 61, 69, 71, 35, 42, 48, 59, 65, 70, 34, 49, 62, 56)
    For ID_Control = 0 To 72
        If ID_Control <> 33 And ID_Control <> 36 And ID_Control <> 38 And ID_Control <> 56 Then
            PAC1.Label(ID_Control).Caption = LoadResString(ID_RES(ID_Control) + ConstIdioma)
        End If
    Next ID_Control
      
'SSOption boxes
    ID_RES = Array(98, 94, 99, 100, 102, 90, 104, 101, 103, 99, 100, 96, 92, 95, 97, 96, 92, 95, 97)
    For ID_Control = 1 To 19
        If ID_Control <> 10 And ID_Control <> 11 Then 'Os SSOptions 10 e 11 foram substituidos por SSCheck
            PAC1.SSOption(ID_Control).Caption = LoadResString(ID_RES(ID_Control - 1) + ConstIdioma)
        End If
    Next ID_Control
    
'SSCheck boxes
    ID_RES = Array(98, 94, 53, 45, 102, 90, 104, 101, 103, 99, 100) '91, 93,
    For ID_Control = 1 To 11
            PAC1.SSCheck(ID_Control).Caption = LoadResString(ID_RES(ID_Control - 1) + ConstIdioma)
    Next ID_Control
    
'Check boxes
    PAC1.Check(1).Caption = LoadResString(107 + ConstIdioma)
    PAC1.Check(2).Caption = LoadResString(105 + ConstIdioma)
    PAC1.Check(3).Caption = LoadResString(106 + ConstIdioma)
    PAC1.Nome_Nac(0).Caption = LoadResString(108 + ConstIdioma)
    PAC1.Nome_Nac(1).Caption = LoadResString(108 + ConstIdioma)
    PAC1.Nome_Nac(2).Caption = LoadResString(108 + ConstIdioma)

'SSCommand buttons
    ID_RES = Array(131, 136, 130, 137, 138, 133, 130, 133, 137, 134, 137, 141, 137, 137, 131, 130, 137, 130, 139, 137, 137, 132, 135, 131, 136, 137, 132, 135, 140, 136, 137, 132, 131, 136)
    For ID_Control = 0 To 33
            PAC1.SSCommand(ID_Control).Caption = LoadResString(ID_RES(ID_Control) + ConstIdioma)
    Next ID_Control
    
'Command Button
        PAC1.Command(1).Caption = LoadResString(142 + ConstIdioma)
  
'Frames
    ID_RES = Array(161, 164, 0, 172, 169, 165, 161, 171, 0, 172, 170, 165, 166, 173, 174, 173, 174, 161, 163, 160, 162, 175, 0, 161, 167, 168)
    For ID_Control = 1 To 26
        If ID_Control <> 3 And ID_Control <> 9 And ID_Control <> 23 Then
            PAC1.SSFrame(ID_Control).Caption = LoadResString(ID_RES(ID_Control - 1) + ConstIdioma)
        End If
    Next ID_Control
   
'Tabs
    PAC1.SSTab(1).TabCaption(0) = LoadResString(190 + ConstIdioma)
    PAC1.SSTab(1).TabCaption(1) = LoadResString(202 + ConstIdioma)
    PAC1.SSTab(1).TabCaption(2) = LoadResString(196 + ConstIdioma)
    PAC1.SSTab(1).TabCaption(3) = LoadResString(191 + ConstIdioma)
    PAC1.SSTab(1).TabCaption(4) = LoadResString(194 + ConstIdioma)
    PAC1.SSTab(1).TabCaption(5) = LoadResString(199 + ConstIdioma)
    PAC1.SSTab(2).TabCaption(0) = LoadResString(192 + ConstIdioma)
    PAC1.SSTab(2).TabCaption(1) = LoadResString(193 + ConstIdioma)
    PAC1.SSTab(2).TabCaption(2) = LoadResString(198 + ConstIdioma)
    PAC1.SSTab(3).TabCaption(0) = LoadResString(201 + ConstIdioma)
    PAC1.SSTab(3).TabCaption(1) = LoadResString(195 + ConstIdioma)
    PAC1.SSTab(4).TabCaption(0) = LoadResString(200 + ConstIdioma)
    PAC1.SSTab(4).TabCaption(1) = LoadResString(197 + ConstIdioma)
    PAC1.SSTab(4).TabCaption(2) = LoadResString(195 + ConstIdioma)

'Combo
    PAC1.Combo(2).AddItem LoadResString(262 + ConstIdioma)
    PAC1.Combo(2).AddItem LoadResString(263 + ConstIdioma)
    PAC1.Combo(4).AddItem LoadResString(262 + ConstIdioma)
    PAC1.Combo(4).AddItem LoadResString(263 + ConstIdioma)
    PAC1.Combo(7).AddItem LoadResString(260 + ConstIdioma)
    PAC1.Combo(7).AddItem LoadResString(261 + ConstIdioma)
    'Novo_Termo.Combo(0).AddItem LoadResString(262 + ConstIdioma)
    'Novo_Termo.Combo(0).AddItem LoadResString(263 + ConstIdioma)


'DBgrid
    PAC1.DBGrid(1).Columns(0).Caption = LoadResString(223 + ConstIdioma)
    PAC1.DBGrid(1).Columns(1).Caption = LoadResString(224 + ConstIdioma)
    PAC1.DBGrid(1).Columns(2).Caption = LoadResString(56 + ConstIdioma)

End Sub

Public Sub CentraForm(Formulario As Form)
'Esta rotina centraliza qualquer formulário que é passado para ela como "Form".

    Formulario.Move (Screen.Width \ 2) - (Formulario.Width \ 2), (Screen.Height \ 2) - (Formulario.Height \ 2)

End Sub


Public Sub Sexo_Civil()
'Esta rotina seleciona qual a opção deve ser ligada para refletir os dados sobre o sexo e o estado civil do corrente Ego.
    
    Dim Sexo As Integer 'Recebe o valor que identifica o sexo do corrente Ego.
    Dim Civil As Integer 'Recebe o valor que identifica o estado civil do corrente Ego.

'Sexo do Ego.
    Sexo = PAC1.DBcp_Ego.Recordset("sexo")
    Select Case Sexo
        Case 0 'Masculino
            PAC1.SSOption(1).Value = True
        Case 1 'Feminino
            PAC1.SSOption(2).Value = True
    End Select
    
'Estado Civil do Ego.
    Civil = PAC1.DBcp_Ego.Recordset("civil")
    'SSOption(5)=solteiro'2'
    'SSOption(6)=casado'3'
    'SSOption(7)=viúvo'4'
    'SSOption(8)=separado'5'
    'SSOption(9)=união irregular'6'
    For Indice = 5 To 9 'Os valores de 5 a 9 representam qual a opção do estado civil está selecionado.
        'Indice - 3 é o valor do estado civil no BD.
        If Civil = Indice - 3 Then PAC1.SSOption(Indice).Value = True
    Next Indice
    
End Sub

Public Sub Procura_Ego(Qual_Combo As String)
'Esta rotina preenche os TextBoxs e ComboBoxs com os dados do corrente Ego.
'Usa o valor ItemData do item atualmente selecionado no combo "Nome Indígena" ou "Nome Nacional"
'como índice de busca no DBcp_Ego.
    
    Dim MeuCriterio As String
    
    MeuCriterio = "ID_Ego = " & Qual_Combo 'Monta o critério baseado no ItemData do item selecionado no combo indicado pelo "Qual_Combo".
    PAC1.DBcp_Ego.Recordset.FindFirst MeuCriterio
    
    If PAC1.DBcp_Ego.Recordset.NoMatch = False Then 'Se findfirst foi bem sucedido então...
    'Tem ajusta o registro de DBcp_casais quando entra no tab(1).
    'NomePreferido
        NomePreferido = PAC1.DBcp_Ego.Recordset("Nome_Preferido")
        Select Case NomePreferido
            Case 1 'Nome Indígena
                PAC1.SSOption(10).Value = True
            Case 2 'Nome Nacional
                PAC1.SSOption(11).Value = True
        End Select
    
    'TextBox
        'Nome_Ind, Nome_Nac, Casa_Mora, Clã
            PAC1.Text(3).Text = PAC1.DBcp_Ego.Recordset("Obs")
            PAC1.Combo(9).Text = IIf(IsNull(PAC1.DBcp_Ego.Recordset("Nome_Ind")) = True, "", PAC1.DBcp_Ego.Recordset("Nome_Ind"))
            PAC1.Combo(10).Text = IIf(IsNull(PAC1.DBcp_Ego.Recordset("Nome_Nac")) = True, "", PAC1.DBcp_Ego.Recordset("Nome_Nac"))
            PAC1.Combo(11).Text = IIf(IsNull(PAC1.DBcp_Ego.Recordset("Casa_Mora")) = True, "", PAC1.DBcp_Ego.Recordset("Casa_Mora"))
            PAC1.Combo(12).Text = PAC1.DBcp_Ego.Recordset("Clã")
            'O valor é testado quanto a ser nulo ou não.
            PAC1.Combo(14).Text = IIf(IsNull(PAC1.DBcp_Ego.Recordset("Ajudante")) = True, "", PAC1.DBcp_Ego.Recordset("Ajudante"))
    
    'ComboBox
        'Lugar_Mora
            If PAC1.DBcp_Ego.Recordset("Lugar_Mora") <> 0 Then
                'Procura o índice em DBcp_Apoio_Lugar, para depois pegar o Texto correspondente.
                PAC1.DBcp_Apoio_Lugar.Recordset.FindFirst "ID_Apoio =" & PAC1.DBcp_Ego.Recordset("Lugar_Mora")
                'Se findfirst foi bem sucedido então...
                If PAC1.DBcp_Apoio_Lugar.Recordset.NoMatch = False Then PAC1.Combo(13).Text = PAC1.DBcp_Apoio_Lugar.Recordset("Apoio")
            Else
                'Caso o corrente Lugar_Mora=0 então o combo.text será vazio
                PAC1.Combo(13).Text = ""
            End If
    
        'Lugar_Nasc
            If PAC1.DBcp_Ego.Recordset("Lugar_Nasc") <> 0 Then
                'Procura o índice em DBcp_Apoio_Lugar, para depois pegar o Texto correspondente.
                PAC1.DBcp_Apoio_Lugar.Recordset.FindFirst "ID_Apoio =" & PAC1.DBcp_Ego.Recordset("Lugar_Nasc")
                'Se findfirst foi bem sucedido então...
                If PAC1.DBcp_Apoio_Lugar.Recordset.NoMatch = False Then PAC1.Combo(0).Text = PAC1.DBcp_Apoio_Lugar.Recordset("Apoio")
            Else
                'Caso o corrente Lugar_Mora=0 então o combo.text será vazio
                PAC1.Combo(0).Text = ""
            End If
    
    'Data
        If PAC1.DBcp_Ego.Recordset("Data") <> Empty Then
            PAC1.MaskCaixa(0).Text = Format(PAC1.DBcp_Ego.Recordset("Data"), "dd/mm/yyyy")
        Else
            PAC1.MaskCaixa(0).Text = "__/__/____"
        End If
    
    'Data de Nascimento
        If PAC1.DBcp_Ego.Recordset("Data_Nasc") <> Empty Then
            PAC1.MaskCaixa(1).Text = Format(PAC1.DBcp_Ego.Recordset("Data_Nasc"), "dd/mm/yyyy")
        Else
            PAC1.MaskCaixa(1).Text = "__/__/____"
        End If
    
    'Data de Falecimento
        If PAC1.DBcp_Ego.Recordset("Data_Falec") <> Empty Then
            PAC1.MaskCaixa(2).Text = Format(PAC1.DBcp_Ego.Recordset("Data_Falec"), "dd/mm/yyyy")
        Else
            PAC1.MaskCaixa(2).Text = "__/__/____"
        End If
    
    
    
    
    
    
    'Família de Orientação
        If PAC1.SSOption(3).Value = True Then Call Orientação
    
    'Família de Procriação
        If PAC1.SSOption(4).Value = True Then Call Procriação
            
    End If

End Sub

Public Sub Enche_Combos()
'Esta rotina enche apenas os combos relativos ao Ego em Casas-Pessoal.
    'PAC1.DBcp_Ego.Refresh
    'PAC1.DBcp_Ego.Refresh 'Reinicializa o DBcp_Ego
    PAC1.DBcp_Ego.Recordset.MoveFirst 'Move para o primeiro registro em DBcp_Ego.
    
    PAC1.DBcp_Apoio_Lugar.Refresh 'Reinicializa o DBcp_Apoio_Lugar
    PAC1.DBcp_Apoio_Lugar.Recordset.MoveFirst 'Move para o primeiro registro em DBcp_Apoio_Lugar.
    
    'Acrescenta o primeiro item do primeiro registro em DBcp_Ego no combo(11).
    PAC1.Combo(11).AddItem PAC1.DBcp_Ego.Recordset("Casa_Mora")
    'Acrescenta o primeiro item do primeiro registro em DBcp_Ego no combo(12)
    PAC1.Combo(12).AddItem PAC1.DBcp_Ego.Recordset("Clã")
    'Acrescenta o primeiro item do primeiro registro em DBcp_Ego combo(14)
    PAC1.Combo(14).AddItem PAC1.DBcp_Ego.Recordset("Ajudante")
    
'Lugar
    Call Enche_Combo_Lugar

    Do Until PAC1.DBcp_Ego.Recordset.EOF = True 'Loop até que chegar ao último registro de DBcp_Ego.
        
'Nome
        If PAC1.DBcp_Ego.Recordset("Nome_Ind") <> "" Then 'Nome_Ind do corrente registro só será tratado se não estiver vazio.
            'Acrescenta um novo item no Combo(9) usando o corrente registro de dbcp_ego.
            PAC1.Combo(9).AddItem PAC1.DBcp_Ego.Recordset("Nome_Ind")
            'Associa o valor de ID_Ego para o novo item inserido usando o ItemData.
            PAC1.Combo(9).ItemData(PAC1.Combo(9).NewIndex) = PAC1.DBcp_Ego.Recordset("ID_Ego")
        End If
        If PAC1.DBcp_Ego.Recordset("Nome_Nac") <> "" Then 'Nome_Nac do corrente registro só será tratado se não estiver vazio.
            'Acrescenta um novo item no Combo(10) usando o corrente registro de dbcp_ego.
            PAC1.Combo(10).AddItem PAC1.DBcp_Ego.Recordset("Nome_Nac")
            'Associa o valor de ID_Ego para o novo item inserido usando o ItemData.
            PAC1.Combo(10).ItemData(PAC1.Combo(10).NewIndex) = PAC1.DBcp_Ego.Recordset("ID_Ego")
        End If
        
'Casa_Mora, Clã, Ajudante.
        'Chama a rotina que testa se o corrente item de PAC1.DBcp_Ego já existe no Combo(11)
        Call Testa_Item(PAC1.DBcp_Ego, PAC1.Combo(11), "Casa_Mora")
        'Chama a rotina que testa se o corrente item de PAC1.DBcp_Ego já existe no Combo(12)
        Call Testa_Item(PAC1.DBcp_Ego, PAC1.Combo(12), "Clã")
        'Chama a rotina que testa se o corrente item de PAC1.DBcp_Ego já existe no Combo(14)
        Call Testa_Item(PAC1.DBcp_Ego, PAC1.Combo(14), "Ajudante")
        
        PAC1.DBcp_Ego.Recordset.MoveNext 'Move para o próximo registro em DBcp_Ego.
    
    Loop
    
    PAC1.DBcp_Ego.Recordset.MoveFirst 'Move para o primeiro registro em DBcp_Ego.

End Sub

Public Sub Grava_Ego()
'Esta rotina grava os dados do Ego, seja um novo Ego ou apenas a modificação feita em um registro já existente.
    
    On Error GoTo MeusErros
    'On Error Resume Next 'Qualquer erro que ocorrer, o PAC seguirá para a próxima linha.
    
'ID_Ego
    'Se o botão "Novo" estiver com o texto "Cancelar", então um novo ID_Ego deve ser lançado, isto
    'não deve acontecer com o botão "Editar", pois ele nunca vai editar ou criar um novo ID_Ego
    If PAC1.SSCommand(0).Caption = LoadResString(135 + Lingua) Then
        PAC1.DBcp_Ego.Recordset("ID_Ego") = PAC1.DBcp_Ego.Recordset.RecordCount
'Nome_Ind
        If PAC1.Combo(9).Text <> "" Then PAC1.DBcp_Ego.Recordset("Nome_Ind") = PAC1.Combo(9).Text & " (" & CStr(PAC1.DBcp_Ego.Recordset.RecordCount) & ")"
'Nome_Nac
        If PAC1.Combo(10).Text <> "" Then PAC1.DBcp_Ego.Recordset("Nome_Nac") = PAC1.Combo(10).Text & " (" & CStr(PAC1.DBcp_Ego.Recordset.RecordCount) & ")"
    Else
        PAC1.DBcp_Ego.Recordset("Nome_Ind") = PAC1.Combo(9).Text
        PAC1.DBcp_Ego.Recordset("Nome_Nac") = PAC1.Combo(10).Text
    End If
'Data
    If PAC1.MaskCaixa(0).ClipText = "" Then
        PAC1.DBcp_Ego.Recordset("Data") = Empty
    Else
        PAC1.DBcp_Ego.Recordset("Data") = PAC1.MaskCaixa(0).Text
    End If

'NomePreferido
    If PAC1.SSOption(10).Value = True Then
        PAC1.DBcp_Ego.Recordset("Nome_Preferido") = 1 'Nome Indígena
    Else
        PAC1.DBcp_Ego.Recordset("Nome_Preferido") = 2 'Nome Nacional
    End If


'Sexo
    PAC1.DBcp_Ego.Recordset("Sexo") = SexoEgo
        
        'PAC1.dbcp_ego.Recordset("Data_Nasc") = CDate(PAC1.MaskEdBox(1).Text)

'Lugar_Mora
    Call Lugar(PAC1.Combo(13), "Lugar_Mora")
    
'Lugar_Nasc
    If PAC1.Combo(0).Text <> "" Then Call Lugar(PAC1.Combo(0), "Lugar_Nasc")
    
'Casa_Mora
    If PAC1.Combo(11).Text = "" Then
        PAC1.DBcp_Ego.Recordset("Casa_Mora") = Empty
    Else
        PAC1.DBcp_Ego.Recordset("Casa_Mora") = PAC1.Combo(11).Text
    End If
    
'Civil
    PAC1.DBcp_Ego.Recordset("Civil") = CivilEgo

'Data de Nascimento
    If PAC1.MaskCaixa(1).ClipText = "" Then
        PAC1.DBcp_Ego.Recordset("Data_Nasc") = Empty
    Else
        PAC1.DBcp_Ego.Recordset("Data_Nasc") = PAC1.MaskCaixa(1).Text
    End If

'Data de Falecimento
    If PAC1.MaskCaixa(2).ClipText = "" Then
        PAC1.DBcp_Ego.Recordset("Data_Falec") = Empty
    Else
        PAC1.DBcp_Ego.Recordset("Data_Falec") = PAC1.MaskCaixa(2).Text
    End If

'Ajudante
    PAC1.DBcp_Ego.Recordset("Ajudante") = PAC1.Combo(14).Text
        
        'PAC1.dbcp_ego.Recordset("Data") = PAC1.MaskEdBox(0).Text

'ID_Anota
    PAC1.DBcp_Ego.Recordset("ID_Anota") = PAC1.Text(1).Text
    
'Pg_Anota
    PAC1.DBcp_Ego.Recordset("Pg_Anota") = IIf(PAC1.Text(2).Text <> "", PAC1.Text(2).Text, 0)
        
        'PAC1.dbcp_ego.recordset("Imprimiu")=
    
'clã
    PAC1.DBcp_Ego.Recordset("clã") = PAC1.Combo(12).Text
    
'Obs
    PAC1.DBcp_Ego.Recordset("Obs") = PAC1.Text(3).Text
    
'Atualiza o dbcp_ego com os novos dados ou editados.
    CorrenteEgo = PAC1.DBcp_Ego.Recordset("ID_Ego")
    PAC1.DBcp_Ego.Recordset.Update
    
    PAC1.DBcp_Ego.Refresh
    EgoAtualizado = 1

MeusErros:
    Beep

End Sub


Public Function Testa_Ego(controle As Control, Campo As String)
'Esta rotina testa se os campos obrigatórios foram preenchidos ou não.
    
    If controle.Text = "" Then 'Se o campo estiver vazio, então...
        'Uma mensagem é dada ao usuário depedendo do "campo" que ele não preencheu.
        'Todo o texto é tirado do PAC1.res, pois depende da língua em uso na interface.
        Mensagem_Erro = MsgBox(LoadResString(280 + Lingua) & Campo & LoadResString(281 + Lingua), 48, LoadResString(281 + Lingua)) 'Atenção!
        Testa_Ego = False 'Devolve o resultado de que o campo está vazio.
    Else 'Caso o campo esteja preenchido...
        Testa_Ego = True 'Devolve o resultado de que o campo está preenchido.
    End If
    
End Function

Public Sub Testa_Item(BD As Control, Comb As Control, Campo As String)
'Esta rotina testa se um item já foi colocado em um determinado combo.
    
    Dim conta As Integer 'Será usada para contar os itens dentro do corrente combo.
    Dim Existe As Integer 'Identificará se o corrente item existe ou não no corrente combo.
    
    Existe = False 'Inicializa a variável como "False".
    
    If BD.Recordset(Campo) <> "" Then 'Se o corrente registro não está vazio, então...
        
        'A quantidade de itens no combo deve ser subtraido de 1, pois o primeiro item sempre
        'terá o índice=0, E a variável "conta" já vai começar com o valor 0.
        For conta = 0 To Comb.ListCount - 1
            Debug.Print CStr(Comb.List(conta)), CStr(BD.Recordset(Campo))
            'Os valores Comb.List(conta) e BD.Recordset(campo) devem ser convertidos para string,
            'dependendo da situação, ele pode vir como valor numérico do Banco de Dados.
            If CStr(Comb.List(conta)) = CStr(BD.Recordset(Campo)) Then 'Se os valores forem exatamente iguais, então...
                Existe = True 'Indica que o item já existe no corrente combo.
                Exit For 'Sai do loop For...Next.
            End If
        
        Next conta
        
        'Se o item não existe no corrente combo "Existe=False", então ele é acrescentado no combo.
        If Existe = False Then Comb.AddItem BD.Recordset(Campo)
        
    End If

End Sub

Public Sub Lugar(Comb As Control, Campo As String)
'Esta rotina verifica se o "lugar de nascimento" e o "lugar que mora" existem,
'caso não existam, um novo índice é criado em DBcp_Apoio_Lugar para este novo lugar
    
    Dim CorrenteLugar As Integer
    Dim Lugar As String 'Receberá a string do nome do lugar para ser manipulada pela função Tira_Apostrofe.
    Dim MeuCritério As String 'Receberá a string que descreve o critério de busca no DBcp_Apoio_Lugar.
    
    CorrenteLugar = PAC1.DBcp_Ego.Recordset(Campo) '####
    Lugar = Comb.Text
    MeuCritério = "Apoio ='" & NovaString(Lugar) & "'"
    
    'O lugar em "Comb.text" é procurado em DBcp_Apoio_Lugar.
    PAC1.DBcp_Apoio_Lugar.Recordset.FindFirst MeuCritério
    'Se o lugar não foi achado, então...
    If PAC1.DBcp_Apoio_Lugar.Recordset.NoMatch = True Then
        'Se dbcp_ego está sendo editado, então...
        If PAC1.DBcp_Ego.EditMode = dbEditInProgress Then
            PAC1.DBcp_Apoio_Lugar.Recordset.FindFirst "ID_Apoio=" & CorrenteLugar
            
            PAC1.DBcp_Apoio_Lugar.Recordset.Edit 'Abre o dbcp_Apoio_Lugar para editar o corrente registro.
            'Associa o novo Lugar
            PAC1.DBcp_Apoio_Lugar.Recordset("Apoio") = Lugar
        
        'Caso o dbcp_ego não está sendo editado...
        Else
            PAC1.DBcp_Apoio_Lugar.Recordset.AddNew 'Abre o dbcp_Apoio_Lugar para receber o novo registro.
            'Associa o novo índice.
            PAC1.DBcp_Apoio_Lugar.Recordset("ID_Apoio") = PAC1.DBcp_Apoio_Lugar.Recordset.RecordCount + 7
            'Associa o novo Lugar
            PAC1.DBcp_Apoio_Lugar.Recordset("Apoio") = Lugar
        End If
        PAC1.DBcp_Apoio_Lugar.Recordset.Update 'O dbcp é atualizado.
        
        'Uma nova busca deve ser executado já que o texto de "Apoio" pode ter sido mudado.
        PAC1.DBcp_Apoio_Lugar.Recordset.FindFirst MeuCritério
    End If
    
    Comb.Text = PAC1.DBcp_Apoio_Lugar.Recordset("Apoio")
    PAC1.DBcp_Ego.Recordset(Campo) = PAC1.DBcp_Apoio_Lugar.Recordset("ID_Apoio")

End Sub

Public Sub Enche_Combo_Lugar()
    Do Until PAC1.DBcp_Apoio_Lugar.Recordset.EOF = True 'Loop até que chegar ao último registro de DBcp_Apoio_Lugar.
        PAC1.Combo(0).AddItem PAC1.DBcp_Apoio_Lugar.Recordset("Apoio")
            'Associa o valor de ID_Ego para o novo item inserido usando o ItemData.
            'PAC1.Combo(0).ItemData(PAC1.Combo(0).NewIndex) = PAC1.DBcp_Apoio_Lugar.Recordset("ID_Apoio")
            'Debug.Print PAC1.Combo(0).ItemData(PAC1.Combo(0).NewIndex)
        PAC1.Combo(13).AddItem PAC1.DBcp_Apoio_Lugar.Recordset("Apoio")
            'Associa o valor de ID_Ego para o novo item inserido usando o ItemData.
            'PAC1.Combo(13).ItemData(PAC1.Combo(13).NewIndex) = PAC1.DBcp_Apoio_Lugar.Recordset("ID_Apoio")
        PAC1.DBcp_Apoio_Lugar.Recordset.MoveNext 'Move para o próximo registro em DBcp_Apoio_Lugar.
    Loop
End Sub

Public Sub Orientação()
        Dim Chave As String
        Dim Critério As String
        PAC1.TreeView1.Nodes.Clear
        If PAC1.DBcp_Ego.Recordset("ID_Pais") <> 0 Then
            EstouCarregando = 1
            PAC1.DBcp_Casais.Recordset.FindFirst "ID_Casal=" & PAC1.DBcp_Ego.Recordset("ID_Pais")
            PAC1.DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & PAC1.DBcp_Casais.Recordset("ID_Conj1")
            QualNomePai = IIf(PAC1.DBcp_Ego_Nomes.Recordset("Nome_Preferido") = 1, PAC1.DBcp_Ego_Nomes.Recordset("Nome_Ind"), PAC1.DBcp_Ego_Nomes.Recordset("Nome_Nac"))
            
            If PAC1.DBcp_Ego.Recordset("Sexo") = 0 Then
                Chave = ImagemCasal("Orientação")
            End If
            
            
            PAC1.DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & PAC1.DBcp_Casais.Recordset("ID_Conj2")
            QualNomeMãe = IIf(PAC1.DBcp_Ego_Nomes.Recordset("Nome_Preferido") = 1, PAC1.DBcp_Ego_Nomes.Recordset("Nome_Ind"), PAC1.DBcp_Ego_Nomes.Recordset("Nome_Nac"))
            
            If PAC1.DBcp_Ego.Recordset("Sexo") = 1 Then
                Chave = ImagemCasal("Orientação")
            End If
           
            
            ' Add Node objects.
            Dim nodX As Node    ' Declare Node variable.
            ' First node with 'Root' as text.
            Set nodX = PAC1.TreeView1.Nodes.Add(, , "Pais", "              " & QualNomePai & "  &  " & QualNomeMãe, Chave)
            
            Critério = "ID_Pais=" & PAC1.DBcp_Ego.Recordset("ID_Pais")
            PAC1.DBcp_Ego_Nomes.Recordset.FindFirst Critério
            If PAC1.DBcp_Ego_Nomes.Recordset.NoMatch = False Then
                Do Until PAC1.DBcp_Ego_Nomes.Recordset.NoMatch = True
                Call filhos(Critério, "Pais")
                Loop
                EstouCarregando = 0
            End If
        End If

End Sub

Public Sub Procriação()
    Dim CriterioCasais As String
    Dim CriterioNomes As String
    Dim Chave As String
    Dim contador As Integer
    PAC1.TreeView1.Nodes.Clear
    contador = 1
    EstouCarregando = 1
'O CASAL.
        CriterioCasais = IIf(PAC1.DBcp_Ego.Recordset("Sexo") = 0, "ID_Conj1 =" & PAC1.DBcp_Ego.Recordset("ID_Ego"), "ID_Conj2 =" & PAC1.DBcp_Ego.Recordset("ID_Ego"))
        
'        If PAC1.DBcp_Ego.Recordset("Sexo") = 0 Then
'            PAC1.DBcp_Casais.Recordset.FindFirst "ID_Conj1 =" & PAC1.DBcp_Ego.Recordset("ID_Ego")
'        Else
'            PAC1.DBcp_Casais.Recordset.FindFirst "ID_Conj2 =" & PAC1.DBcp_Ego.Recordset("ID_Ego")
'        End If
        PAC1.DBcp_Casais.Recordset.FindFirst CriterioCasais
        'Se Encontrou o ID do corrente ego no BD dos casais, então...
'        If PAC1.DBcp_Casais.Recordset.NoMatch = False Then
            
        Do While PAC1.DBcp_Casais.Recordset.NoMatch = False
            PAC1.DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & PAC1.DBcp_Casais.Recordset("ID_Conj1")
            QualNomeConj1 = IIf(PAC1.DBcp_Ego_Nomes.Recordset("Nome_Preferido") = 1, PAC1.DBcp_Ego_Nomes.Recordset("Nome_Ind"), PAC1.DBcp_Ego_Nomes.Recordset("Nome_Nac"))
            If PAC1.DBcp_Ego.Recordset("Sexo") = 0 Then
                Chave = ImagemCasal("Procriação")
            End If
            PAC1.DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & PAC1.DBcp_Casais.Recordset("ID_Conj2")
            QualNomeConj2 = IIf(PAC1.DBcp_Ego_Nomes.Recordset("Nome_Preferido") = 1, PAC1.DBcp_Ego_Nomes.Recordset("Nome_Ind"), PAC1.DBcp_Ego_Nomes.Recordset("Nome_Nac"))
            If PAC1.DBcp_Ego.Recordset("Sexo") = 1 Then
                Chave = ImagemCasal("Procriação")
            End If
            ' Primeiro Node sendo os "Pais".
            Set nodX = PAC1.TreeView1.Nodes.Add(, , "Pais" & contador, "              " & QualNomeConj1 & "  &  " & QualNomeConj2, Chave)
'OS FILHOS
            CriterioNomes = "ID_Pais=" & PAC1.DBcp_Casais.Recordset("ID_Casal")
            PAC1.DBcp_Ego_Nomes.Recordset.FindFirst CriterioNomes
            If PAC1.DBcp_Ego_Nomes.Recordset.NoMatch = False Then
                Do Until PAC1.DBcp_Ego_Nomes.Recordset.NoMatch = True
                    If PAC1.DBcp_Ego.Recordset("ID_Ego") <> PAC1.DBcp_Ego_Nomes.Recordset("ID_Ego") Then
                        Call filhos(CriterioNomes, "Pais" & contador)
                    End If
                Loop
            End If
            contador = contador + 1
            PAC1.DBcp_Casais.Recordset.FindNext CriterioCasais
        Loop
'        End If
    EstouCarregando = 0
End Sub

Public Sub filhos(Critério As String, Pais As String)
    Dim Ícone As Integer
    Dim nodX As Node
    QualNomeFilho = IIf(PAC1.DBcp_Ego_Nomes.Recordset("Nome_Preferido") = 1, PAC1.DBcp_Ego_Nomes.Recordset("Nome_Ind"), PAC1.DBcp_Ego_Nomes.Recordset("Nome_Nac"))
    If PAC1.DBcp_Ego_Nomes.Recordset("Sexo") = 0 Then
        Ícone = IIf(PAC1.DBcp_Ego_Nomes.Recordset("ID_Ego") = PAC1.DBcp_Ego.Recordset("ID_Ego"), 6, 8)
    Else
        Ícone = IIf(PAC1.DBcp_Ego_Nomes.Recordset("ID_Ego") = PAC1.DBcp_Ego.Recordset("ID_Ego"), 2, 4)
    End If
    If PAC1.DBcp_Ego_Nomes.Recordset("Data_Falec") <> "" Then Ícone = Ícone - 1
    ' This next node is a child of Node 1 ("Root").
    Set nodX = PAC1.TreeView1.Nodes.Add(Pais, tvwChild, , QualNomeFilho, Ícone)
    nodX.EnsureVisible
    PAC1.DBcp_Ego_Nomes.Recordset.FindNext Critério
End Sub

Public Function ImagemCasal(QualTipo As String) As String
    Dim Chave As String
    If PAC1.DBcp_Casais.Recordset("Civil") = 3 Then Chave = Chave & "B"
    If PAC1.DBcp_Casais.Recordset("Civil") = 5 Then Chave = Chave & "C"
    If PAC1.DBcp_Casais.Recordset("Civil") = 6 Then Chave = Chave & "D"
    If QualTipo = "Procriação" Then
        If PAC1.DBcp_Ego.Recordset("Sexo") = 0 Then Chave = Chave & "2"
        If PAC1.DBcp_Ego.Recordset("Sexo") = 1 Then Chave = Chave & "3"
        If PAC1.DBcp_Casais.Recordset("ID_Conj1") = PAC1.DBcp_Ego.Recordset("ID_Ego") Then
            QuemBuscar1 = "ID_Conj1"
            QuemBuscar2 = "ID_Conj2"
        Else
            QuemBuscar1 = "ID_Conj2"
            QuemBuscar2 = "ID_Conj1"
        End If
    Else
        Chave = Chave & "1"
        QuemBuscar1 = "ID_Conj1"
        QuemBuscar2 = "ID_Conj2"

    End If
        PAC1.DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & PAC1.DBcp_Casais.Recordset(QuemBuscar1)
        If Format(PAC1.DBcp_Ego_Nomes.Recordset("Data_Falec"), "dd/mm/yyyy") = Empty Then
            PAC1.DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & PAC1.DBcp_Casais.Recordset(QuemBuscar2)
            If Format(PAC1.DBcp_Ego_Nomes.Recordset("Data_Falec"), "dd/mm/yyyy") = Empty Then
                Chave = Chave & "a"
            Else
                Chave = Chave & "b"
            End If
        Else
            PAC1.DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & PAC1.DBcp_Casais.Recordset(QuemBuscar2)
            If Format(PAC1.DBcp_Ego_Nomes.Recordset("Data_Falec"), "dd/mm/yyyy") = Empty Then
                Chave = Chave & "c"
            Else
                Chave = Chave & "d"
            End If
        End If
    ImagemCasal = Chave
End Function


Public Function NovaString(sInput As String) As String
   Dim sNewString As String
   Dim I As Integer

   For I = 1 To Len(sInput)
      If Mid$(sInput, I, 1) = "'" Then
         sNewString = sNewString & "''"
      Else
         sNewString = sNewString & Mid$(sInput, I, 1)
      End If
   Next I

   NovaString = sNewString
End Function
