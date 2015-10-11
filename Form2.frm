VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Novo_Termo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acrescentar Novo Termo T�cnico"
   ClientHeight    =   4956
   ClientLeft      =   2244
   ClientTop       =   3288
   ClientWidth     =   5880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4956
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame Frame 
      Height          =   3012
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5652
      _Version        =   65536
      _ExtentX        =   9970
      _ExtentY        =   5313
      _StockProps     =   14
      Caption         =   "Forme o novo Termo T�cnico cliclando nos bot�es abaixo:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox Text 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   888
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2000
         Width           =   5400
      End
      Begin Threed.SSCommand Botao_Termo 
         Height          =   492
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   868
         _StockProps     =   78
         Caption         =   "Pai"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   1
         Outline         =   0   'False
         MouseIcon       =   "Form2.frx":0000
      End
   End
   Begin Threed.SSCommand Botao_Novo_Termo 
      Height          =   252
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   4596
      Width           =   960
      _Version        =   65536
      _ExtentX        =   1693
      _ExtentY        =   444
      _StockProps     =   78
      Caption         =   "&OK"
   End
   Begin Threed.SSCommand Botao_Novo_Termo 
      Height          =   252
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   4596
      Width           =   960
      _Version        =   65536
      _ExtentX        =   1693
      _ExtentY        =   444
      _StockProps     =   78
      Caption         =   "&Cancelar"
   End
   Begin Threed.SSCommand Botao_Novo_Termo 
      Height          =   252
      Index           =   2
      Left            =   3720
      TabIndex        =   5
      Top             =   4596
      Width           =   960
      _Version        =   65536
      _ExtentX        =   1693
      _ExtentY        =   444
      _StockProps     =   78
      Caption         =   "&Apagar"
   End
   Begin Threed.SSFrame Frame 
      Height          =   1300
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   5652
      _Version        =   65536
      _ExtentX        =   9970
      _ExtentY        =   2293
      _StockProps     =   14
      Caption         =   "Observa��o"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox Text 
         Height          =   888
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   5400
      End
   End
End
Attribute VB_Name = "Novo_Termo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Botao_Novo_Termo_Click(Index As Integer)
    Select Case Index
        Case 0 'OK
            Dim Novo_Termo_Inserido As Integer 'Declara a vari�vel que receber� o ID do novo termo.
            Dim Tip As Integer 'Declara a vari�vel do Tipo de termo de tratamento.
            Dim Sex As Integer 'Declara a vari�vel do Sexo do ego.
            Dim Vai As Integer 'Declara a vari�vel de controle do loop
            PAC1.DB_Temp.DatabaseName = "dbpa.mdb"
            PAC1.DB_Temp.RecordSource = "Termos_Tec" 'Configura o DB_Temp para esta tabela.
            PAC1.DB_Temp.Refresh 'Reinicia o DB_Temp
            PAC1.DB_Temp.Recordset.AddNew
            If Lingua = 0 Then
                PAC1.DB_Temp.Recordset("Termo_Tec") = Text(0).Text
                PAC1.DB_Temp.Recordset("Termo_Tec_IN") = Text(0).Tag
            Else
                PAC1.DB_Temp.Recordset("Termo_Tec") = Text(0).Tag
                PAC1.DB_Temp.Recordset("Termo_Tec_IN") = Text(0).Text
            End If
            PAC1.DB_Temp.Recordset("Trilha") = Novo_Termo.Tag
            PAC1.DB_Temp.Recordset("Obs") = Text(1).Text
            PAC1.DB_Temp.Recordset.Update
            PAC1.DB_Temp.Recordset.MoveLast 'Vai para o �ltimo registro, pois ele foi o rec�m inserido.
            Novo_Termo_Inserido = PAC1.DB_Temp.Recordset("ID_Termo_Tec")  'Pega aqui o ID do registro.
            'Lan�a o novo termo iserido pelo usu�rio no banco de dados para ser pesquisado em todos _
             os ambientes (Masculino-Refer�ncia e tratamento, Feminino-Refer�ncia e tratamento.
            PAC1.DB_Temp.RecordSource = "Termos_Confirmados" 'Configura o DB_Temp para esta tabela.
            PAC1.DB_Temp.Refresh 'Reinicia o DB_Temp
            For Vai = 1 To 4 'O loop vai rodar 4 vezes.
                Tip = IIf(Vai < 3, 0, 1) 'Nas duas primeiras voltas a vari�vel Tip valer� 0, depois passar� a valer 1
                Sex = IIf(Vai = 1 Or Vai = 3, 1, 0) 'Na volta 1 e 3 a vari�vel Sex valer� 1, nas outras voltas valer� 0
                PAC1.DB_Temp.Recordset.AddNew
                PAC1.DB_Temp.Recordset("ID_Termo_Tec") = Novo_Termo_Inserido
                PAC1.DB_Temp.Recordset("ID_Tipo") = Tip
                PAC1.DB_Temp.Recordset("Sexo_Ego") = Sex
                PAC1.DB_Temp.Recordset.Update
            Next Vai
            PAC1.Label(33) = PAC1.DBpa_Termos.Recordset.RecordCount 'Coloca o n�mero de registros do DBpa no label 33.
            Unload Novo_Termo 'Descarrega o formul�rio Novo_Termo
        Case 1 'CANCELAR
            Unload Novo_Termo 'Descarrega o formul�rio Novo_Termo
        Case 2 'APAGAR
            Text(0).Text = "" 'Limpa a caixa de texto
            Text(0).Tag = "" 'Limpa o Tag da caixa de texto
            Text(1).Text = "" 'Limpa a caixa de texto de observa��es.
            Novo_Termo.Tag = "" 'Limpa o Tag do formul�rio
            Frame(0).Tag = "" 'Limpa o Tag do frame
    End Select
End Sub
Private Sub Botao_Termo_Click(Index As Integer)
'Escolhe o termo de parentesco conforme selecionado pelo usu�rio.
'A vari�vel Lingua=0 � portugu�s e Lingua=2300 � Ingl�s.
    Select Case Index
        Case 0
            Debug.Print Monta_Termo(Lingua, "Pai", "Father", "1")
        Case 1
            Debug.Print Monta_Termo(Lingua, "M�e", "Mother", "2")
        Case 2
            Debug.Print Monta_Termo(Lingua, "Irm�o", "Brother", "3")
        Case 3
            Debug.Print Monta_Termo(Lingua, "Irm�", "Sister", "4")
        Case 4
            Debug.Print Monta_Termo(Lingua, "Filho", "Son", "5")
        Case 5
            Debug.Print Monta_Termo(Lingua, "Filha", "Daughter", "6")
        Case 6
            Debug.Print Monta_Termo(Lingua, "Esposo", "Husband", "7")
        Case 7
            Debug.Print Monta_Termo(Lingua, "Esposa", "Wife", "8")
    End Select
    Debug.Print Text(0).Tag

End Sub

Private Sub Form_Load()
    Dim I As Integer
'CENTRA O FORMUL�RIO NA TELA.
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2
    
'Lingua=0 ===> Portugu�s     Lingua=2300 ===> Ingl�s
    Botao_Novo_Termo(0).Caption = LoadResString(132 + Lingua) 'Inclui o texto "Ok" na l�ngua corrente
    Botao_Novo_Termo(1).Caption = LoadResString(135 + Lingua) 'Inclui o texto "Cancelar" na l�ngua corrente
    Botao_Novo_Termo(2).Caption = LoadResString(133 + Lingua) 'Inclui o texto "Apagar" na l�ngua corrente
    Botao_Termo(0).Caption = IIf(Lingua = 0, "Pai", "Father") 'Escolhe conforme a l�ngua
    
'OS BOT�ES COM OS TERMOS DE PARENTESCO B�SICO.
    For I = 1 To 7 'Este loop de 7 voltas carrega 7 novos bot�es e configura-os com a lingua atual.
        Load Botao_Termo(I) 'Carrega o novo bot�o.
        If I > 3 Then 'Do 4� bot�o em diante a localiza��o ser� na linha de baixo.
            Botao_Termo(I).Top = Botao_Termo(I - 4).Height + Botao_Termo(I - 4).Top + 200
            Botao_Termo(I).Left = (Botao_Termo(I - 4).Left + Botao_Termo(I - 4).Width) - Botao_Termo(0).Width
        Else 'At� o 3� bot�o todos ficam na mesma linha.
            Botao_Termo(I).Left = 200 + Botao_Termo(I - 1).Left + Botao_Termo(I - 1).Width
        End If
        Select Case I 'Bloco que coloca o texto no bot�o conforme a l�ngua selecionada.
            Case 1
                Botao_Termo(I).Caption = IIf(Lingua = 0, "M�e", "Mother")
            Case 2
                Botao_Termo(I).Caption = IIf(Lingua = 0, "Irm�o", "Brother")
            Case 3
                Botao_Termo(I).Caption = IIf(Lingua = 0, "Irm�", "Sister")
            Case 4
                Botao_Termo(I).Caption = IIf(Lingua = 0, "Filho", "Son")
            Case 5
                Botao_Termo(I).Caption = IIf(Lingua = 0, "Filha", "Daughter")
            Case 6
                Botao_Termo(I).Caption = IIf(Lingua = 0, "Esposo", "Husband")
            Case 7
                Botao_Termo(I).Caption = IIf(Lingua = 0, "Esposa", "Wife")
        End Select
        Botao_Termo(I).Visible = True
    Next I
End Sub
Public Function Monta_Termo(Lingua As Integer, Primario As String, Secundario As String, Trilha As String) As String
'Esta fun��o monta o novo termo que o usu�rio quer inserir no Banco de Dados.
    Dim Artigo As String 'Recebe o artigo "do" ou "da" para ser usado no termo em portugu�s.
    'O artigo � determinado pelo n�mero da trilha. Os pares s�o femininos e �mpares s�o masculino.
    Artigo = Switch(Trilha = "1", " do ", Trilha = "2", " da ", _
                    Trilha = "3", " do ", Trilha = "4", " da ", _
                    Trilha = "5", " do ", Trilha = "6", " da ", _
                    Trilha = "7", " do ", Trilha = "8", " da ")
    If Lingua = 0 Then 'Processa aqui se a lingua corrente for portugu�s.
        If Text(0).Text = "" Then 'Se caixa de texto estiver vazia.
            Text(0).Text = Primario 'A primeira palavra � colocada.
            Text(0).Tag = Secundario 'A primeira palavra na outra lingua � guadada no Tag.
        Else 'Se a caixa de texto j� contem alguma coisa...
            Text(0).Text = Text(0).Text & Artigo & Primario 'A palavra � juntada no final do texto que j� est� na caixa juntamente com o artigo.
            Text(0).Tag = Secundario & "'s " & Text(0).Tag '� igual � linha anterior, s� que com a outra lingua.
        End If
         Novo_Termo.Tag = Trilha & Novo_Termo.Tag 'A trilha vai sendo montada com este TAG.
    Else 'Se a lingua corrente n�o � o portugu�s, o processo � aqui....
        If Text(0).Text = "" Then
            Text(0).Text = Secundario
            Text(0).Tag = Primario
        Else
            Text(0).Text = Text(0).Text & "'s " & Secundario
            'O artigo � escolhido pelo n�mero de trilha da palavra anteriormente selecionada e armazenado em frame(0).tag
            Artigo = Switch(Frame(0).Tag = "1", " do ", Frame(0).Tag = "2", " da ", _
                            Frame(0).Tag = "3", " do ", Frame(0).Tag = "4", " da ", _
                            Frame(0).Tag = "5", " do ", Frame(0).Tag = "6", " da ", _
                            Frame(0).Tag = "7", " do ", Frame(0).Tag = "8", " da ")
            Text(0).Tag = Primario & Artigo & Text(0).Tag
        End If
        Frame(0).Tag = Trilha 'A trilha da palavra atual � armazenada aqui.
        Novo_Termo.Tag = Novo_Termo.Tag & Trilha
    End If
    Monta_Termo = Novo_Termo.Tag
End Function
Private Sub Text_GotFocus(Index As Integer)
    If Index = 0 Then Botao_Termo(0).SetFocus    'Tira o foco da Caixa de texto, pois o usu�rio n�o pode edita-la.
End Sub
