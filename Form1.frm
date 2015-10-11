VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F72CC888-5ADC-101B-A56C-00AA003668DC}#1.0#0"; "ANIBTN32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form PAC1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAC - Programa de Análise Cultural 1.0"
   ClientHeight    =   6888
   ClientLeft      =   -48
   ClientTop       =   240
   ClientWidth     =   9564
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6888
   ScaleWidth      =   9564
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data DB_Temp 
      Caption         =   "DB_Temp"
      Connect         =   "Access"
      DatabaseName    =   "D:\MACS\PAC\DBpa.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Data DBpa_Termos 
      Caption         =   "DBpa_Termos"
      Connect         =   "Access"
      DatabaseName    =   "D:\MACS\PAC\DBpa.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Consulta_Geral"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Data DBcp_Ego_Nomes 
      Caption         =   "DBcp_Ego_Nomes"
      Connect         =   "Access"
      DatabaseName    =   "d:\macs\pac\DBcp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   4032
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from EGO"
      Top             =   6636
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Data DBcp_Ego_Feminino 
      Caption         =   "DBcp_Ego_Feminino"
      Connect         =   "Access"
      DatabaseName    =   "d:\macs\pac\DBcp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   2064
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * from EGO where Ego.sexo=1 and Ego.Nome_Ind<>''"
      Top             =   6612
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Data DBcp_Ego_Masculino 
      Caption         =   "DBcp_Ego_Masculino"
      Connect         =   "Access"
      DatabaseName    =   "d:\macs\pac\DBcp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   324
      Left            =   324
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select *  from EGO where sexo=0 And Ego.Nome_Ind<>''"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Data DBcp_Casais 
      Caption         =   "DBcp_Casais"
      Connect         =   "Access"
      DatabaseName    =   "d:\macs\pac\DBcp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   192
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CASAIS"
      Top             =   6900
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Data DBcp_Ego 
      Caption         =   "DBcp_Ego"
      Connect         =   "Access"
      DatabaseName    =   "d:\macs\pac\DBcp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   2064
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EGO"
      Top             =   6900
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Data DBcp_Apoio_Lugar 
      Caption         =   "DBcp_Apoio_Lugar"
      Connect         =   "Access"
      DatabaseName    =   "d:\macs\pac\DBcp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   6048
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from APOIO where id_apoio>6"
      Top             =   6900
      Visible         =   0   'False
      Width           =   2076
   End
   Begin VB.Data DBdc_Pesquisador 
      Caption         =   "DBdc_Pesquisador"
      Connect         =   "Access"
      DatabaseName    =   "d:\macs\pac\DBdc.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   3984
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PESQUISADOR"
      Top             =   6900
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.CommandButton Command 
      Caption         =   "Info"
      Height          =   252
      Index           =   1
      Left            =   8208
      TabIndex        =   238
      Top             =   48
      Width           =   840
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6810
      Index           =   1
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   9465
      _ExtentX        =   16701
      _ExtentY        =   12002
      _Version        =   327681
      Style           =   1
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   617
      BackColor       =   8421504
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Casas/Pessoal"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SSTab(2)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Parentesco"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSTab(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Acontecimentos"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab(4)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Descrições"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSFrame(21)"
      Tab(3).Control(1)=   "SSFrame(19)"
      Tab(3).Control(2)=   "SSFrame(23)"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "Form1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Administração"
      TabPicture(5)   =   "Form1.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "SSFrame(25)"
      Tab(5).Control(1)=   "SSFrame(27)"
      Tab(5).ControlCount=   2
      Begin TabDlg.SSTab SSTab 
         Height          =   6468
         Index           =   2
         Left            =   -75000
         TabIndex        =   1
         Top             =   336
         Width           =   9470
         _ExtentX        =   16701
         _ExtentY        =   11409
         _Version        =   327681
         TabOrientation  =   1
         Style           =   1
         TabHeight       =   600
         TabMaxWidth     =   1411
         WordWrap        =   0   'False
         BackColor       =   8421504
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Ego"
         TabPicture(0)   =   "Form1.frx":00A8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Papel(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Total"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SSCommand(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "SSCommand(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "SSCommand(3)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "SSCommand(1)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "SSFrame(2)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "SSFrame(1)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Família Nuclear"
         TabPicture(1)   =   "Form1.frx":00C4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSFrame(8)"
         Tab(1).Control(1)=   "SSFrame(7)"
         Tab(1).Control(2)=   "Papel(2)"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Análise Geral"
         TabPicture(2)   =   "Form1.frx":00E0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Papel(3)"
         Tab(2).Control(1)=   "SSFrame(9)"
         Tab(2).ControlCount=   2
         Begin Threed.SSFrame SSFrame 
            Height          =   912
            Index           =   1
            Left            =   84
            TabIndex        =   79
            Top             =   12
            Width           =   9264
            _Version        =   65536
            _ExtentX        =   16341
            _ExtentY        =   1609
            _StockProps     =   14
            Caption         =   "Dados de controle"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   14
               ItemData        =   "Form1.frx":00FC
               Left            =   1380
               List            =   "Form1.frx":00FE
               Sorted          =   -1  'True
               TabIndex        =   3
               Top             =   480
               Width           =   2652
            End
            Begin VB.TextBox Text 
               DataField       =   "ID_Anota"
               DataSource      =   "DBcp_Ego"
               Height          =   300
               Index           =   1
               Left            =   7728
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   480
               Width           =   732
            End
            Begin VB.TextBox Text 
               DataField       =   "Pg_Anota"
               DataSource      =   "DBcp_Ego"
               Height          =   300
               Index           =   2
               Left            =   8544
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   480
               Width           =   588
            End
            Begin MSMask.MaskEdBox MaskCaixa 
               Bindings        =   "Form1.frx":0100
               Height          =   288
               Index           =   0
               Left            =   168
               TabIndex        =   2
               Top             =   480
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   508
               _Version        =   327681
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSDBCtls.DBCombo DBCombo 
               Bindings        =   "Form1.frx":010F
               DataField       =   "ID_Pesquisador"
               DataSource      =   "DBcp_Ego"
               Height          =   288
               Index           =   2
               Left            =   4296
               TabIndex        =   4
               Top             =   480
               Width           =   3084
               _ExtentX        =   5440
               _ExtentY        =   508
               _Version        =   327681
               Style           =   2
               BackColor       =   16777215
               ForeColor       =   0
               ListField       =   "Pesq"
               BoundColumn     =   "ID_Pesquisador"
               Text            =   "DBCombo(2)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Página:"
               Height          =   192
               Index           =   5
               Left            =   8544
               TabIndex        =   174
               Top             =   288
               Width           =   552
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Caderno:"
               Height          =   192
               Index           =   4
               Left            =   7728
               TabIndex        =   173
               Top             =   288
               Width           =   660
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Pesquisador:"
               Height          =   192
               Index           =   3
               Left            =   4284
               TabIndex        =   172
               Top             =   288
               Width           =   960
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Ajudante:"
               Height          =   192
               Index           =   2
               Left            =   1392
               TabIndex        =   171
               Top             =   288
               Width           =   672
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Data:"
               Height          =   192
               Index           =   1
               Left            =   168
               TabIndex        =   170
               Top             =   288
               Width           =   396
               WordWrap        =   -1  'True
            End
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   5076
            Index           =   2
            Left            =   84
            TabIndex        =   80
            Top             =   948
            Width           =   9264
            _Version        =   65536
            _ExtentX        =   16341
            _ExtentY        =   8954
            _StockProps     =   14
            Caption         =   "Ego"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   13
               ItemData        =   "Form1.frx":012A
               Left            =   1525
               List            =   "Form1.frx":012C
               Sorted          =   -1  'True
               TabIndex        =   8
               Top             =   528
               Width           =   3660
            End
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   0
               ItemData        =   "Form1.frx":012E
               Left            =   168
               List            =   "Form1.frx":0130
               Sorted          =   -1  'True
               TabIndex        =   16
               Top             =   3312
               Width           =   2604
            End
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   12
               ItemData        =   "Form1.frx":0132
               Left            =   2960
               List            =   "Form1.frx":0134
               Sorted          =   -1  'True
               TabIndex        =   17
               Top             =   3312
               Width           =   2220
            End
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   11
               Left            =   168
               Sorted          =   -1  'True
               TabIndex        =   7
               Top             =   528
               Width           =   1164
            End
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   10
               Left            =   408
               Sorted          =   -1  'True
               TabIndex        =   10
               Top             =   1680
               Width           =   4770
            End
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   9
               Left            =   408
               Sorted          =   -1  'True
               TabIndex        =   9
               Top             =   1104
               Width           =   4770
            End
            Begin VB.TextBox Text 
               Height          =   1056
               Index           =   3
               Left            =   4560
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   24
               Top             =   3876
               Width           =   4512
            End
            Begin Threed.SSFrame SSFrame 
               Height          =   828
               Index           =   4
               Left            =   168
               TabIndex        =   11
               Top             =   2112
               Width           =   1476
               _Version        =   65536
               _ExtentX        =   2611
               _ExtentY        =   1460
               _StockProps     =   14
               Caption         =   "Sexo"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   3
               Begin Threed.SSOption SSOption 
                  Height          =   252
                  Index           =   2
                  Left            =   288
                  TabIndex        =   13
                  Top             =   480
                  Width           =   888
                  _Version        =   65536
                  _ExtentX        =   1566
                  _ExtentY        =   445
                  _StockProps     =   78
                  Caption         =   "Feminino"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSOption SSOption 
                  Height          =   252
                  Index           =   1
                  Left            =   288
                  TabIndex        =   12
                  Top             =   240
                  Width           =   972
                  _Version        =   65536
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   78
                  Caption         =   "Masculino"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   -1  'True
               End
            End
            Begin Threed.SSFrame SSFrame 
               Height          =   1020
               Index           =   6
               Left            =   168
               TabIndex        =   18
               Top             =   3792
               Width           =   4164
               _Version        =   65536
               _ExtentX        =   7345
               _ExtentY        =   1799
               _StockProps     =   14
               Caption         =   "Estado Civil"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   3
               Begin Threed.SSOption SSOption 
                  Height          =   300
                  Index           =   5
                  Left            =   288
                  TabIndex        =   19
                  Top             =   240
                  Width           =   924
                  _Version        =   65536
                  _ExtentX        =   1630
                  _ExtentY        =   529
                  _StockProps     =   78
                  Caption         =   "Solteiro"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   -1  'True
               End
               Begin Threed.SSOption SSOption 
                  Height          =   300
                  Index           =   6
                  Left            =   288
                  TabIndex        =   20
                  Top             =   600
                  Width           =   924
                  _Version        =   65536
                  _ExtentX        =   1630
                  _ExtentY        =   529
                  _StockProps     =   78
                  Caption         =   "Casado"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSOption SSOption 
                  Height          =   300
                  Index           =   7
                  Left            =   1320
                  TabIndex        =   21
                  Top             =   240
                  Width           =   1092
                  _Version        =   65536
                  _ExtentX        =   1926
                  _ExtentY        =   529
                  _StockProps     =   78
                  Caption         =   "Viúvo"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSOption SSOption 
                  Height          =   300
                  Index           =   8
                  Left            =   1320
                  TabIndex        =   22
                  Top             =   600
                  Width           =   1140
                  _Version        =   65536
                  _ExtentX        =   2011
                  _ExtentY        =   529
                  _StockProps     =   78
                  Caption         =   "Separado"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSOption SSOption 
                  Height          =   300
                  Index           =   9
                  Left            =   2520
                  TabIndex        =   23
                  Top             =   240
                  Width           =   1536
                  _Version        =   65536
                  _ExtentX        =   2709
                  _ExtentY        =   529
                  _StockProps     =   78
                  Caption         =   "União irregular"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin Threed.SSFrame SSFrame 
               Height          =   828
               Index           =   5
               Left            =   1990
               TabIndex        =   167
               Top             =   2112
               Width           =   3180
               _Version        =   65536
               _ExtentX        =   5609
               _ExtentY        =   1461
               _StockProps     =   14
               Caption         =   "Data"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   3
               Begin MSMask.MaskEdBox MaskCaixa 
                  Height          =   252
                  Index           =   1
                  Left            =   360
                  TabIndex        =   14
                  Top             =   420
                  Width           =   1008
                  _ExtentX        =   1778
                  _ExtentY        =   445
                  _Version        =   327681
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MaskCaixa 
                  Bindings        =   "Form1.frx":0136
                  Height          =   252
                  Index           =   2
                  Left            =   1800
                  TabIndex        =   15
                  Top             =   420
                  Width           =   1008
                  _ExtentX        =   1778
                  _ExtentY        =   445
                  _Version        =   327681
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Falecimento"
                  Height          =   192
                  Index           =   11
                  Left            =   1800
                  TabIndex        =   169
                  Top             =   228
                  Width           =   888
               End
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Nascimento"
                  Height          =   192
                  Index           =   10
                  Left            =   360
                  TabIndex        =   168
                  Top             =   228
                  Width           =   864
               End
            End
            Begin Threed.SSFrame SSFrame 
               Height          =   3576
               Index           =   3
               Left            =   5395
               TabIndex        =   175
               Top             =   0
               Width           =   3864
               _Version        =   65536
               _ExtentX        =   6816
               _ExtentY        =   6308
               _StockProps     =   14
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   3
               Begin ComctlLib.TreeView TreeView1 
                  Height          =   2976
                  Left            =   180
                  TabIndex        =   247
                  Top             =   216
                  Width           =   3504
                  _ExtentX        =   6181
                  _ExtentY        =   5249
                  _Version        =   327682
                  LabelEdit       =   1
                  Style           =   7
                  Appearance      =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSOption SSOption 
                  Height          =   396
                  Index           =   4
                  Left            =   1968
                  TabIndex        =   177
                  Top             =   3156
                  Width           =   1092
                  _Version        =   65536
                  _ExtentX        =   1926
                  _ExtentY        =   699
                  _StockProps     =   78
                  Caption         =   "Procriação"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSOption SSOption 
                  Height          =   396
                  Index           =   3
                  Left            =   816
                  TabIndex        =   176
                  Top             =   3156
                  Width           =   1092
                  _Version        =   65536
                  _ExtentX        =   1926
                  _ExtentY        =   699
                  _StockProps     =   78
                  Caption         =   "Orientação"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   -1  'True
               End
            End
            Begin Threed.SSOption SSOption 
               Height          =   240
               Index           =   11
               Left            =   144
               TabIndex        =   250
               ToolTipText     =   "Nome Preferido"
               Top             =   1704
               Width           =   204
               _Version        =   65536
               _ExtentX        =   360
               _ExtentY        =   423
               _StockProps     =   78
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption SSOption 
               Height          =   240
               Index           =   10
               Left            =   144
               TabIndex        =   251
               ToolTipText     =   "Nome Preferido"
               Top             =   1128
               Width           =   204
               _Version        =   65536
               _ExtentX        =   360
               _ExtentY        =   423
               _StockProps     =   78
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Clã:"
               Height          =   192
               Index           =   13
               Left            =   2960
               TabIndex        =   166
               Top             =   3120
               Width           =   288
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Observação: "
               Height          =   240
               Index           =   14
               Left            =   4560
               TabIndex        =   164
               Top             =   3684
               Width           =   1212
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Nome Indígena:"
               Height          =   192
               Index           =   8
               Left            =   408
               TabIndex        =   85
               Top             =   912
               Width           =   1140
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Nome Nacional:"
               Height          =   192
               Index           =   9
               Left            =   408
               TabIndex        =   84
               Top             =   1488
               Width           =   1176
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Lugar de Nasc."
               Height          =   192
               Index           =   12
               Left            =   168
               TabIndex        =   83
               Top             =   3120
               Width           =   1068
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Lugar que mora"
               Height          =   192
               Index           =   7
               Left            =   1525
               TabIndex        =   82
               Top             =   324
               Width           =   1128
            End
            Begin VB.Label Label 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Casa N°:"
               Height          =   192
               Index           =   6
               Left            =   168
               TabIndex        =   81
               Top             =   336
               Width           =   636
            End
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   5076
            Index           =   8
            Left            =   -74904
            TabIndex        =   86
            Top             =   948
            Width           =   9264
            _Version        =   65536
            _ExtentX        =   16341
            _ExtentY        =   8954
            _StockProps     =   14
            Caption         =   "Relacionamentos"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
            Begin ComctlLib.ListView ListView1 
               Height          =   1260
               Left            =   132
               TabIndex        =   42
               Top             =   2196
               Width           =   5280
               _ExtentX        =   9313
               _ExtentY        =   2223
               View            =   2
               Arrange         =   1
               LabelEdit       =   1
               Sorted          =   -1  'True
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               _Version        =   327682
               Icons           =   "ImageList2"
               SmallIcons      =   "ImageList2"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin Threed.SSFrame SSFrame 
               Height          =   5076
               Index           =   0
               Left            =   5520
               TabIndex        =   256
               Top             =   0
               Width           =   3744
               _Version        =   65536
               _ExtentX        =   6604
               _ExtentY        =   8954
               _StockProps     =   14
               ForeColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   3
               Begin Threed.SSCommand SSCommand 
                  Height          =   252
                  Index           =   7
                  Left            =   1020
                  TabIndex        =   48
                  Top             =   4728
                  Width           =   840
                  _Version        =   65536
                  _ExtentX        =   1482
                  _ExtentY        =   445
                  _StockProps     =   78
                  Caption         =   "Apagar"
               End
               Begin Threed.SSCommand SSCommand 
                  Height          =   252
                  Index           =   8
                  Left            =   2076
                  TabIndex        =   49
                  Top             =   4728
                  Width           =   840
                  _Version        =   65536
                  _ExtentX        =   1482
                  _ExtentY        =   445
                  _StockProps     =   78
                  Caption         =   "Imprimir"
               End
               Begin ComctlLib.TreeView TreeView2 
                  Height          =   4296
                  Left            =   120
                  TabIndex        =   47
                  Top             =   336
                  Width           =   3492
                  _ExtentX        =   6160
                  _ExtentY        =   7578
                  _Version        =   327682
                  LabelEdit       =   1
                  Sorted          =   -1  'True
                  Style           =   7
                  Appearance      =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  Caption         =   "Famílias:"
                  Height          =   192
                  Index           =   22
                  Left            =   132
                  TabIndex        =   257
                  Top             =   144
                  Width           =   648
               End
            End
            Begin VB.PictureBox Membro_Família 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               Height          =   192
               Index           =   2
               Left            =   2892
               Picture         =   "Form1.frx":014B
               ScaleHeight     =   192
               ScaleWidth      =   216
               TabIndex        =   255
               Top             =   540
               Width           =   216
            End
            Begin VB.PictureBox Membro_Família 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               Height          =   192
               Index           =   1
               Left            =   2400
               Picture         =   "Form1.frx":028D
               ScaleHeight     =   192
               ScaleWidth      =   216
               TabIndex        =   254
               Top             =   540
               Width           =   216
            End
            Begin VB.PictureBox Membro_Família 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               Height          =   192
               Index           =   3
               Left            =   2652
               Picture         =   "Form1.frx":03CF
               ScaleHeight     =   192
               ScaleWidth      =   216
               TabIndex        =   253
               Top             =   960
               Width           =   216
            End
            Begin VB.TextBox Text 
               DataField       =   "Obs"
               DataSource      =   "DBcp_Casais"
               Height          =   924
               Index           =   0
               Left            =   144
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   43
               Top             =   3708
               Width           =   5280
            End
            Begin Threed.SSCommand SSCommand 
               Height          =   252
               Index           =   33
               Left            =   2400
               TabIndex        =   45
               Top             =   4740
               Width           =   840
               _Version        =   65536
               _ExtentX        =   1482
               _ExtentY        =   445
               _StockProps     =   78
               Caption         =   "Editar"
            End
            Begin Threed.SSCommand SSCommand 
               Height          =   252
               Index           =   32
               Left            =   1392
               TabIndex        =   44
               Top             =   4740
               Width           =   840
               _Version        =   65536
               _ExtentX        =   1482
               _ExtentY        =   445
               _StockProps     =   78
               Caption         =   "Novo"
            End
            Begin Threed.SSCheck Nome_Nac 
               Height          =   204
               Index           =   2
               Left            =   2592
               TabIndex        =   39
               Top             =   1536
               Width           =   876
               _Version        =   65536
               _ExtentX        =   1545
               _ExtentY        =   360
               _StockProps     =   78
               Caption         =   "Nacional"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSCheck Nome_Nac 
               Height          =   204
               Index           =   1
               Left            =   4440
               TabIndex        =   37
               Top             =   816
               Width           =   876
               _Version        =   65536
               _ExtentX        =   1545
               _ExtentY        =   360
               _StockProps     =   78
               Caption         =   "Nacional"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSCheck Nome_Nac 
               Height          =   204
               Index           =   0
               Left            =   192
               TabIndex        =   35
               Top             =   816
               Width           =   876
               _Version        =   65536
               _ExtentX        =   1545
               _ExtentY        =   360
               _StockProps     =   78
               Caption         =   "Nacional"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDBCtls.DBCombo DBCombo 
               Bindings        =   "Form1.frx":0511
               DataSource      =   "DBcp_Casais"
               Height          =   288
               Index           =   13
               Left            =   1356
               TabIndex        =   38
               ToolTipText     =   "teste"
               Top             =   1200
               Width           =   3180
               _ExtentX        =   5609
               _ExtentY        =   508
               _Version        =   327681
               Enabled         =   0   'False
               MatchEntry      =   -1  'True
               Style           =   2
               BackColor       =   16777215
               ForeColor       =   0
               ListField       =   "Nome_Ind"
               BoundColumn     =   "ID_Ego"
               Text            =   "DBCombo(13)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSCommand SSCommand 
               Height          =   252
               Index           =   4
               Left            =   1008
               TabIndex        =   40
               Top             =   1704
               Width           =   840
               _Version        =   65536
               _ExtentX        =   1482
               _ExtentY        =   445
               _StockProps     =   78
               Caption         =   "Inserir"
               Enabled         =   0   'False
            End
            Begin Threed.SSCommand SSCommand 
               Height          =   252
               Index           =   5
               Left            =   4032
               TabIndex        =   41
               Top             =   1704
               Width           =   840
               _Version        =   65536
               _ExtentX        =   1482
               _ExtentY        =   445
               _StockProps     =   78
               Caption         =   "Apagar"
               Enabled         =   0   'False
            End
            Begin MSDBCtls.DBCombo DBCombo 
               Bindings        =   "Form1.frx":0532
               DataField       =   "ID_Conj2"
               DataSource      =   "DBcp_Casais"
               Height          =   288
               Index           =   12
               Left            =   3204
               TabIndex        =   36
               Top             =   480
               Width           =   2208
               _ExtentX        =   3895
               _ExtentY        =   508
               _Version        =   327681
               Enabled         =   0   'False
               MatchEntry      =   -1  'True
               Style           =   2
               BackColor       =   16777215
               ForeColor       =   0
               ListField       =   "Nome_Ind"
               BoundColumn     =   "ID_Ego"
               Text            =   "DBCombo(12)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDBCtls.DBCombo DBCombo 
               Bindings        =   "Form1.frx":054E
               DataField       =   "ID_Conj1"
               DataSource      =   "DBcp_Casais"
               Height          =   288
               Index           =   11
               Left            =   168
               TabIndex        =   34
               Top             =   480
               Width           =   2172
               _ExtentX        =   3831
               _ExtentY        =   508
               _Version        =   327681
               Enabled         =   0   'False
               MatchEntry      =   -1  'True
               Style           =   2
               BackColor       =   16777215
               ForeColor       =   0
               ListField       =   "Nome_Ind"
               BoundColumn     =   "ID_Ego"
               Text            =   "DBCombo(11)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSCommand SSCommand 
               Height          =   252
               Index           =   6
               Left            =   3396
               TabIndex        =   46
               Top             =   4740
               Width           =   840
               _Version        =   65536
               _ExtentX        =   1482
               _ExtentY        =   445
               _StockProps     =   78
               Caption         =   "Gravar"
               Enabled         =   0   'False
            End
            Begin AniBtn.AniPushButton AniButton 
               Height          =   252
               Left            =   2616
               TabIndex        =   252
               Top             =   504
               Width           =   276
               _Version        =   65536
               _ExtentX        =   487
               _ExtentY        =   444
               _StockProps     =   111
               BackColor       =   -2147483633
               Picture         =   "Form1.frx":0578
               Cycle           =   1
               TextXpos        =   0
               TextYpos        =   0
               PictDrawMode    =   2
               ButtonVersion   =   1024
               ClearFirst      =   -1  'True
               HideFocusBox    =   -1  'True
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Observação:"
               Height          =   192
               Index           =   72
               Left            =   132
               TabIndex        =   248
               Top             =   3504
               Width           =   936
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00000000&
               X1              =   2748
               X2              =   2748
               Y1              =   696
               Y2              =   1068
            End
            Begin VB.Label Figura 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00004000&
               BackStyle       =   0  'Transparent
               Caption         =   "é"
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   13.8
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   312
               Index           =   2
               Left            =   4344
               TabIndex        =   91
               Top             =   1896
               Width           =   252
            End
            Begin VB.Label Figura 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00004000&
               BackStyle       =   0  'Transparent
               Caption         =   "ê"
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   13.8
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   312
               Index           =   1
               Left            =   1296
               TabIndex        =   90
               Top             =   1944
               Width           =   252
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Filhos:"
               Height          =   192
               Index           =   23
               Left            =   1368
               TabIndex        =   89
               Top             =   1008
               Width           =   480
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Mãe biológica:"
               Height          =   192
               Index           =   21
               Left            =   3192
               TabIndex        =   88
               Top             =   288
               Width           =   1068
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Pai biológico:"
               Height          =   192
               Index           =   20
               Left            =   168
               TabIndex        =   87
               Top             =   288
               Width           =   984
            End
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   912
            Index           =   7
            Left            =   -74916
            TabIndex        =   179
            Top             =   12
            Width           =   9264
            _Version        =   65536
            _ExtentX        =   16341
            _ExtentY        =   1609
            _StockProps     =   14
            Caption         =   "Dados de controle"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
            Begin VB.ComboBox Combo 
               Enabled         =   0   'False
               Height          =   288
               Index           =   15
               ItemData        =   "Form1.frx":085C
               Left            =   1380
               List            =   "Form1.frx":085E
               Sorted          =   -1  'True
               TabIndex        =   30
               Top             =   480
               Width           =   2652
            End
            Begin VB.TextBox Text 
               DataField       =   "Pg_Anota"
               DataSource      =   "DBcp_Casais"
               Height          =   300
               Index           =   5
               Left            =   8544
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   480
               Width           =   588
            End
            Begin VB.TextBox Text 
               DataField       =   "ID_Anota"
               DataSource      =   "DBcp_Casais"
               Height          =   300
               Index           =   4
               Left            =   7728
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   32
               Top             =   480
               Width           =   732
            End
            Begin MSMask.MaskEdBox MaskCaixa 
               DataField       =   "Data"
               DataSource      =   "DBcp_Casais"
               Height          =   288
               Index           =   3
               Left            =   168
               TabIndex        =   29
               Top             =   480
               Width           =   996
               _ExtentX        =   1757
               _ExtentY        =   508
               _Version        =   327681
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSDBCtls.DBCombo DBCombo 
               Bindings        =   "Form1.frx":0860
               DataField       =   "ID_Pesquisador"
               DataSource      =   "DBcp_Casais"
               Height          =   288
               Index           =   10
               Left            =   4296
               TabIndex        =   31
               Top             =   480
               Width           =   3084
               _ExtentX        =   5440
               _ExtentY        =   508
               _Version        =   327681
               Enabled         =   0   'False
               BackColor       =   16777215
               ForeColor       =   0
               ListField       =   "Pesq"
               BoundColumn     =   "ID_Pesquisador"
               Text            =   "DBCombo(10)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Data:"
               Height          =   192
               Index           =   15
               Left            =   168
               TabIndex        =   184
               Top             =   288
               Width           =   396
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Ajudante:"
               Height          =   192
               Index           =   16
               Left            =   1392
               TabIndex        =   183
               Top             =   288
               Width           =   672
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Pesquisador:"
               Height          =   192
               Index           =   17
               Left            =   4284
               TabIndex        =   182
               Top             =   288
               Width           =   960
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Caderno:"
               Height          =   192
               Index           =   18
               Left            =   7728
               TabIndex        =   181
               Top             =   288
               Width           =   660
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Página:"
               Height          =   192
               Index           =   19
               Left            =   8544
               TabIndex        =   180
               Top             =   288
               Width           =   552
            End
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   5964
            Index           =   9
            Left            =   -74916
            TabIndex        =   185
            Top             =   12
            Width           =   9264
            _Version        =   65536
            _ExtentX        =   16341
            _ExtentY        =   10520
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   21
               ItemData        =   "Form1.frx":087B
               Left            =   2628
               List            =   "Form1.frx":087D
               Sorted          =   -1  'True
               TabIndex        =   61
               Top             =   3132
               Width           =   2184
            End
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   20
               ItemData        =   "Form1.frx":087F
               Left            =   168
               List            =   "Form1.frx":0881
               Sorted          =   -1  'True
               TabIndex        =   60
               Top             =   3132
               Width           =   2364
            End
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   19
               ItemData        =   "Form1.frx":0883
               Left            =   168
               List            =   "Form1.frx":0885
               Sorted          =   -1  'True
               TabIndex        =   53
               Top             =   1548
               Width           =   4668
            End
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   18
               ItemData        =   "Form1.frx":0887
               Left            =   168
               List            =   "Form1.frx":0889
               Sorted          =   -1  'True
               TabIndex        =   52
               Top             =   972
               Width           =   4668
            End
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   17
               ItemData        =   "Form1.frx":088B
               Left            =   1416
               List            =   "Form1.frx":088D
               Sorted          =   -1  'True
               TabIndex        =   51
               Top             =   396
               Width           =   3420
            End
            Begin VB.ComboBox Combo 
               Height          =   288
               Index           =   16
               ItemData        =   "Form1.frx":088F
               Left            =   168
               List            =   "Form1.frx":0891
               Sorted          =   -1  'True
               TabIndex        =   50
               Top             =   396
               Width           =   1116
            End
            Begin VB.ListBox List 
               Height          =   5040
               Index           =   4
               Left            =   5040
               Sorted          =   -1  'True
               TabIndex        =   77
               Top             =   396
               Width           =   4092
            End
            Begin VB.TextBox Text 
               Height          =   288
               Index           =   6
               Left            =   168
               ScrollBars      =   2  'Vertical
               TabIndex        =   75
               Top             =   5496
               Width           =   4644
            End
            Begin Threed.SSFrame SSFrame 
               Height          =   924
               Index           =   10
               Left            =   168
               TabIndex        =   54
               Top             =   1908
               Width           =   1476
               _Version        =   65536
               _ExtentX        =   2603
               _ExtentY        =   1630
               _StockProps     =   14
               Caption         =   "Sexo"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   3
               Begin Threed.SSCheck SSCheck 
                  Height          =   252
                  Index           =   2
                  Left            =   288
                  TabIndex        =   56
                  Top             =   540
                  Width           =   972
                  _Version        =   65536
                  _ExtentX        =   1714
                  _ExtentY        =   444
                  _StockProps     =   78
                  Caption         =   "Feminino"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck SSCheck 
                  Height          =   252
                  Index           =   1
                  Left            =   288
                  TabIndex        =   55
                  Top             =   264
                  Width           =   972
                  _Version        =   65536
                  _ExtentX        =   1715
                  _ExtentY        =   445
                  _StockProps     =   78
                  Caption         =   "Masculino"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin Threed.SSFrame SSFrame 
               Height          =   1644
               Index           =   12
               Left            =   168
               TabIndex        =   62
               Top             =   3576
               Width           =   1596
               _Version        =   65536
               _ExtentX        =   2815
               _ExtentY        =   2900
               _StockProps     =   14
               Caption         =   "Estado Civil"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   3
               Begin Threed.SSCheck SSCheck 
                  Height          =   216
                  Index           =   8
                  Left            =   192
                  TabIndex        =   66
                  Top             =   1008
                  Width           =   1212
                  _Version        =   65536
                  _ExtentX        =   2138
                  _ExtentY        =   381
                  _StockProps     =   78
                  Caption         =   "Separado"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck SSCheck 
                  Height          =   252
                  Index           =   9
                  Left            =   192
                  TabIndex        =   67
                  Top             =   1248
                  Width           =   1356
                  _Version        =   65536
                  _ExtentX        =   2392
                  _ExtentY        =   445
                  _StockProps     =   78
                  Caption         =   "União irregular"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck SSCheck 
                  Height          =   252
                  Index           =   7
                  Left            =   192
                  TabIndex        =   65
                  Top             =   768
                  Width           =   1080
                  _Version        =   65536
                  _ExtentX        =   1905
                  _ExtentY        =   444
                  _StockProps     =   78
                  Caption         =   "Viúvo"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck SSCheck 
                  Height          =   252
                  Index           =   6
                  Left            =   192
                  TabIndex        =   64
                  Top             =   528
                  Width           =   1116
                  _Version        =   65536
                  _ExtentX        =   1968
                  _ExtentY        =   444
                  _StockProps     =   78
                  Caption         =   "Casado"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck SSCheck 
                  Height          =   252
                  Index           =   5
                  Left            =   192
                  TabIndex        =   63
                  Top             =   288
                  Width           =   1092
                  _Version        =   65536
                  _ExtentX        =   1926
                  _ExtentY        =   444
                  _StockProps     =   78
                  Caption         =   "Solteiro"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin Threed.SSFrame SSFrame 
               Height          =   1620
               Index           =   11
               Left            =   1920
               TabIndex        =   68
               Top             =   3600
               Width           =   2892
               _Version        =   65536
               _ExtentX        =   5101
               _ExtentY        =   2857
               _StockProps     =   14
               Caption         =   "Período"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   3
               Begin Threed.SSCheck SSCheck 
                  Height          =   252
                  Index           =   4
                  Left            =   168
                  TabIndex        =   72
                  Top             =   972
                  Width           =   1488
                  _Version        =   65536
                  _ExtentX        =   2625
                  _ExtentY        =   444
                  _StockProps     =   78
                  Caption         =   "Falescimento:"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck SSCheck 
                  Height          =   252
                  Index           =   3
                  Left            =   168
                  TabIndex        =   69
                  Top             =   288
                  Width           =   1476
                  _Version        =   65536
                  _ExtentX        =   2603
                  _ExtentY        =   444
                  _StockProps     =   78
                  Caption         =   "Nascimento:"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin MSMask.MaskEdBox MaskCaixa 
                  Height          =   288
                  Index           =   4
                  Left            =   240
                  TabIndex        =   70
                  Top             =   540
                  Width           =   996
                  _ExtentX        =   1757
                  _ExtentY        =   508
                  _Version        =   327681
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MaskCaixa 
                  Height          =   288
                  Index           =   5
                  Left            =   1680
                  TabIndex        =   71
                  Top             =   540
                  Width           =   996
                  _ExtentX        =   1757
                  _ExtentY        =   508
                  _Version        =   327681
                  MaxLength       =   10
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MaskCaixa 
                  Height          =   288
                  Index           =   6
                  Left            =   252
                  TabIndex        =   73
                  Top             =   1212
                  Width           =   996
                  _ExtentX        =   1757
                  _ExtentY        =   508
                  _Version        =   327681
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MaskCaixa 
                  Height          =   288
                  Index           =   7
                  Left            =   1680
                  TabIndex        =   74
                  Top             =   1212
                  Width           =   996
                  _ExtentX        =   1757
                  _ExtentY        =   508
                  _Version        =   327681
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  Caption         =   "até:"
                  Height          =   192
                  Index           =   0
                  Left            =   1356
                  TabIndex        =   249
                  Top             =   1212
                  Width           =   264
               End
               Begin VB.Label Label 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  Caption         =   "até:"
                  Height          =   192
                  Index           =   29
                  Left            =   1356
                  TabIndex        =   186
                  Top             =   564
                  Width           =   264
               End
            End
            Begin Threed.SSFrame SSFrame 
               Height          =   924
               Index           =   13
               Left            =   1956
               TabIndex        =   57
               Top             =   1920
               Width           =   2856
               _Version        =   65536
               _ExtentX        =   5038
               _ExtentY        =   1630
               _StockProps     =   14
               Caption         =   "Família"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   3
               Begin Threed.SSCheck SSCheck 
                  Height          =   252
                  Index           =   10
                  Left            =   216
                  TabIndex        =   58
                  Top             =   348
                  Width           =   1236
                  _Version        =   65536
                  _ExtentX        =   2180
                  _ExtentY        =   444
                  _StockProps     =   78
                  Caption         =   "Orientação"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin Threed.SSCheck SSCheck 
                  Height          =   252
                  Index           =   11
                  Left            =   1632
                  TabIndex        =   59
                  Top             =   350
                  Width           =   1104
                  _Version        =   65536
                  _ExtentX        =   1947
                  _ExtentY        =   444
                  _StockProps     =   78
                  Caption         =   "Procriação"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   7.8
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin Threed.SSCommand SSCommand 
               Height          =   252
               Index           =   9
               Left            =   6168
               TabIndex        =   76
               Top             =   5592
               Width           =   840
               _Version        =   65536
               _ExtentX        =   1482
               _ExtentY        =   445
               _StockProps     =   78
               Caption         =   "Buscar"
            End
            Begin Threed.SSCommand SSCommand 
               Height          =   252
               Index           =   10
               Left            =   7344
               TabIndex        =   78
               Top             =   5592
               Width           =   840
               _Version        =   65536
               _ExtentX        =   1482
               _ExtentY        =   445
               _StockProps     =   78
               Caption         =   "Imprimir"
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Casa  n° :"
               Height          =   192
               Index           =   24
               Left            =   168
               TabIndex        =   194
               Top             =   192
               Width           =   672
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Lugar que mora:"
               Height          =   192
               Index           =   25
               Left            =   1440
               TabIndex        =   193
               Top             =   192
               Width           =   1416
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Lugar de Nascimento:"
               Height          =   192
               Index           =   30
               Left            =   168
               TabIndex        =   192
               Top             =   2940
               Width           =   1572
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Nome Nacional:"
               Height          =   192
               Index           =   28
               Left            =   168
               TabIndex        =   191
               Top             =   1356
               Width           =   1404
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Nome Indígena:"
               Height          =   192
               Index           =   27
               Left            =   168
               TabIndex        =   190
               Top             =   780
               Width           =   1380
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Texto na Observação: "
               Height          =   192
               Index           =   32
               Left            =   168
               TabIndex        =   189
               Top             =   5304
               Width           =   1992
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Clã:"
               Height          =   192
               Index           =   31
               Left            =   2652
               TabIndex        =   188
               Top             =   2940
               Width           =   288
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Resultados:"
               Height          =   192
               Index           =   26
               Left            =   5040
               TabIndex        =   187
               Top             =   204
               Width           =   1092
            End
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   1
            Left            =   5088
            TabIndex        =   26
            Top             =   6168
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Editar"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   3
            Left            =   7248
            TabIndex        =   28
            Top             =   6168
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Imprimir"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   2
            Left            =   5988
            TabIndex        =   27
            Top             =   6168
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Gravar"
            Enabled         =   0   'False
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   0
            Left            =   4164
            TabIndex        =   25
            Top             =   6168
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Novo"
         End
         Begin VB.Label Total 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   252
            Left            =   8448
            TabIndex        =   246
            Top             =   6168
            Width           =   876
         End
         Begin VB.Label Papel 
            Height          =   108
            Index           =   3
            Left            =   -75000
            TabIndex        =   178
            Top             =   -96
            Width           =   1344
         End
         Begin VB.Label Papel 
            Height          =   108
            Index           =   1
            Left            =   12
            TabIndex        =   150
            Top             =   -48
            Width           =   1344
         End
         Begin VB.Label Papel 
            Height          =   108
            Index           =   2
            Left            =   -74988
            TabIndex        =   149
            Top             =   -96
            Width           =   1344
         End
      End
      Begin TabDlg.SSTab SSTab 
         Height          =   6468
         Index           =   3
         Left            =   0
         TabIndex        =   141
         Top             =   336
         Width           =   9470
         _ExtentX        =   16701
         _ExtentY        =   11409
         _Version        =   327681
         TabOrientation  =   1
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   600
         BackColor       =   -2147483636
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Geral"
         TabPicture(0)   =   "Form1.frx":0893
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Papel(4)"
         Tab(0).Control(1)=   "Label(34)"
         Tab(0).Control(2)=   "Label(36)"
         Tab(0).Control(3)=   "Label(35)"
         Tab(0).Control(4)=   "Label(38)"
         Tab(0).Control(5)=   "Label(37)"
         Tab(0).Control(6)=   "Label(33)"
         Tab(0).Control(7)=   "HScroll1"
         Tab(0).Control(8)=   "SSPanel(0)"
         Tab(0).Control(9)=   "DBGrid(1)"
         Tab(0).Control(10)=   "SSCommand(11)"
         Tab(0).Control(11)=   "SSCommand(12)"
         Tab(0).Control(12)=   "SSFrame(15)"
         Tab(0).Control(13)=   "SSFrame(14)"
         Tab(0).Control(14)=   "Combo(2)"
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "Planejamento"
         TabPicture(1)   =   "Form1.frx":08AF
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Papel(5)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label(40)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label(41)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label(39)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label(74)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "DBCombo(0)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "SSCommand(13)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "SSFrame(17)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "SSFrame(16)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Combo(4)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "SSPanel(2)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "HScroll_Parente"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "VScroll_Parente"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "List(5)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "List(6)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).ControlCount=   15
         Begin VB.ListBox List 
            Height          =   4656
            Index           =   6
            Left            =   2520
            TabIndex        =   277
            Top             =   888
            Width           =   2268
         End
         Begin VB.ListBox List 
            Height          =   4656
            Index           =   5
            Left            =   120
            TabIndex        =   276
            Top             =   888
            Width           =   2268
         End
         Begin VB.VScrollBar VScroll_Parente 
            Height          =   4430
            LargeChange     =   2000
            Left            =   9000
            SmallChange     =   1000
            TabIndex        =   275
            Top             =   888
            Width           =   252
         End
         Begin VB.HScrollBar HScroll_Parente 
            Height          =   252
            LargeChange     =   2000
            Left            =   4920
            SmallChange     =   1000
            TabIndex        =   274
            Top             =   5300
            Width           =   4092
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   4428
            Index           =   2
            Left            =   4920
            TabIndex        =   272
            Top             =   888
            Width           =   4092
            _Version        =   65536
            _ExtentX        =   7218
            _ExtentY        =   7810
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Begin VB.PictureBox Mapa_Parente 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               FillStyle       =   0  'Solid
               FontTransparent =   0   'False
               ForeColor       =   &H80000008&
               Height          =   2000
               Left            =   120
               ScaleHeight     =   2004
               ScaleWidth      =   6204
               TabIndex        =   273
               Top             =   1176
               Width           =   6204
            End
         End
         Begin VB.ComboBox Combo 
            Height          =   288
            Index           =   2
            ItemData        =   "Form1.frx":08CB
            Left            =   -73176
            List            =   "Form1.frx":08CD
            Style           =   2  'Dropdown List
            TabIndex        =   196
            Top             =   288
            Width           =   2820
         End
         Begin VB.ComboBox Combo 
            Height          =   288
            Index           =   4
            Left            =   3624
            Style           =   2  'Dropdown List
            TabIndex        =   195
            Top             =   288
            Width           =   1980
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   588
            Index           =   16
            Left            =   5760
            TabIndex        =   198
            Top             =   96
            Width           =   1212
            _Version        =   65536
            _ExtentX        =   2138
            _ExtentY        =   1037
            _StockProps     =   14
            Caption         =   "Sexo do Ego"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSOption SSOption 
               Height          =   300
               Index           =   16
               Left            =   192
               TabIndex        =   200
               Top             =   240
               Width           =   444
               _Version        =   65536
               _ExtentX        =   783
               _ExtentY        =   529
               _StockProps     =   78
               Caption         =   "M"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   -1  'True
            End
            Begin Threed.SSOption SSOption 
               Height          =   300
               Index           =   17
               Left            =   720
               TabIndex        =   199
               Top             =   240
               Width           =   396
               _Version        =   65536
               _ExtentX        =   699
               _ExtentY        =   529
               _StockProps     =   78
               Caption         =   "F"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   588
            Index           =   17
            Left            =   7248
            TabIndex        =   201
            Top             =   96
            Width           =   1644
            _Version        =   65536
            _ExtentX        =   2900
            _ExtentY        =   1037
            _StockProps     =   14
            Caption         =   "Vista"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSOption SSOption 
               Height          =   300
               Index           =   19
               Left            =   912
               TabIndex        =   203
               Top             =   240
               Width           =   636
               _Version        =   65536
               _ExtentX        =   1122
               _ExtentY        =   529
               _StockProps     =   78
               Caption         =   "Mapa"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption SSOption 
               Height          =   300
               Index           =   18
               Left            =   240
               TabIndex        =   202
               Top             =   240
               Width           =   588
               _Version        =   65536
               _ExtentX        =   1037
               _ExtentY        =   529
               _StockProps     =   78
               Caption         =   "Lista"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   588
            Index           =   14
            Left            =   -69240
            TabIndex        =   204
            Top             =   96
            Width           =   1212
            _Version        =   65536
            _ExtentX        =   2138
            _ExtentY        =   1037
            _StockProps     =   14
            Caption         =   "Sexo do Ego"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSOption SSOption 
               Height          =   300
               Index           =   12
               Left            =   192
               TabIndex        =   206
               Top             =   240
               Width           =   444
               _Version        =   65536
               _ExtentX        =   783
               _ExtentY        =   529
               _StockProps     =   78
               Caption         =   "M"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   -1  'True
            End
            Begin Threed.SSOption SSOption 
               Height          =   300
               Index           =   13
               Left            =   720
               TabIndex        =   205
               Top             =   240
               Width           =   396
               _Version        =   65536
               _ExtentX        =   699
               _ExtentY        =   529
               _StockProps     =   78
               Caption         =   "F"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   588
            Index           =   15
            Left            =   -67752
            TabIndex        =   207
            Top             =   96
            Width           =   1884
            _Version        =   65536
            _ExtentX        =   3323
            _ExtentY        =   1037
            _StockProps     =   14
            Caption         =   "Vista"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSOption SSOption 
               Height          =   300
               Index           =   15
               Left            =   912
               TabIndex        =   209
               Top             =   240
               Width           =   876
               _Version        =   65536
               _ExtentX        =   1545
               _ExtentY        =   529
               _StockProps     =   78
               Caption         =   "Mapa"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin Threed.SSOption SSOption 
               Height          =   300
               Index           =   14
               Left            =   240
               TabIndex        =   208
               Top             =   240
               Width           =   588
               _Version        =   65536
               _ExtentX        =   1037
               _ExtentY        =   529
               _StockProps     =   78
               Caption         =   "Lista"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   -1  'True
            End
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   13
            Left            =   5376
            TabIndex        =   146
            Top             =   5736
            Width           =   1068
            _Version        =   65536
            _ExtentX        =   1884
            _ExtentY        =   444
            _StockProps     =   78
            Caption         =   "Imprimir"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   12
            Left            =   -74640
            TabIndex        =   143
            Top             =   456
            Width           =   1188
            _Version        =   65536
            _ExtentX        =   2096
            _ExtentY        =   444
            _StockProps     =   78
            Caption         =   "Imprimir"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   11
            Left            =   -74640
            TabIndex        =   142
            Top             =   96
            Width           =   1188
            _Version        =   65536
            _ExtentX        =   2096
            _ExtentY        =   444
            _StockProps     =   78
            Caption         =   "Novo Termo"
         End
         Begin MSDBCtls.DBCombo DBCombo 
            Bindings        =   "Form1.frx":08CF
            DataSource      =   "DBpa_Termos"
            Height          =   288
            Index           =   0
            Left            =   120
            TabIndex        =   271
            Top             =   288
            Width           =   3324
            _ExtentX        =   5863
            _ExtentY        =   508
            _Version        =   327681
            MatchEntry      =   -1  'True
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "Termo_Tec"
            BoundColumn     =   "Trilha"
            Text            =   "DBCombo(0)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDBGrid.DBGrid DBGrid 
            Bindings        =   "Form1.frx":08F3
            Height          =   4440
            Index           =   1
            Left            =   -74880
            OleObjectBlob   =   "Form1.frx":0909
            TabIndex        =   165
            Top             =   816
            Width           =   9216
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   4188
            Index           =   0
            Left            =   -74880
            TabIndex        =   267
            Top             =   816
            Width           =   9216
            _Version        =   65536
            _ExtentX        =   16256
            _ExtentY        =   7387
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Begin VB.PictureBox Mapa_Termo 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               FillStyle       =   0  'Solid
               FontTransparent =   0   'False
               ForeColor       =   &H80000008&
               Height          =   4188
               Left            =   120
               ScaleHeight     =   4188
               ScaleWidth      =   30804
               TabIndex        =   268
               Top             =   120
               Width           =   30800
            End
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   252
            LargeChange     =   2000
            Left            =   -74880
            SmallChange     =   1000
            TabIndex        =   269
            Top             =   5000
            Width           =   9204
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Parentes:"
            Height          =   192
            Index           =   74
            Left            =   2520
            TabIndex        =   278
            Top             =   624
            Width           =   684
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   288
            Index           =   33
            Left            =   -70320
            TabIndex        =   270
            Top             =   288
            Width           =   528
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Class. de Primo:"
            ForeColor       =   &H00000000&
            Height          =   192
            Index           =   37
            Left            =   -69744
            TabIndex        =   244
            Top             =   5544
            Width           =   1200
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BorderStyle     =   1  'Fixed Single
            Height          =   264
            Index           =   38
            Left            =   -68544
            TabIndex        =   243
            Top             =   5544
            Width           =   2820
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Class. de 1ª Geração Ascend.:"
            ForeColor       =   &H00000000&
            Height          =   192
            Index           =   35
            Left            =   -74904
            TabIndex        =   242
            Top             =   5544
            Width           =   2220
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            BorderStyle     =   1  'Fixed Single
            Height          =   264
            Index           =   36
            Left            =   -72648
            TabIndex        =   241
            Top             =   5544
            Width           =   2436
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Classificação:"
            Height          =   192
            Index           =   34
            Left            =   -73176
            TabIndex        =   197
            Top             =   96
            Width           =   1008
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Termo:"
            Height          =   192
            Index           =   39
            Left            =   120
            TabIndex        =   161
            Top             =   96
            Width           =   576
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Egos:"
            Height          =   192
            Index           =   41
            Left            =   120
            TabIndex        =   148
            Top             =   624
            Width           =   528
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Classificação:"
            Height          =   192
            Index           =   40
            Left            =   3624
            TabIndex        =   147
            Top             =   96
            Width           =   1080
            WordWrap        =   -1  'True
         End
         Begin VB.Label Papel 
            Height          =   60
            Index           =   5
            Left            =   1332
            TabIndex        =   145
            Top             =   0
            Width           =   1020
         End
         Begin VB.Label Papel 
            Height          =   108
            Index           =   4
            Left            =   -73668
            TabIndex        =   144
            Top             =   -48
            Width           =   1020
         End
      End
      Begin TabDlg.SSTab SSTab 
         Height          =   6468
         Index           =   4
         Left            =   -75000
         TabIndex        =   92
         Top             =   336
         Width           =   9470
         _ExtentX        =   16701
         _ExtentY        =   11409
         _Version        =   327681
         TabOrientation  =   1
         Style           =   1
         Tab             =   2
         TabHeight       =   600
         BackColor       =   -2147483636
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Editar"
         TabPicture(0)   =   "Form1.frx":14AA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Check(1)"
         Tab(0).Control(1)=   "SSFrame(18)"
         Tab(0).Control(2)=   "SSCommand(16)"
         Tab(0).Control(3)=   "SSCommand(15)"
         Tab(0).Control(4)=   "SSCommand(14)"
         Tab(0).Control(5)=   "Papel(6)"
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Analizar"
         TabPicture(1)   =   "Form1.frx":14C6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Check(2)"
         Tab(1).Control(1)=   "SSCommand(19)"
         Tab(1).Control(2)=   "DBList(2)"
         Tab(1).Control(3)=   "DBCombo(23)"
         Tab(1).Control(4)=   "SSCommand(18)"
         Tab(1).Control(5)=   "SSCommand(17)"
         Tab(1).Control(6)=   "Papel(7)"
         Tab(1).Control(7)=   "Label(50)"
         Tab(1).Control(8)=   "Label(49)"
         Tab(1).Control(9)=   "Label(48)"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "Planejamento"
         TabPicture(2)   =   "Form1.frx":14E2
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Papel(8)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label(51)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label(53)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Label(52)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "DBGrid(3)"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "SSCommand(20)"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "List(2)"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "List(3)"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "List(1)"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "Check(3)"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).ControlCount=   10
         Begin VB.CheckBox Check 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Caption         =   "Perguntas Confirmadas"
            Height          =   252
            Index           =   3
            Left            =   6960
            TabIndex        =   236
            Top             =   5856
            Width           =   1932
         End
         Begin VB.ListBox List 
            Height          =   240
            Index           =   1
            Left            =   144
            TabIndex        =   234
            Top             =   285
            Width           =   2748
         End
         Begin VB.ListBox List 
            Height          =   1008
            Index           =   3
            Left            =   144
            TabIndex        =   232
            Top             =   1152
            Width           =   8700
         End
         Begin VB.ListBox List 
            Height          =   240
            Index           =   2
            Left            =   2928
            TabIndex        =   231
            Top             =   288
            Width           =   5916
         End
         Begin VB.CheckBox Check 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Caption         =   "Completado"
            Height          =   252
            Index           =   2
            Left            =   -67305
            TabIndex        =   94
            Top             =   5850
            Width           =   1164
         End
         Begin VB.CheckBox Check 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Height          =   255
            Index           =   1
            Left            =   -67095
            TabIndex        =   93
            Top             =   5850
            Width           =   960
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   1404
            Index           =   18
            Left            =   -74916
            TabIndex        =   95
            Top             =   48
            Width           =   8844
            _Version        =   65536
            _ExtentX        =   15600
            _ExtentY        =   2477
            _StockProps     =   14
            Caption         =   "Dados de controle"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
            Begin VB.TextBox Text 
               Height          =   300
               Index           =   8
               Left            =   8112
               TabIndex        =   97
               Text            =   "Text3"
               Top             =   1008
               Width           =   588
            End
            Begin VB.TextBox Text 
               Height          =   300
               Index           =   7
               Left            =   7248
               TabIndex        =   96
               Text            =   "Text3"
               Top             =   1008
               Width           =   732
            End
            Begin MSDBCtls.DBCombo DBCombo 
               Height          =   288
               Index           =   20
               Left            =   156
               TabIndex        =   151
               Top             =   480
               Width           =   8544
               _ExtentX        =   15071
               _ExtentY        =   508
               _Version        =   327681
               BackColor       =   16777215
               ForeColor       =   0
               Text            =   "DBCombo(20)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDBCtls.DBCombo DBCombo 
               Height          =   288
               Index           =   21
               Left            =   1200
               TabIndex        =   102
               Top             =   1008
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   508
               _Version        =   327681
               BackColor       =   16777215
               ForeColor       =   0
               Text            =   "DBCombo(21)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDBCtls.DBCombo DBCombo 
               Height          =   288
               Index           =   22
               Left            =   4032
               TabIndex        =   101
               Top             =   1008
               Width           =   3084
               _ExtentX        =   5440
               _ExtentY        =   508
               _Version        =   327681
               BackColor       =   16777215
               ForeColor       =   0
               Text            =   "DBCombo(22)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Data:"
               Height          =   192
               Index           =   43
               Left            =   144
               TabIndex        =   105
               Top             =   816
               Width           =   396
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Título:"
               Height          =   192
               Index           =   42
               Left            =   144
               TabIndex        =   100
               Top             =   288
               Width           =   588
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Ajudante:"
               Height          =   192
               Index           =   44
               Left            =   1200
               TabIndex        =   104
               Top             =   816
               Width           =   672
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Pesquisador:"
               Height          =   192
               Index           =   45
               Left            =   4032
               TabIndex        =   103
               Top             =   816
               Width           =   960
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Caderno:"
               Height          =   192
               Index           =   46
               Left            =   7248
               TabIndex        =   99
               Top             =   816
               Width           =   660
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Página:"
               Height          =   192
               Index           =   47
               Left            =   8112
               TabIndex        =   98
               Top             =   816
               Width           =   552
            End
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   19
            Left            =   -69240
            TabIndex        =   245
            Top             =   5856
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Imprimir"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   20
            Left            =   5376
            TabIndex        =   237
            Top             =   5856
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Imprimir"
         End
         Begin MSDBGrid.DBGrid DBGrid 
            Height          =   720
            Index           =   3
            Left            =   144
            Negotiate       =   -1  'True
            OleObjectBlob   =   "Form1.frx":14FE
            TabIndex        =   229
            Top             =   3132
            Width           =   4800
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   16
            Left            =   -69192
            TabIndex        =   210
            Top             =   5856
            Width           =   828
            _Version        =   65536
            _ExtentX        =   1461
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Imprimir"
         End
         Begin MSDBCtls.DBList DBList 
            Height          =   432
            Index           =   2
            Left            =   -69288
            TabIndex        =   160
            Top             =   4668
            Width           =   3228
            _ExtentX        =   5694
            _ExtentY        =   762
            _Version        =   327681
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDBCtls.DBCombo DBCombo 
            Height          =   288
            Index           =   23
            Left            =   -74904
            TabIndex        =   156
            Top             =   288
            Width           =   5532
            _ExtentX        =   9758
            _ExtentY        =   508
            _Version        =   327681
            BackColor       =   16777215
            ForeColor       =   0
            Text            =   "DBCombo(23)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   18
            Left            =   -70440
            TabIndex        =   155
            Top             =   5856
            Width           =   828
            _Version        =   65536
            _ExtentX        =   1461
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Lista"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   255
            Index           =   17
            Left            =   -71685
            TabIndex        =   154
            Top             =   5850
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1461
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Gravar"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   15
            Left            =   -70440
            TabIndex        =   153
            Top             =   5856
            Width           =   828
            _Version        =   65536
            _ExtentX        =   1461
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Gravar"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   14
            Left            =   -71688
            TabIndex        =   152
            Top             =   5856
            Width           =   828
            _Version        =   65536
            _ExtentX        =   1461
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Novo"
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Acontecimento:"
            Height          =   192
            Index           =   52
            Left            =   2928
            TabIndex        =   235
            Top             =   96
            Width           =   1104
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Perguntas Gerais:"
            Height          =   192
            Index           =   53
            Left            =   144
            TabIndex        =   233
            Top             =   960
            Width           =   1284
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Pergunta:"
            Height          =   192
            Index           =   51
            Left            =   144
            TabIndex        =   230
            Top             =   96
            Width           =   684
         End
         Begin VB.Label Papel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   96
            Index           =   8
            Left            =   2304
            TabIndex        =   228
            Top             =   -48
            Width           =   1380
         End
         Begin VB.Label Papel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   144
            Index           =   6
            Left            =   -72696
            TabIndex        =   163
            Top             =   -48
            Width           =   1380
         End
         Begin VB.Label Papel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   192
            Index           =   7
            Left            =   -72696
            TabIndex        =   162
            Top             =   -96
            Width           =   1380
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Referências:"
            Height          =   192
            Index           =   50
            Left            =   -69288
            TabIndex        =   159
            Top             =   4476
            Width           =   912
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Esboço universal:"
            Height          =   192
            Index           =   49
            Left            =   -69288
            TabIndex        =   158
            Top             =   96
            Width           =   1284
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Título:"
            Height          =   192
            Index           =   48
            Left            =   -74904
            TabIndex        =   157
            Top             =   96
            Width           =   540
         End
      End
      Begin Threed.SSFrame SSFrame 
         Height          =   6180
         Index           =   25
         Left            =   -74844
         TabIndex        =   220
         Top             =   444
         Width           =   4032
         _Version        =   65536
         _ExtentX        =   7112
         _ExtentY        =   10901
         _StockProps     =   14
         Caption         =   "Informações"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "Form1.frx":1ED3
            Height          =   2472
            Left            =   156
            OleObjectBlob   =   "Form1.frx":1EEE
            TabIndex        =   258
            Top             =   3552
            Width           =   3672
         End
         Begin VB.TextBox Text 
            Height          =   1356
            Index           =   17
            Left            =   144
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   226
            Top             =   1680
            Width           =   3660
         End
         Begin VB.TextBox Text 
            Height          =   288
            Index           =   16
            Left            =   132
            TabIndex        =   225
            Top             =   1056
            Width           =   3660
         End
         Begin VB.TextBox Text 
            Height          =   288
            Index           =   15
            Left            =   132
            TabIndex        =   221
            Top             =   432
            Width           =   3660
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   31
            Left            =   1632
            TabIndex        =   222
            Top             =   5424
            Visible         =   0   'False
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "OK"
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Pesquisadores:"
            Height          =   216
            Index           =   71
            Left            =   156
            TabIndex        =   239
            Top             =   3204
            Width           =   1128
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Localização:"
            Height          =   192
            Index           =   70
            Left            =   132
            TabIndex        =   227
            Top             =   1488
            Width           =   900
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Aldeia:"
            Height          =   192
            Index           =   69
            Left            =   132
            TabIndex        =   224
            Top             =   864
            Width           =   504
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tribo:"
            Height          =   192
            Index           =   68
            Left            =   132
            TabIndex        =   223
            Top             =   240
            Width           =   420
         End
      End
      Begin Threed.SSFrame SSFrame 
         Height          =   6180
         Index           =   27
         Left            =   -70656
         TabIndex        =   240
         Top             =   444
         Width           =   4956
         _Version        =   65536
         _ExtentX        =   8742
         _ExtentY        =   10901
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox Combo 
            Height          =   288
            Index           =   7
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   259
            Top             =   444
            Width           =   2172
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   1260
            Index           =   26
            Left            =   204
            TabIndex        =   260
            Top             =   1032
            Width           =   2412
            _Version        =   65536
            _ExtentX        =   4255
            _ExtentY        =   2223
            _StockProps     =   14
            Caption         =   "Confirmações"
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
               Height          =   300
               Index           =   14
               Left            =   1824
               MaxLength       =   2
               TabIndex        =   262
               Top             =   720
               Width           =   300
            End
            Begin VB.TextBox Text 
               Height          =   300
               Index           =   13
               Left            =   1824
               MaxLength       =   2
               TabIndex        =   261
               Top             =   288
               Width           =   300
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Pergunta:"
               Height          =   192
               Index           =   66
               Left            =   1104
               TabIndex        =   264
               Top             =   336
               Width           =   684
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Termo de Parentesco:"
               Height          =   192
               Index           =   67
               Left            =   192
               TabIndex        =   263
               Top             =   768
               Width           =   1608
            End
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Interface em desenvolvimento."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1104
            Index           =   73
            Left            =   732
            TabIndex        =   266
            Top             =   3456
            Width           =   3756
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Idioma da Interface:"
            Height          =   192
            Index           =   65
            Left            =   228
            TabIndex        =   265
            Top             =   240
            Width           =   1404
         End
      End
      Begin Threed.SSFrame SSFrame 
         Height          =   6060
         Index           =   21
         Left            =   -74952
         TabIndex        =   106
         Top             =   540
         Width           =   8940
         _Version        =   65536
         _ExtentX        =   15769
         _ExtentY        =   10689
         _StockProps     =   14
         Caption         =   "Editando Descrição"
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         ShadowStyle     =   1
         Begin VB.TextBox Text 
            Height          =   300
            Index           =   10
            Left            =   96
            TabIndex        =   119
            Text            =   "Achando sangue no pátio"
            Top             =   432
            Width           =   4764
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   4188
            Index           =   22
            Left            =   4992
            TabIndex        =   125
            Top             =   1440
            Width           =   3852
            _Version        =   65536
            _ExtentX        =   6795
            _ExtentY        =   7387
            _StockProps     =   14
            Caption         =   "Mapa"
            ForeColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSCommand SSCommand 
               Height          =   255
               Index           =   23
               Left            =   630
               TabIndex        =   129
               Top             =   3840
               Width           =   780
               _Version        =   65536
               _ExtentX        =   1376
               _ExtentY        =   450
               _StockProps     =   78
               Caption         =   "Novo"
            End
            Begin Threed.SSCommand SSCommand 
               Height          =   255
               Index           =   24
               Left            =   1530
               TabIndex        =   128
               Top             =   3840
               Width           =   780
               _Version        =   65536
               _ExtentX        =   1376
               _ExtentY        =   450
               _StockProps     =   78
               Caption         =   "Editar"
            End
            Begin Threed.SSCommand SSCommand 
               Height          =   255
               Index           =   25
               Left            =   2490
               TabIndex        =   127
               Top             =   3840
               Width           =   780
               _Version        =   65536
               _ExtentX        =   1376
               _ExtentY        =   450
               _StockProps     =   78
               Caption         =   "Imprimir"
            End
            Begin MSDBCtls.DBCombo DBCombo 
               Height          =   288
               Index           =   26
               Left            =   96
               TabIndex        =   126
               Top             =   288
               Width           =   3660
               _ExtentX        =   6456
               _ExtentY        =   508
               _Version        =   327681
               BackColor       =   16777215
               ForeColor       =   0
               Text            =   "DBCombo(26)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSPanel SSPanel 
            Height          =   348
            Index           =   1
            Left            =   96
            TabIndex        =   211
            Top             =   1392
            Width           =   4812
            _Version        =   65536
            _ExtentX        =   8488
            _ExtentY        =   614
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   6
               Left            =   1680
               TabIndex        =   213
               Text            =   "12"
               Top             =   36
               Width           =   588
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   5
               Left            =   48
               TabIndex        =   212
               Text            =   "TimesNewRoman"
               Top             =   36
               Width           =   1596
            End
            Begin Threed.SSRibbon SSRibbon 
               Height          =   252
               Index           =   4
               Left            =   3600
               TabIndex        =   219
               Top             =   48
               Width           =   300
               _Version        =   65536
               _ExtentX        =   529
               _ExtentY        =   445
               _StockProps     =   65
               BackColor       =   12632256
               GroupNumber     =   4
               GroupAllowAllUp =   -1  'True
               Autosize        =   1
               RoundedCorners  =   0   'False
               BevelWidth      =   1
               Outline         =   0   'False
            End
            Begin Threed.SSRibbon SSRibbon 
               Height          =   252
               Index           =   5
               Left            =   3936
               TabIndex        =   218
               Top             =   48
               Width           =   300
               _Version        =   65536
               _ExtentX        =   529
               _ExtentY        =   445
               _StockProps     =   65
               BackColor       =   12632256
               GroupNumber     =   4
               GroupAllowAllUp =   -1  'True
               Autosize        =   1
               RoundedCorners  =   0   'False
               BevelWidth      =   1
               Outline         =   0   'False
            End
            Begin Threed.SSRibbon SSRibbon 
               Height          =   252
               Index           =   6
               Left            =   4272
               TabIndex        =   217
               Top             =   48
               Width           =   300
               _Version        =   65536
               _ExtentX        =   529
               _ExtentY        =   445
               _StockProps     =   65
               BackColor       =   12632256
               GroupNumber     =   4
               GroupAllowAllUp =   -1  'True
               Autosize        =   1
               RoundedCorners  =   0   'False
               BevelWidth      =   1
               Outline         =   0   'False
            End
            Begin Threed.SSRibbon SSRibbon 
               Height          =   252
               Index           =   3
               Left            =   3120
               TabIndex        =   216
               Top             =   48
               Width           =   300
               _Version        =   65536
               _ExtentX        =   529
               _ExtentY        =   445
               _StockProps     =   65
               BackColor       =   12632256
               GroupNumber     =   3
               GroupAllowAllUp =   -1  'True
               Autosize        =   1
               RoundedCorners  =   0   'False
               BevelWidth      =   1
               Outline         =   0   'False
            End
            Begin Threed.SSRibbon SSRibbon 
               Height          =   252
               Index           =   2
               Left            =   2784
               TabIndex        =   215
               Top             =   48
               Width           =   300
               _Version        =   65536
               _ExtentX        =   529
               _ExtentY        =   445
               _StockProps     =   65
               BackColor       =   12632256
               GroupNumber     =   2
               GroupAllowAllUp =   -1  'True
               Autosize        =   1
               RoundedCorners  =   0   'False
               BevelWidth      =   1
               Outline         =   0   'False
            End
            Begin Threed.SSRibbon SSRibbon 
               Height          =   252
               Index           =   1
               Left            =   2448
               TabIndex        =   214
               Top             =   48
               Width           =   300
               _Version        =   65536
               _ExtentX        =   529
               _ExtentY        =   445
               _StockProps     =   65
               BackColor       =   12632256
               GroupAllowAllUp =   -1  'True
               Autosize        =   1
               RoundedCorners  =   0   'False
               BevelWidth      =   1
               Outline         =   0   'False
            End
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   26
            Left            =   6144
            TabIndex        =   131
            Top             =   5712
            Width           =   732
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "OK"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   252
            Index           =   27
            Left            =   7068
            TabIndex        =   130
            Top             =   5712
            Width           =   780
            _Version        =   65536
            _ExtentX        =   1376
            _ExtentY        =   445
            _StockProps     =   78
            Caption         =   "Cancelar"
         End
         Begin MSDBCtls.DBCombo DBCombo 
            Height          =   288
            Index           =   25
            Left            =   1392
            TabIndex        =   122
            Top             =   960
            Width           =   3504
            _ExtentX        =   6181
            _ExtentY        =   508
            _Version        =   327681
            BackColor       =   16777215
            ForeColor       =   0
            Text            =   "Steve Armour"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Universal:"
            ForeColor       =   &H00000000&
            Height          =   192
            Index           =   58
            Left            =   4992
            TabIndex        =   124
            Top             =   240
            Width           =   792
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Título:"
            Height          =   192
            Index           =   57
            Left            =   96
            TabIndex        =   123
            Top             =   240
            Width           =   432
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Pesquisador:"
            Height          =   192
            Index           =   60
            Left            =   1392
            TabIndex        =   121
            Top             =   768
            Width           =   912
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Data:"
            Height          =   192
            Index           =   59
            Left            =   96
            TabIndex        =   120
            Top             =   768
            Width           =   396
            WordWrap        =   -1  'True
         End
      End
      Begin Threed.SSFrame SSFrame 
         Height          =   6060
         Index           =   19
         Left            =   -74952
         TabIndex        =   132
         Top             =   540
         Width           =   8940
         _Version        =   65536
         _ExtentX        =   15769
         _ExtentY        =   10689
         _StockProps     =   14
         Caption         =   "Editando Mapa"
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         ShadowStyle     =   1
         Begin VB.TextBox Text 
            Height          =   300
            Index           =   9
            Left            =   96
            TabIndex        =   133
            Text            =   "Text3"
            Top             =   432
            Width           =   4764
         End
         Begin Threed.SSFrame SSFrame 
            Height          =   4764
            Index           =   20
            Left            =   4992
            TabIndex        =   134
            Top             =   720
            Width           =   3852
            _Version        =   65536
            _ExtentX        =   6795
            _ExtentY        =   8403
            _StockProps     =   14
            Caption         =   "Acontecimento"
            ForeColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin MSDBCtls.DBCombo DBCombo 
               Height          =   288
               Index           =   24
               Left            =   96
               TabIndex        =   135
               Top             =   240
               Width           =   3660
               _ExtentX        =   6456
               _ExtentY        =   508
               _Version        =   327681
               BackColor       =   16777215
               ForeColor       =   0
               Text            =   "DBCombo(24)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   348
            Index           =   22
            Left            =   5136
            TabIndex        =   137
            Top             =   5616
            Width           =   1116
            _Version        =   65536
            _ExtentX        =   1969
            _ExtentY        =   614
            _StockProps     =   78
            Caption         =   "Cancelar"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   348
            Index           =   21
            Left            =   3600
            TabIndex        =   136
            Top             =   5616
            Width           =   1116
            _Version        =   65536
            _ExtentX        =   1969
            _ExtentY        =   614
            _StockProps     =   78
            Caption         =   "OK"
         End
         Begin VB.Label Label 
            BorderStyle     =   1  'Fixed Single
            Height          =   264
            Index           =   56
            Left            =   4992
            TabIndex        =   140
            Top             =   432
            Width           =   3828
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Título:"
            Height          =   192
            Index           =   54
            Left            =   96
            TabIndex        =   139
            Top             =   240
            Width           =   432
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Título da Descrição:"
            ForeColor       =   &H00000000&
            Height          =   192
            Index           =   55
            Left            =   4992
            TabIndex        =   138
            Top             =   240
            Width           =   1452
            WordWrap        =   -1  'True
         End
      End
      Begin Threed.SSFrame SSFrame 
         Height          =   5964
         Index           =   23
         Left            =   -74952
         TabIndex        =   107
         Top             =   588
         Width           =   8940
         _Version        =   65536
         _ExtentX        =   15769
         _ExtentY        =   10520
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSFrame SSFrame 
            Height          =   1356
            Index           =   24
            Left            =   96
            TabIndex        =   108
            Top             =   156
            Width           =   8700
            _Version        =   65536
            _ExtentX        =   15346
            _ExtentY        =   2392
            _StockProps     =   14
            Caption         =   "Dados de controle"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
            Begin VB.TextBox Text 
               Height          =   300
               Index           =   11
               Left            =   144
               TabIndex        =   110
               Top             =   960
               Width           =   4860
            End
            Begin VB.TextBox Text 
               Height          =   300
               Index           =   12
               Left            =   5280
               TabIndex        =   109
               Top             =   960
               Width           =   2172
            End
            Begin MSDBCtls.DBCombo DBCombo 
               Height          =   288
               Index           =   27
               Left            =   156
               TabIndex        =   115
               Top             =   432
               Width           =   8460
               _ExtentX        =   14923
               _ExtentY        =   508
               _Version        =   327681
               BackColor       =   16777215
               ForeColor       =   0
               Text            =   "DBCombo(27)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Autor:"
               Height          =   192
               Index           =   63
               Left            =   5280
               TabIndex        =   114
               Top             =   768
               Width           =   408
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Universal:"
               Height          =   192
               Index           =   62
               Left            =   144
               TabIndex        =   113
               Top             =   768
               Width           =   720
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Título:"
               Height          =   192
               Index           =   61
               Left            =   144
               TabIndex        =   112
               Top             =   240
               Width           =   432
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Data:"
               Height          =   192
               Index           =   64
               Left            =   7680
               TabIndex        =   111
               Top             =   768
               Width           =   396
               WordWrap        =   -1  'True
            End
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   348
            Index           =   30
            Left            =   5616
            TabIndex        =   118
            Top             =   5520
            Width           =   1116
            _Version        =   65536
            _ExtentX        =   1969
            _ExtentY        =   614
            _StockProps     =   78
            Caption         =   "Imprimir"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   348
            Index           =   29
            Left            =   3984
            TabIndex        =   117
            Top             =   5520
            Width           =   1116
            _Version        =   65536
            _ExtentX        =   1969
            _ExtentY        =   614
            _StockProps     =   78
            Caption         =   "Editar"
         End
         Begin Threed.SSCommand SSCommand 
            Height          =   348
            Index           =   28
            Left            =   2400
            TabIndex        =   116
            Top             =   5520
            Width           =   1116
            _Version        =   65536
            _ExtentX        =   1969
            _ExtentY        =   614
            _StockProps     =   78
            Caption         =   "Nova"
         End
      End
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   8796
      Top             =   6804
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":272D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":278B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":27E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":29C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2B9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2BFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2C59
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2E33
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":300D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":31E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":33C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":359B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3775
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":37D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3831
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":388F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":38ED
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   8292
      Top             =   6720
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   16
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   48
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3A3F
            Key             =   "A3b"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3C11
            Key             =   "A3a"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3DE3
            Key             =   "A1d"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3FB5
            Key             =   "A1c"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4187
            Key             =   "A2b"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4359
            Key             =   "A2a"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":452B
            Key             =   "A1b"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":46FD
            Key             =   "A1a"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":48CF
            Key             =   "B1a"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4AA1
            Key             =   "B1b"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4C73
            Key             =   "B1c"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4E45
            Key             =   "B1d"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5017
            Key             =   "B2a"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":51E9
            Key             =   "B2b"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":53BB
            Key             =   "B2c"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":558D
            Key             =   "B2d"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":575F
            Key             =   "B3a"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5931
            Key             =   "B3b"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5B03
            Key             =   "B3c"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5CD5
            Key             =   "B3d"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5EA7
            Key             =   "C1a"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6079
            Key             =   "C1b"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":624B
            Key             =   "C1c"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":641D
            Key             =   "C1d"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":65EF
            Key             =   "C2a"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":67C1
            Key             =   "C2b"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6993
            Key             =   "C2c"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6B65
            Key             =   "C2d"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6D37
            Key             =   "C3a"
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6F09
            Key             =   "C3b"
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":70DB
            Key             =   "C3c"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":72AD
            Key             =   "C3d"
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":747F
            Key             =   "D1a"
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7651
            Key             =   "D1b"
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7823
            Key             =   "D1c"
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":79F5
            Key             =   "D1d"
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7BC7
            Key             =   "D2a"
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7D99
            Key             =   "D2b"
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7F6B
            Key             =   "D2c"
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":813D
            Key             =   "D2d"
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":830F
            Key             =   "D3a"
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":84E1
            Key             =   "D3b"
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":86B3
            Key             =   "D3c"
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":8885
            Key             =   "D3d"
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":8A57
            Key             =   "Homem"
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":8BA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":8CBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":8E95
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "PAC1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim it() As Integer
Public contador As Integer
Public ID_Nome_DBcombo11 As Integer 'Esta variavel captura o ID_Ego Masculino.
Public ID_Nome_DBcombo12 As Integer 'Esta variavel captura o ID_Ego Feminino.
Public ID_Nome_DBcombo13 As Integer 'Esta variavel captura o ID_Ego para ser usado no momento de inserir um novo filho na lista.
Public DBcombo11_clicou As Integer
Public DBcombo12_clicou As Integer
Public DBcombo13_clicou As Integer 'Esta variavel captura o ID_Ego para ser usado no momento de inserir um novo filho na lista.
Public TreeView2_Clicou As Integer 'Esta variavel captura o ID_Ego para ser usado no momento de inserir um novo filho na lista.
Public ÍndiceAtual As Integer
Public Civil As Integer 'Recebe o índice que representa o estado civil.
Public Foi_Clicado As Integer  'Indica que um combo ou outro controle foi clicado.
Public Proximo_Casal As Integer
Public ID_Ego_TB As Integer


Private Sub DBGrid2_Click()
'Coluna 1=Nome do parente; coluna 2=Termo tecnico; coluna 3=termo indigena
End Sub

Private Sub DBGrid3_Click()
' NOTA PARA MACS: Este grid tem que ter 3 colunas:
' uma para a pergunta ampliada, outra para mostrar as confirmações,
' e outra para fazer a ligação com um acontecimento novo.
' Assim, automaticamente quando ele clique no lugar na coluna e row,
' a janela de acontecimentos abre (UM Problema: o que acontece se ele já
' tem um acontecimento não salva dentro da sessão de acontecimentos?),
' ele bate o acontecimento, salva a primeira vez, e
' e o índice do acontecimento está ligado e vista na coluna 3.
' Deve voltar ao sessão de perguntas depois dele salvar o acontecimento.
End Sub



Private Sub AniButton_Click()
    Civil = Switch(AniButton.Value = 1, 6, AniButton.Value = 2, 3, AniButton.Value = 3, 5)
End Sub

Private Sub Combo_Click(Index As Integer)
    'Caso o PAC ainda está carregando, a variável EstouCarregando terá o valor 1
    'e avisando que a rotina daqui não deve ser executada.
    If EstouCarregando = 1 Or BancoCP_EGO = 0 Then Exit Sub
    Dim ConstIdioma As String 'Identifica qual a língua usada na interface.
    
    Select Case Index 'Esta variável recebe o índice de todos os combos, por isso não pode ser usada para identificar os combos abaixo.
        
        Case 2  'Seleciona o tipo de termo de parentesco. Referência ou Tratamento.
            Combo(4).ListIndex = Combo(2).ListIndex 'Iguala o combo(4) com o combo(2)
            Call Lista_Termos
            If SSOption(15).Value = True Then Call SSOption_Click(15, True)
        
        Case 4  'Seleciona o tipo de termo de parentesco. Referência ou Tratamento.
            'Se o tab Parentesco/Planejamento estiver ativo, o combo(2) de Parentesco/Geral será igualado ao combo(4).
            'É por este combo(2)(Parentesco/Geral ) que a rotina de listar termos é ativada.
            If SSTab(3).Tab = 1 Then Combo(2).ListIndex = Combo(4).ListIndex
        
        Case 7
            If Foi_Clicado = 1 Then 'Este teste não deixa a rotina entrar num loop por ser selecionado
                Foi_Clicado = 0     'um item no combo e tentar rodar a rotina de troca de idioma indefinidamente.
            Else
                Unload Novo_Termo
                If Combo(7).ListIndex = 0 Then 'Se o combo "Idioma" for igual a 0 (Português), então...
                    ConstIdioma = 0 'Põe o valor 0=Português nesta variável.
                Else 'Caso o combo "Idioma" for diferente de 0...
                    ConstIdioma = 2300 'Põe o valor 2300=Inglês nesta variável.
                End If
                
                Combo(7).Clear 'Limpa o combo "Idioma", pois ele vai ser peenchido com o novo idioma selecionado.
                Combo(2).Clear 'Limpa o combo "Classificação", pois ele vai ser peenchido com o novo idioma selecionado.
                Combo(4).Clear
                'Novo_Termo.Combo(0).Clear 'Limpa o combo "Classificação" na janela Novo_Termo, pois ele vai ser peenchido com o novo idioma selecionado.
                'Grava no Pac.ini o valor do novo idioma selecionado.
                a = WritePrivateProfileString("IDIOMA", "Idioma", ConstIdioma, "pac.ini")
                Lingua = ConstIdioma 'Muda esta variável pública para o valor do novo idioma selecionado.
                Idioma (ConstIdioma) 'Muda toda a interface com o novo idioma selecionado.
                DBGrid(1).Columns(0).DataField = IIf(Lingua = 0, "Termo_Tec", "Termo_Tec_IN")
                'DBGrid(1).ClearFields
                DBGrid(1).Refresh
                Select Case Lingua
                    Case 0      'Português
                        Foi_Clicado = 1 'Indica que um item foi selecionado para evitar o loop.
                        PAC1.Combo(7).ListIndex = 0 'Português
                    Case 2300    'Inglês
                        Foi_Clicado = 1 'Indica que um item foi selecionado para evitar o loop.
                        PAC1.Combo(7).ListIndex = 1 'Inglês
                End Select
                PAC1.Combo(2).ListIndex = 0 'seleciona o primeiro item na lista.
                PAC1.Combo(4).ListIndex = 0
                DBCombo(0).Text = "" 'Deixa o dbcombo limpo para o usuário escolher um novo termo.
                'Novo_Termo.Combo(0).ListIndex = 0 'seleciona o primeiro item na lista.
            End If

        Case 9  'Nome_Indígena
            'Se o nome indígena não é vazio então...
            If PAC1.Combo(9).ListIndex < 0 Then PAC1.Combo(9).ListIndex = 0
            If PAC1.Combo(9).Text <> "" Then Procura_Ego (PAC1.Combo(9).ItemData(PAC1.Combo(9).ListIndex)) 'Procura pelo valor do ItemData que será igual ao ID_Ego
            Total.Caption = DBcp_Ego.Recordset("ID_Ego") & " - " & DBcp_Ego.Recordset.RecordCount 'Atualiza o Label Total.
        
        Case 10 'Nome_Nacional
            'Se o nome nacional não é vazio então...
            If PAC1.Combo(10).Text <> "" Then Procura_Ego (PAC1.Combo(10).ItemData(PAC1.Combo(10).ListIndex)) 'Procura pelo valor do ItemData que será igual ao ID_Ego
            Total.Caption = DBcp_Ego.Recordset("ID_Ego") & " - " & DBcp_Ego.Recordset.RecordCount 'Atualiza o Label Total.
    
    End Select
    
End Sub


Private Sub DBCombo_Click(Index As Integer, Area As Integer)
    Dim MeuCritério As String
    Dim Figura As Integer
    
    'Lista de termos técnicos no Parentesco/Planejamento.
    If Index = 0 And Area = 2 And DBCombo(0).BoundText <> "" Then
        Dim TB As String
        Dim TB_Anterior As String
        Dim Passos As Integer
        Dim Sexo As String
        Dim Trilha As String
        Dim Nova_Trilha As String
        Dim Irmão_Critério As String
        Dim Esposa_Critério As String
        Dim Proximo_Casal As Integer
        Dim Esposa_Outra_Camada As Integer
        Dim ID_Ego_TB As Integer
        Dim Zero As Integer
        Dim denovo As Integer
        Dim Qual_Conj As String
        Dim Qt As Integer
        Dim Sequ As Integer
        Dim MaisUm As Integer
        List(5).Clear
        List(6).Clear
        Mapa_Parente.Cls
        Sexo = IIf(SSOption(16), "0", "1") 'Seleciona o sexo, 0>M  1>F
        MeuCritério = "Select * from Ego where sexo=" & Sexo
        Trilha = DBCombo(0).BoundText
        DB_Temp.DatabaseName = "dbcp.mdb"
        DB_Temp.RecordSource = MeuCritério 'Lembre-se que todos os egos neste db_temp tem o mesmo sexo.
        DB_Temp.Refresh
        DB_Temp.Recordset.MoveLast
        EstouCarregando = 1
        Proximo_Casal = 0
        DB_Temp.Recordset.MoveFirst
        Camada = 1
        ReDim Tree_Filhos(0 To DBcp_Ego.Recordset.RecordCount - 1, 0 To DBcp_Ego.Recordset.RecordCount - 1, 100, 20)
        ReDim Arvore(0 To DBcp_Ego.Recordset.RecordCount - 1, 0 To 5000, 0 To 1)
        Esposa_Outra_Camada = 0
        Do While DB_Temp.Recordset.EOF = False
            Ego_Inicial = DB_Temp.Recordset("id_ego") 'Mantendo a ligação com o ego inicial
            If Ego_Inicial = 2 Then Stop
    Do
        If Esposa_Outra_Camada = 0 Then
            'Vai entrar aqui quando procura os parentes da(s) esposa(s).
            If Len(Trilha) > 1 And Left(Trilha, 1) = "8" Then
                'Se o ego selecionado for o Feminino, então deve sair deste procedimento, _
                 pois a condição esposa de uma mulher não vale.
                If SSOption(17).Value = True Then Exit Sub
                'Este critério procura pelas uniões feitas pelo corrente ego. _
                 Observe que ele pode ter ou teve várias famílias de procriação. _
                 As uniões desfeitas por separação não entram aqui.
                MeuCritério = "ID_Conj1= " & Ego_Inicial _
                               & " and Civil <> 5"
                Call Acha_Esposas(MeuCritério, 0, Trilha)
                'O proximo bloco testa se o corrente ego tem esposa(s). Se tem, então _
                 o ID dela(s) passa para a variável publica Procurar_De_Quem para que todos _
                 os TB na sequencia sejam os da(s) esposa(s).
                If Arvore(Ego_Inicial, 0, 0) <> "" Then
                    ReDim Procurar_De_Quem(0 To CInt(Arvore(Ego_Inicial, 0, 0)))
                    For Qt = 1 To CInt(Arvore(Ego_Inicial, 0, 0))
                        Procurar_De_Quem(Qt) = CInt(Arvore(Ego_Inicial, Qt, 0))
                    Next
                    Procurar_De_Quem(0) = Arvore(Ego_Inicial, 0, 0)
                End If
            Else 'Aqui é pq está procurando os parentes do próprio ego.
                ReDim Procurar_De_Quem(0 To 1)
                Procurar_De_Quem(1) = DB_Temp.Recordset("ID_Ego")
                Procurar_De_Quem(0) = 1
            End If
        ElseIf Esposa_Outra_Camada = 1 Then
                'O proximo bloco testa se o corrente ego tem esposa(s). Se tem, então _
                 o ID dela(s) passa para a variável publica Procurar_De_Quem para que todos _
                 os TB na sequencia sejam os da(s) esposa(s).
                If Arvore(Ego_Inicial, 0, 0) <> "" Then
                    ReDim Procurar_De_Quem(0 To CInt(Arvore(Ego_Inicial, 0, 0)))
                    For Qt = 1 To CInt(Arvore(Ego_Inicial, 0, 0))
                        Procurar_De_Quem(Qt) = CInt(Arvore(Ego_Inicial, Qt, 0))
                        Arvore(Ego_Inicial, Qt, 0) = ""
                    Next
                    Procurar_De_Quem(0) = Arvore(Ego_Inicial, 0, 0)
                End If
        End If
        'Se procuramos parentes das esposas do ego inicial, então a variável MaisUm vai ciclar _
         mais que uma vez.
        For MaisUm = 1 To Procurar_De_Quem(0)
            'Se estamos procurando parentes da esposa, mas nenhuma esposa do corrente ego _
             foi encontrada, então pulamos para o próximo ego.
            If Procurar_De_Quem(1) <> DB_Temp.Recordset("ID_Ego") _
            And Arvore(Ego_Inicial, 0, 0) = "" Then Exit For
            denovo = 0
            For Passos = 1 To Len(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha)) 'Passeia pela string capturando cada TB.
                If Passos = 1 Then 'Se for o primeiro Termo Básico.
                    ID_Ego_TB = -1 'Esta variável é zerada aqui.
                    Proximo_Casal = 0 'Nenhum casal é selecionado.
                End If
                'Pega cada Termo Básico individualmente de dentro da string.
                TB = Mid(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha), Passos, 1)
                Select Case TB
                    
                    Case "1" 'Pai
                        denovo = 0 'zera esta variável caso seja necessário entrar em filhos depois daqui.
                        Proximo_Casal = TB_Pais(MaisUm, Proximo_Casal, "ID_Conj1", Passos, Esposa_Outra_Camada, Trilha, Nova_Trilha)
                        If Proximo_Casal = 0 Then
                            ID_Ego_TB = -1
                            Exit For
                        Else
                            ID_Ego_TB = DBcp_Ego.Recordset("ID_Ego") 'Recebe o ID do homem que responde pelo TB.
                        End If
                            
                    Case "2" 'Mãe
                        denovo = 0 'zera esta variável caso seja necessário entrar em filhos depois daqui.
                        Proximo_Casal = TB_Pais(MaisUm, Proximo_Casal, "ID_Conj2", Passos, Esposa_Outra_Camada, Trilha, Nova_Trilha)
                        If Proximo_Casal = 0 Then
                            ID_Ego_TB = -1
                            Exit For
                        Else
                            ID_Ego_TB = DBcp_Ego.Recordset("ID_Ego") 'Recebe o ID do homem que responde pelo TB.
                        End If
                        
                    Case "3" 'Irmão
                        denovo = 0 'zera esta variável caso seja necessário entrar em filhos depois daqui.
                        Proximo_Casal = TB_Irmãos(MaisUm, Proximo_Casal, ID_Ego_TB, " sexo=0", Passos, Esposa_Outra_Camada, Trilha, Nova_Trilha)
                        If Proximo_Casal = 0 Then
                            ID_Ego_TB = -1
                            Exit For
                        Else
                            ID_Ego_TB = DBcp_Ego.Recordset("ID_Ego") 'Recebe o ID do homem que responde pelo TB.
                        End If
                    
                    Case "4" 'Irmã
                    'If Ego_Inicial = 13 Then Stop
                        denovo = 0 'zera esta variável caso seja necessário entrar em filhos depois daqui.
                        Proximo_Casal = TB_Irmãos(MaisUm, Proximo_Casal, ID_Ego_TB, " sexo=1", Passos, Esposa_Outra_Camada, Trilha, Nova_Trilha)
                        If Proximo_Casal = 0 Then
                            ID_Ego_TB = -1
                            Exit For
                        Else
                            ID_Ego_TB = DBcp_Ego.Recordset("ID_Ego") 'Recebe o ID do homem que responde pelo TB.
                        End If

                    Case "5" 'Filho
            'If Ego_Inicial = 13 Then Stop
                        Sexo = "Sexo=0"
                        'Se a Trilha é maior que 1 então pode ser estejamos procurando _
                         filha do filho ou filho da filha. Isto exije uma mudança da variável _
                         Sexo, pois estamos tratando de sexo invertido na busca dos cônjuges.
                        If Passos > 1 Then
                            If Mid(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha), Passos - 1, 1) = "6" _
                            And denovo >= 1 Then Sexo = "Sexo=1"
                        End If
                        denovo = TB_Prole(TB, denovo, Sexo, MaisUm, Proximo_Casal, ID_Ego_TB, Passos, Esposa_Outra_Camada, Trilha, Nova_Trilha)
                    
                    Case "6" 'Filha
             'If Ego_Inicial = 12 Then Stop
                        Sexo = "Sexo=1"
                        'Se a Trilha é maior que 1 então pode ser estejamos procurando _
                         filha do filho ou filho da filha. Isto exije uma mudança da variável _
                         Sexo, pois estamos tratando de sexo invertido na busca dos cônjuges.
                        If Passos > 1 Then
                            If Mid(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha), Passos - 1, 1) = "5" _
                            And denovo >= 1 Then Sexo = "Sexo=0"
                        End If
                        denovo = TB_Prole(TB, denovo, Sexo, MaisUm, Proximo_Casal, ID_Ego_TB, Passos, Esposa_Outra_Camada, Trilha, Nova_Trilha)
                    
                    Case "7" 'Esposo
                        If Ego_Inicial = 2 Then Stop
                        'Se o ego selecionado for o masculino, então deve sair deste procedimento, _
                         pois a condição esposo de um homem não vale.
                        If Len(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha)) = 1 And SSOption(16).Value = True Then Exit Sub
                        If TB_Conjuges("ID_Conj2=", Esposa_Outra_Camada, Trilha, Nova_Trilha, Passos) = 1 Then
                            pulaUma = 1
                            Exit For
                        End If

                    
                    Case "8" 'Esposa
                    'Só entra aqui se está procurando a esposa do ego inicial ou se
                    If Procurar_De_Quem(1) = DB_Temp.Recordset("ID_Ego") Or Passos > 2 Then
                        'Caso a busca seja pela(s) esposa(s) do corrente ego...
                        If Len(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha)) = 1 Then
                            'Se o ego selecionado for o Feminino, então deve sair deste procedimento, _
                             pois a condição esposa de uma mulher não vale.
                            If SSOption(17).Value = True Then Exit Sub
                            'Este critério procura pelas uniões feitas pelo corrente ego. _
                             Observe que ele pode ter ou teve várias famílias de procriação. _
                             As uniões desfeitas por separação não entram aqui.
                            MeuCritério = "ID_Conj1= " & DB_Temp.Recordset("ID_Ego") _
                                           & " and Civil <> 5"
                            Call Acha_Esposas(MeuCritério, Passos, IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha))
                        'Vai entrar aqui quando procura a(s) esposa(s) que não _
                         sejam do próprio ego.
                        Else
                            'Esta variável pega o termo básico anterior ao tb esposa. Estamos _
                             procurando esposa(s) do pai, do irmão e do filho.
                            TB_Anterior = Mid(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha), Passos - 1, 1)
                            Select Case TB_Anterior
                                Case "1" 'Pai
                                    'Lembre que o ID do pais está guardado nesta variável: Arvore(ego_inicial, 1, 0)
                                    'Este critério procura pelas uniões feitas pelo pai do ego. _
                                     Estamos procurando por aquelas esposas do pai que não seja _
                                     a própria mãe do ego. As uniões desfeitas pelo pai não importam aqui.
                                    MeuCritério = "ID_Conj1= " & Arvore(Ego_Inicial, 1, 0) & _
                                                   " and ID_Casal<> " & DB_Temp.Recordset("ID_Pais") & _
                                                   " and Civil<> 5"
                                    Call Acha_Esposas(MeuCritério, Passos, IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha))
                                Case "3" 'Irmão
                                    'Já sabemos quantos irmãos o ego em db_temp tem, pois _
                                     o processo já passou pela seção de busca de irmãos. _
                                     Carreguei os irmãos nas variáveis Cada_Irmão() e esvaziei _
                                     as variáveis Arvore() por segurança, pois esta variável _
                                     será manipulada pela função Acha_Esposas()
                                    If Arvore(Ego_Inicial, 0, 0) <> "" Then
                                        qt_irmão = CInt(Arvore(Ego_Inicial, 0, 0))
                                        ReDim Cada_Irmão(qt_irmão) As Integer
                                        For Sequ = 1 To qt_irmão
                                            Cada_Irmão(Sequ) = CInt(Arvore(Ego_Inicial, Sequ, 0))
                                            Arvore(Ego_Inicial, Sequ, 0) = ""
                                            Arvore(Ego_Inicial, Sequ, 1) = ""
                                        Next Sequ
                                        Arvore(Ego_Inicial, 0, 0) = ""
                                        For Qt = 1 To qt_irmão
                                            'Este critério procura as esposas dos irmãos atuais ou que já morreram. _
                                             As esposas de casamentos desfeitos por separação não contam.
                                            Irmão_Critério = "ID_Conj1= " & Cada_Irmão(Qt) & " and Civil<> 5"
                                            'Esta função será chamada com cada irmão selecionado.
                                            Call Acha_Esposas(Irmão_Critério, Passos, IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha))
                                        Next Qt
                                    End If
                                Case "5" 'Filho
                                    'Já sabemos quantos filhos o ego em db_temp tem, pois _
                                     o processo já passou pela seção de busca de filhos. _
                                     Carreguei os filhos nas variáveis Cada_Irmão() e esvaziei _
                                     as variáveis Arvore() por segurança, pois esta variável _
                                     será manipulada pela função Acha_Esposas()
                                    If Arvore(Ego_Inicial, 0, 0) <> "" Then
                                        qt_irmão = CInt(Arvore(Ego_Inicial, 0, 0))
                                        ReDim Cada_Irmão(qt_irmão) As Integer
                                        For Sequ = 1 To qt_irmão
                                            Cada_Irmão(Sequ) = CInt(Arvore(Ego_Inicial, Sequ, 0))
                                            Arvore(Ego_Inicial, Sequ, 0) = ""
                                            Arvore(Ego_Inicial, Sequ, 1) = ""
                                        Next Sequ
                                        Arvore(Ego_Inicial, 0, 0) = ""
                                        For Qt = 1 To qt_irmão
                                            'Este critério procura as esposas dos irmãos atuais ou que já morreram. _
                                             As esposas de casamentos desfeitos por separação não contam.
                                            Irmão_Critério = "ID_Conj1= " & Cada_Irmão(Qt) & " and Civil<> 5"
                                            'Esta função será chamada com cada irmão selecionado.
                                            Call Acha_Esposas(Irmão_Critério, Passos, IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha))
                                        Next Qt
                                    End If
                            End Select
                            
                        End If
                        'Caso nenhuma esposa seja encontrada, então o For-Next é interrompido _
                         para que outro ego inicial seja escolhido para outra busca.
                        If Arvore(Ego_Inicial, 0, 0) = "" Then Exit For
                        'Caso estamos procurando parentes de esposas do irmão, pai ou filho. Isto _
                         siginifica que o termo tecnico não é composto apenas de Esposa, nem o primeiro TB e _
                         nem o último TB é Esposa. Left(Trilha, 1) <> "8")
                         'If DB_Temp.Recordset("ID_Ego") = 0 Then Stop
                         qqqq = Arvore(41, 0, 0)
                         Sequ = 0
                         posi = IIf(InStr(Trilha, "8") = 2, 2, 3)
                         Sequ = InStr(posi, Trilha, 8)
                        If Len(Trilha) > 2 And Sequ <> 0 And Right(Trilha, 1) <> "8" Then
                            Nova_Trilha = Mid(Trilha, Passos + 1, Len(Trilha) - (Passos))
                            Esposa_Outra_Camada = 1
                            pulaUma = 1
                            Exit For
                        
                        End If
                    End If
                End Select
            Next Passos
            If pulaUma = 0 Then
                If MaisUm = Procurar_De_Quem(0) And Esposa_Outra_Camada = 1 Then
                    Esposa_Outra_Camada = 0: Exit For
                End If
            Else
                pulaUma = 0
            End If
        Next MaisUm
        If Esposa_Outra_Camada = 0 Then Exit Do
    Loop
        DB_Temp.Recordset.MoveNext
        Loop
        If List(5).ListCount <> 0 Then List(5).ListIndex = 0 'Seleciona o primeiro ego da lista.
        EstouCarregando = 0
    End If
    
    If Index = 11 And Area = 2 And DBCombo(11).Text <> "" Then
        MeuCritério = "ID_Ego = " & Val(DBCombo(11).BoundText)
        DBcp_Ego_Masculino.Recordset.FindFirst MeuCritério
        ID_Nome_DBcombo11 = Val(DBCombo(11).BoundText)
        DBcombo11_clicou = 1
        Figura = IIf(IsNull(DBcp_Ego_Masculino.Recordset("data_falec")) = True, 10, 9)
        Membro_Família(1).Picture = ImageList2.ListImages(Figura).Picture
        
    End If
    
    If Index = 12 And Area = 2 And DBCombo(12).Text <> "" Then
        'If Nome_Nac(1).Value = False Then
            'DBcp_Ego_Feminino.Recordset.FindFirst "Nome_Ind='" & NovaString(DBCombo(12).Text) & "'"
        'Else
            'DBcp_Ego_Feminino.Recordset.FindFirst "Nome_Nac='" & NovaString(DBCombo(12).Text) & "'"
        'End If
        'ID_Nome_DBcombo12 = DBcp_Ego_Feminino.Recordset("ID_Ego")
        
        MeuCritério = "ID_Ego = " & Val(DBCombo(12).BoundText)
        DBcp_Ego_Feminino.Recordset.FindFirst MeuCritério
        ID_Nome_DBcombo12 = Val(DBCombo(12).BoundText)
        DBcombo12_clicou = 1
        Figura = IIf(IsNull(DBcp_Ego_Feminino.Recordset("data_falec")) = True, 12, 11)
        Membro_Família(2).Picture = ImageList2.ListImages(Figura).Picture
    End If
    
    If Index = 13 And Area = 2 Then
        If DBCombo(13).Text <> "" Then
             'Se a opção do nome nacional não estiver selecionada, então...
             'If Nome_Nac(2).Value = False Then
                 'É feita uma busca pelo nome indígena.
                 'O apóstrofe (') é trocado por ('') pela função NovaString(). Sem isto o findfirst não funciona.
                 'DBcp_Ego_Nomes.Recordset.FindFirst "Nome_Ind='" & NovaString(DBCombo(13).Text) & "'"
             'Caso a opção do nome nacional estiver selecionada...
             'Else
                 'É feita uma busca pelo nome nacional.
                 'O apóstrofe (') é trocado por ('') pela função NovaString(). Sem isto o findfirst não funciona.
                 'DBcp_Ego_Nomes.Recordset.FindFirst "Nome_Nac='" & NovaString(DBCombo(13).Text) & "'"
             'End If
             'O índice do ego selecionado em combo(13) é amazenado nesta variável pública.
             'ID_Nome_DBcombo13 = DBcp_Ego_Nomes.Recordset("ID_Ego")
             MeuCritério = "ID_Ego = " & Val(DBCombo(13).BoundText)
             DBcp_Ego_Nomes.Recordset.FindFirst MeuCritério
             ID_Nome_DBcombo13 = Val(DBCombo(13).BoundText)
             'Esta variável indica para outro controles que o dbcombo(13) foi clicado.
             DBcombo13_clicou = 1
             'Mostra o ícone apropriado para o ego selecionado.
             Call Qual_Icon(DBcp_Ego_Nomes, Membro_Família(3))
        Else
             'Não mostra nada, poís nenhum ego está selecionado.
             Membro_Família(3).Picture = ImageList2.ListImages(16).Picture
        End If
    End If
End Sub


Private Sub DBcp_Casais_Reposition()
    Dim Critério As String
    Dim QualNome As String
    Dim itmX As ListItem
    Dim Ícone As Integer
    Dim Figura As Integer
    If EstouCarregando = 0 Then 'Se o PAC não está sendo carregado agora, então...
    
        EstouCarregando = 1
        AniButton.Value = Switch(PAC1.DBcp_Casais.Recordset("civil") = 6, 1, PAC1.DBcp_Casais.Recordset("civil") = 3, 2, PAC1.DBcp_Casais.Recordset("civil") = 5, 3)
        Line1.BorderStyle = IIf(AniButton.Value = 1, 3, 1)
        'Coloca os filhos do casal no listview1
        PAC1.ListView1.ListItems.Clear
        Combo(15).Text = DBcp_Casais.Recordset("Ajudante") 'Ajusta o nome do Ajudante do corrente registro.
        Critério = "ID_Pais=" & DBcp_Casais.Recordset("ID_Casal")
        DBcp_Ego.Recordset.FindFirst Critério
        If DBcp_Ego.Recordset.NoMatch = False Then
            Do Until DBcp_Ego.Recordset.NoMatch = True
                QualNome = IIf(DBcp_Ego.Recordset("Nome_Preferido") = 1, DBcp_Ego.Recordset("Nome_Ind"), DBcp_Ego.Recordset("Nome_Nac"))
                
                If PAC1.DBcp_Ego.Recordset("Sexo") = 0 Then
                    Ícone = 8
                Else
                    Ícone = 4
                End If
                If PAC1.DBcp_Ego.Recordset("Data_Falec") <> "" Then Ícone = Ícone - 1
                
                Set itmX = ListView1.ListItems.Add(, QualNome, QualNome, , Ícone)
                ListView1.ListItems.Item(QualNome).Tag = PAC1.DBcp_Ego.Recordset("ID_Ego")
                DBcp_Ego.Recordset.FindNext Critério
            Loop
            Nome_Nac(2).Enabled = True
        Else
            Nome_Nac(2).Enabled = False
        End If
        'Coloca o nome do homem no combo e ajusta o ícone.
        ID_Nome_DBcombo11 = DBcp_Casais.Recordset("ID_Conj1")
        DBcp_Ego.Recordset.FindFirst "ID_Ego=" & DBcp_Casais.Recordset("ID_Conj1")
        Figura = IIf(IsNull(DBcp_Ego.Recordset("data_falec")) = True, 10, 9)
        Membro_Família(1).Picture = ImageList2.ListImages(Figura).Picture
        TreeView2_Clicou = 1
        DBcombo11_clicou = 1
        DBcombo12_clicou = 1
        DBcombo13_clicou = 1
        If DBcp_Ego.Recordset("Nome_Preferido") = 1 Then
            Nome_Nac(0).Value = False
        Else
            Nome_Nac(0).Value = True
        End If
        
        'Coloca o nome da mulher no combo e ajusta o ícone.
        ID_Nome_DBcombo12 = DBcp_Casais.Recordset("ID_Conj2")
        DBcp_Ego.Recordset.FindFirst "ID_Ego=" & DBcp_Casais.Recordset("ID_Conj2")
        Figura = IIf(IsNull(DBcp_Ego.Recordset("data_falec")) = True, 12, 11)
        Membro_Família(2).Picture = ImageList2.ListImages(Figura).Picture

        If DBcp_Ego.Recordset("Nome_Preferido") = 1 Then
            Nome_Nac(1).Value = False
        Else
            Nome_Nac(1).Value = True
        End If
        If ListView1.ListItems.Count <> 0 Then '(0).ListCount <> 0 Then
            Call ListView1_ItemClick(ListView1.SelectedItem)
        Else
            DBCombo(13).Text = ""
        End If
        TreeView2_Clicou = 0
        EstouCarregando = 0
    End If
End Sub

Private Sub DBcp_Ego_Reposition()
'Quando DBcp_Ego muda de um registro para outro, por seleção do nome indígena ou nome nacional,
'as opções de Sexo e Estado Civil precisam ser ajustadas na interface da acordo com o corrente Ego.
    
    If EstouCarregando = 0 And BancoCP_EGO = 1 Then 'Se o PAC não está sendo carregado agora, então...
        Call Sexo_Civil 'Chama a rotina que ajusta o sexo e o estado civil do corrente Ego.
    End If

End Sub


Private Sub DBGrid_AfterDelete(Index As Integer)
    DB_Temp.RecordSource = "Termos_Tec" 'Configura o DB_Temp para esta tabela.
    DB_Temp.Refresh 'Reinicia o DB_Temp
    DB_Temp.Recordset.FindFirst "ID_Termo_Tec=" & CInt(DB_Temp.Tag) 'Procura o termo a ser apagado pelo indice guardado em db_temp.tag
    If DB_Temp.Recordset.NoMatch = False Then DB_Temp.Recordset.Delete
End Sub

Private Sub DBGrid_AfterInsert(Index As Integer)
'Lança o novo termo iserido pelo usuário no banco de dados para ser pesquisado em todos_
'os ambientes (Masculino-Referência e tratamento, Feminino-Referência e tratamento.
'        Dim Novo_Termo_Inserido As Integer 'Declara a variável que receberá o ID do novo termo.
'        Dim Tip As Integer 'Declara a variável do Tipo de termo de tratamento.
'        Dim Sex As Integer 'Declara a variável do Sexo do ego.
'        Dim Vai As Integer 'Declara a variável de controle do loop
'        DB_Temp.RecordSource = "Termos_Tec" 'Configura o DB_Temp para esta tabela.
'        DB_Temp.Refresh 'Reinicia o DB_Temp
'        DB_Temp.Recordset.MoveLast 'Vai para o último registro, pois ele foi o recém inserido.
'        Novo_Termo_Inserido = DB_Temp.Recordset("ID_Termo_Tec")  'Pega aqui o ID do registro.
'        DB_Temp.RecordSource = "Termos_Confirmados" 'Configura o DB_Temp para esta tabela.
'        DB_Temp.Refresh 'Reinicia o DB_Temp
        
'        For Vai = 1 To 4 'O loop vai rodar 4 vezes.
'            Tip = IIf(Vai < 3, 0, 1) 'Nas duas primeiras voltas a variável Tip valerá 0, depois passará a valer 1
'            Sex = IIf(Vai = 1 Or Vai = 3, 1, 0) 'Na volta 1 e 3 a variável Sex valerá 1, nas outras voltas valerá 0
'            DB_Temp.Recordset.AddNew
'            DB_Temp.Recordset("ID_Termo_Tec") = Novo_Termo_Inserido
'            DB_Temp.Recordset("ID_Tipo") = Tip
'            DB_Temp.Recordset("Sexo_Ego") = Sex
'            DB_Temp.Recordset.Update
'        Next Vai
        
'        Label(33) = DBpa_Termos.Recordset.RecordCount 'Coloca o número de registros do DBpa no label 33.
End Sub

Private Sub DBGrid_BeforeColEdit(Index As Integer, ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    Select Case Index
        Case 1  'Não permite que os 305 termos básicos sejam editados.
                'Os termos inseridos pelo usuário iniciam com o registro número 306.
            
            If DBGrid(1).AddNewMode <> 1 Then 'AddNewMode = 1 indicar que um novo registro está sendo lançado.
                If DBpa_Termos.Recordset("id_termo") < 306 Then Cancel = True
            End If
    End Select


End Sub

Private Sub DBGrid_BeforeDelete(Index As Integer, Cancel As Integer)
    Select Case Index
        Case 1  'Não permite que os 305 termos básicos sejam apagados.
                'Os termos inseridos pelo usuário iniciam com o registro número 306.
            If DBpa_Termos.Recordset("id_termo") < 306 Then
                Cancel = True
            Else
                DB_Temp.Tag = CStr(DBpa_Termos.Recordset("ID_Termo_Tec"))
            End If
    End Select
End Sub

Private Sub DBGrid_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    DBCombo(0).Text = "" 'Quando o usuário vai para um registro, o dbcombo(0) precisa ser limpo para que o usuário escolha outro termo.
End Sub

Private Sub DBGrid1_AfterUpdate()
    'Lista dos pesquisadores na janela de configuração.
    'O nome deste grid deve ser mudado para melhor identificação.
    DBCombo(2).ReFill 'Campo pesquisador em casas/pessoal/ego
    DBCombo(10).ReFill 'Campo pesquisador em casas/pessoal/família nuclear
End Sub

Private Sub Form_Activate()
    'Depois que o form tórnasse ativo, esta variável avisa que
    'o form não mais está sendo carregado.
    EstouCarregando = 0
    EgoAtualizado = 1
    TreeView1.Style = tvwTreelinesPictureText ' Estilo 5.
    PAC1.TreeView1.ImageList = PAC1.ImageList1
    PAC1.TreeView1.Indentation = 0
    TreeView2.Style = tvwTreelinesPictureText ' Estilo 5.
    PAC1.TreeView2.ImageList = PAC1.ImageList1
    If BancoCP_EGO = 1 Then
        DBcp_Ego_Nomes.RecordSource = "select * from EGO" ' where Nome_Ind<>''"   'Ego.ID_Ego, Ego.Nome_Ind, Ego.Nome_Nac, Ego.Nome_Preferido, Ego.ID_Pais, Ego.Sexo, Ego.Data_Falesc, Ego.Civil
        DBcp_Ego_Nomes.Refresh
        SSOption(3).Value = True
        
        If BancoCP_CASAIS = 1 Then Call Familias
    
        DBcp_Ego_Nomes.RecordSource = "select * from EGO  where Nome_Ind<>''"   'Ego.ID_Ego, Ego.Nome_Ind, Ego.Nome_Nac, Ego.Nome_Preferido, Ego.ID_Pais, Ego.Sexo, Ego.Data_Falesc, Ego.Civil
        DBcp_Ego_Nomes.Refresh
   
    End If
    Mapa_Parente.Width = SSPanel(2).Width: Mapa_Parente.Height = SSPanel(2).Height 'Ajusta o tamanho do mapa_parente
    Mapa_Parente.Top = 0: Mapa_Parente.Left = 0 'Ajusta a posição do mapa_patente
    Mapa_Termo.Top = 0: Mapa_Termo.Left = 0 'Ajusta a posição do mapa_termo

    'Seleciona o tipo de termo de parentesco para "referência".
    Combo(2).ListIndex = 0 'Referência
    Combo(4).ListIndex = 0
    Call Lista_Termos 'Chama o processo que enche a lista de termos de parentesco.
    'Combo(13).ToolTipText = "Name of the location where Ego lives. this can be a village, or the name of a geographic place."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    KeyCode = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim Critério As String
    Dim Tribo As String
    Dim Aldeia As String
    Dim Localização As String
    'Prepara as variáveis para o tamanho máximo esperado
    Tribo = Space$(50)
    Aldeia = Space$(50)
    Localização = Space$(200)
    
    On Error Resume Next 'Qualquer erro que ocorrer, o PAC seguirá para a próxima linha.
    
    Call CentraForm(Me) 'Chama a rotina que centraliza o formulário.
    
    EstouCarregando = 1 'Informa para algumas rotinas que o form está sendo carregado.
    
'Esta próxima seção lê do PAC.ini valores de configuração do programa.
    
    Lingua = GetPrivateProfileInt("IDIOMA", "Idioma", 0, "PAC.ini") 'O valor que representa o idioma que será usado na interface e armazena na variável Lingua.
    Text(13).Text = GetPrivateProfileInt("CONFIRMAÇÃO", "Pergunta Ampliada", 2, "PAC.ini") 'O números de confirmações para as perguntas ampliadas
    Text(14).Text = GetPrivateProfileInt("CONFIRMAÇÃO", "Parentesco", 2, "PAC.ini") 'O número de confirmações para os termos de parentesco.
    r = GetPrivateProfileString("INFORMAÇÃO", "Tribo", "", Tribo, Len(Tribo), "PAC.ini") 'Qual a tribo.
    Text(15).Text = Left$(Tribo, InStr(Tribo, Chr$(0)) - 1)
    r = GetPrivateProfileString("INFORMAÇÃO", "Aldeia", "", Aldeia, Len(Aldeia), "PAC.ini") 'Qual a aldeia.
    Text(16).Text = Left$(Aldeia, InStr(Aldeia, Chr$(0)) - 1)
    r = GetPrivateProfileString("INFORMAÇÃO", "Localização", "", Localização, Len(Localização), "PAC.ini") 'Qual a localização da aldeia.
    Text(17).Text = Left$(Localização, InStr(Localização, Chr$(0)) - 1)
    
    Call Idioma(Lingua) 'Chama a rotina que controla o idioma e passa o valor contido em Lingua.
    
    'Ativa o item no combo(7) correspondente ao idioma
    Select Case Lingua
            Case 0 'Português
                Combo(7).ListIndex = 0
            Case 2300 'Inglês
                Combo(7).ListIndex = 1
    End Select
    
    'Inicia o Banco de Dados sobre o EGO.
    PAC1.DBcp_Ego.Refresh
    If PAC1.DBcp_Ego.Recordset.EOF <> True And PAC1.DBcp_Ego.Recordset.BOF <> True Then
        Call Enche_Combos   'Chama a rotina que enche os combos relatívos ao Ego.
        Combo(9).ListIndex = 0 'Ativa o primeiro nome na lista de Nomes Indígenas.
        'Combo(9).Tag = Combo(9).ListIndex
        
        Critério = "Nome_Ind='" & Combo(9).List(0) & "'" 'Procura o nome indígena ativo da lista no Banco de Dados.
        DBcp_Ego.Recordset.FindFirst Critério
        Call Sexo_Civil 'Chama a rotina que ativa as opções sobre o sexo e o estado civil do corrente Ego.
        
        'Chama a rotina que preenche os Combos e os Texts com os dados do corrente Ego.
        Call Procura_Ego(PAC1.Combo(9).ItemData(PAC1.Combo(9).ListIndex))
        
        'Combo(9).ListIndex + 1 Registra no Label Total, qual é o índice na lista do corrente nome indígena e qual o número total de nomes na lista.
        Total.Caption = DBcp_Ego.Recordset("ID_Ego") & " - " & DBcp_Ego.Recordset.RecordCount
        
        'Inicia o Banco de Dados sobre os CASAIS.
        PAC1.DBcp_Casais.Refresh
        If PAC1.DBcp_Casais.Recordset.EOF <> True And PAC1.DBcp_Casais.Recordset.BOF <> True Then
            
            Do Until PAC1.DBcp_Casais.Recordset.EOF = True 'Loop até que chegar ao último registro de DBcp_Ego.
                
                'Ajudante.
                'Chama a rotina que testa se o corrente item de PAC1.DBcp_Casais já existe no Combo(15)
                Call Testa_Item(PAC1.DBcp_Casais, PAC1.Combo(15), "Ajudante")
                
                PAC1.DBcp_Casais.Recordset.MoveNext 'Move para o próximo registro em DBcp_Casais.
            
            Loop
            
            PAC1.DBcp_Casais.Recordset.MoveFirst 'Move para o primeiro registro em DBcp_Casais.
            Combo(15).Text = DBcp_Casais.Recordset("Ajudante") 'Ajusta o nome do Ajudante do corrente registro.
            BancoCP_CASAIS = 1 'Indica que a tabela CASAIS já contém algum dado.
        Else
            BancoCP_CASAIS = 0 'Indica que a tabela CASAIS está vazia.
        End If
        
        BancoCP_EGO = 1 'Indica que a tabela EGO já contém algum dado.
    Else
        SSCommand(1).Enabled = False 'Desliga o botão "Editar"
        BancoCP_EGO = 0 'Indica que a tabela EGO está vazia.
        Total.Caption = "0 - 0"
    End If
End Sub

Private Sub ListView_Click(Index As Integer)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Novo_Termo
End Sub

Private Sub HScroll_Parente_Change()
    Mapa_Parente.Left = HScroll_Parente.Value * -1
End Sub

Private Sub HScroll_Parente_Scroll()
    Mapa_Parente.Left = HScroll_Parente.Value * -1
End Sub

Private Sub HScroll1_Change()
    Mapa_Termo.Left = HScroll1.Value * -1
End Sub

Private Sub HScroll1_Scroll()
    Mapa_Termo.Left = HScroll1.Value * -1
End Sub

Private Sub List_Click(Index As Integer)
    Select Case Index
        Case 0
            EstouCarregando = 1
            DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & List(0).ItemData(List(0).ListIndex)
            
            If DBcp_Ego_Nomes.Recordset("Nome_Preferido") = 1 Then
                Nome_Nac(2).Value = False
                DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & List(0).ItemData(List(0).ListIndex)
                DBCombo(13).Text = DBcp_Ego_Nomes.Recordset("Nome_Ind")
                ID_Nome_DBcombo13 = DBcp_Ego_Nomes.Recordset("ID_Ego")
            Else
                Nome_Nac(2).Value = True
                DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & List(0).ItemData(List(0).ListIndex)
                DBCombo(13).Text = DBcp_Ego_Nomes.Recordset("Nome_Nac")
                ID_Nome_DBcombo13 = DBcp_Ego_Nomes.Recordset("ID_Ego")
            End If
            EstouCarregando = 0
        Case 5
            Dim Nome As String
            EstouCarregando = 1
            List(6).Clear 'Limpa a lista de parentes
            Trilha = DBCombo(0).BoundText
            If Right(Trilha, 1) = "5" Then 'Filho
                For Qt = 1 To CInt(Arvore(List(5).ItemData(List(5).ListIndex), 0, 0)) '1000
                    If Arvore(List(5).ItemData(List(5).ListIndex), Qt, 1) <> "" Then
                        List(6).AddItem Arvore(List(5).ItemData(List(5).ListIndex), Qt, 1)
                    Else
                        Exit For
                    End If
                Next Qt
            ElseIf Trilha = "3" Then 'Irmão
                For Qt = 1 To 1000
                    If Arvore(List(5).ItemData(List(5).ListIndex), Qt, 1) <> "" Then
                        List(6).AddItem Arvore(List(5).ItemData(List(5).ListIndex), Qt, 1)
                    Else
                        Exit For
                    End If
                Next Qt
            Else
                erere = Arvore(0, 0, 0)
                For Qt = 1 To Arvore(List(5).ItemData(List(5).ListIndex), 0, 0) 'Todos os outros
                    If Arvore(List(5).ItemData(List(5).ListIndex), Qt, 1) <> "" Then
                        List(6).AddItem Arvore(List(5).ItemData(List(5).ListIndex), Qt, 1)
                    Else
                        Exit For
                    End If
                Next Qt
            End If
            List(6).ListIndex = 0 'Seleciona o primeiro parente na lista.
            EstouCarregando = 0
        Case 6
            Call Mapinha(DBCombo(0).BoundText, List(6).Text)
    End Select
End Sub


Private Sub List_DblClick(Index As Integer)
    Select Case Index
        Case 4
            Dim contando As Integer
            SSTab(2).Tab = 0
            DBcp_Ego.Recordset.FindFirst "ID_Ego = " & List(4).ItemData(List(4).ListIndex)
            If DBcp_Ego.Recordset("Nome_Ind") <> "" Then
                For contando = 0 To Combo(9).ListCount
                    If Combo(9).List(contando) = DBcp_Ego.Recordset("Nome_Ind") Then
                        Combo(9).ListIndex = contando
                    End If
                Next contando
            Else
                For contando = 0 To Combo(10).ListCount
                    If Combo(10).List(contando) = DBcp_Ego.Recordset("Nome_Nac") Then
                        Combo(10).ListIndex = contando
                    End If
                Next contando
            End If
        End Select
End Sub


Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
    'EstouCarregando = 1
    'DBcp_Ego.Recordset.FindFirst "ID_Ego=" & Item.Tag
    'If DBcp_Ego.Recordset("Nome_Preferido") = 1 Then
        'Nome_Nac(2).Value = False
        'DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & Item.Tag
        'ID_Nome_DBcombo13 = DBcp_Ego.Recordset("ID_Ego")
    'Else
        'Nome_Nac(2).Value = True
        'DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & Item.Tag
        'DBcp_Ego.Recordset.FindFirst "ID_Ego=" & Item.Tag
        'ID_Nome_DBcombo13 = DBcp_Ego.Recordset("ID_Ego")
    'End If
    'ID_Nome_DBcombo13 = DBcp_Ego.Recordset("ID_Ego")
    DBcombo13_clicou = 1
    'EstouCarregando = 0
    DBCombo(13).BoundText = Str(Item.Tag)
    ID_Nome_DBcombo13 = Str(Item.Tag)
    Call DBCombo_Click(13, 2)
End Sub


Private Sub Mapa_Parente_Resize()
Debug.Print Mapa_Parente.Height
            HScroll_Parente.Max = Mapa_Parente.Width - SSPanel(2).Width
            VScroll_Parente.Max = Mapa_Parente.Height - SSPanel(2).Height
End Sub


Private Sub Nome_Nac_Click(Index As Integer, Value As Integer)
    Dim ÍndiceAnterior As Integer
    Dim clicou As Integer
    Select Case Index
        
        Case 0
            ÍndiceAnterior = ID_Nome_DBcombo11
            If Value = False Then
                DBcp_Ego_Masculino.RecordSource = "Select * from EGO where sexo=0 And Ego.Nome_Ind<>'' Order by Ego.Nome_Ind"
                DBcp_Ego_Masculino.Refresh
                DBCombo(11).ListField = "Nome_Ind"
                clicou = DBcombo11_clicou
                ID_Nome_DBcombo11 = Ajustar_Pessoa(clicou, DBcp_Ego_Masculino, DBCombo(11), "Nome_Ind", ÍndiceAnterior, Membro_Família(1))
            Else
                DBcp_Ego_Masculino.RecordSource = "Select * from EGO where sexo=0 And Ego.Nome_Nac<>'' Order by Ego.Nome_Nac"
                DBcp_Ego_Masculino.Refresh
                DBCombo(11).ListField = "Nome_Nac"
                clicou = DBcombo11_clicou
                ID_Nome_DBcombo11 = Ajustar_Pessoa(clicou, DBcp_Ego_Masculino, DBCombo(11), "Nome_Nac", ÍndiceAnterior, Membro_Família(1))
            End If
        
        Case 1
            ÍndiceAnterior = ID_Nome_DBcombo12
            If Value = False Then
                DBcp_Ego_Feminino.RecordSource = "Select * from EGO where sexo=1 And Ego.Nome_Ind<>'' Order by Ego.Nome_Ind"
                DBcp_Ego_Feminino.Refresh
                DBCombo(12).ListField = "Nome_Ind"
                clicou = DBcombo12_clicou
                ID_Nome_DBcombo12 = Ajustar_Pessoa(clicou, DBcp_Ego_Feminino, DBCombo(12), "Nome_Ind", ÍndiceAnterior, Membro_Família(2))
            Else
                DBcp_Ego_Feminino.RecordSource = "Select * from EGO where sexo=1 And Ego.Nome_Nac<>'' Order by Ego.Nome_Nac"
                DBcp_Ego_Feminino.Refresh
                DBCombo(12).ListField = "Nome_Nac"
                clicou = DBcombo12_clicou
                ID_Nome_DBcombo12 = Ajustar_Pessoa(clicou, DBcp_Ego_Feminino, DBCombo(12), "Nome_Nac", ÍndiceAnterior, Membro_Família(2))
            End If
        
        Case 2
            ÍndiceAtual = ID_Nome_DBcombo13
            If Value = False Then
                DBcp_Ego_Nomes.RecordSource = "select * from EGO where Nome_Ind<>'' order by Ego.Nome_Ind"
                DBcp_Ego_Nomes.Refresh
                DBCombo(13).ListField = "Nome_Ind"
                clicou = DBcombo13_clicou
                ID_Nome_DBcombo13 = Ajustar_Pessoa(clicou, DBcp_Ego_Nomes, DBCombo(13), "Nome_Ind", ÍndiceAtual, Membro_Família(3))
            Else
                DBcp_Ego_Nomes.RecordSource = "select * from EGO where Ego.Nome_Nac<>'' order by Ego.Nome_Nac"
                DBcp_Ego_Nomes.Refresh
                DBCombo(13).ListField = "Nome_Nac"
                clicou = DBcombo13_clicou
                ID_Nome_DBcombo13 = Ajustar_Pessoa(clicou, DBcp_Ego_Nomes, DBCombo(13), "Nome_Nac", ÍndiceAtual, Membro_Família(3))
            End If
    End Select
End Sub

Private Sub SSCheck_Click(Index As Integer, Value As Integer)
    Select Case Index
        Case 3
            If Value = False Then
                MaskCaixa(4).Text = "__/__/____"
                MaskCaixa(5).Text = "__/__/____"
            End If
        Case 4
            If Value = False Then
                MaskCaixa(6).Text = "__/__/____"
                MaskCaixa(7).Text = "__/__/____"
            End If
    End Select
End Sub

Private Sub SSCommand_Click(Index As Integer)
    
    Dim ConstIdioma As String 'Identifica qual a língua usada na interface.
    
    Select Case Index
        
        Case 0 'Acrestando um novo registro(Ego).
            
            'Se o botão "Novo" estiver com o texto "Novo", então...
            'O motivo disto, é que o botão pode estar com o texto "Cancelar" e aí as ações devem ser outras.
            If SSCommand(0).Caption = LoadResString(131 + Lingua) Then
                
                DBcp_Ego.ReadOnly = False 'Torna o dbcp_ego atualizável.
                DBcp_Ego.Refresh 'Reinicializa o dbcp_ego.
                'É necessário colocar este valor nesta variável, para evitar que as ações em combo_click
                'sejam acionadas quando o usuário estive simplesmente escolhendo um nome e não fazendo busca de Ego.
                EstouCarregando = 1
                DBcp_Ego.Recordset.AddNew 'Prepara o dbcp_ego para receber um novo registro.
                
                Text(1).Locked = False 'Destrava este controle
                Text(2).Locked = False 'Destrava este controle
                Text(3).Locked = False 'Destrava este controle
                DBCombo(2).Locked = False 'Destrava este controle
            
                Call Desliga_Tab(PAC1.SSTab(1).Tab, PAC1.SSTab(1), False) 'Chama a rotina que Liga-Desliga os tabs
                Call Desliga_Tab(PAC1.SSTab(2).Tab, PAC1.SSTab(2), False) 'Chama a rotina que Liga-Desliga os tabs
                  
                SSCommand(1).Enabled = False 'Desliga o botão "Editar"
                SSCommand(2).Enabled = True 'Liga o botão "Gravar"
                
                With MaskCaixa(0)
                    .SetFocus 'Põe o controle para a Data em foco.
                    .Text = Format(Date, "dd/mm/yyyy") 'Insere a data atual do sistema.
                End With
                Text(3).Text = "" 'Limpa este controle - OBS.
                MaskCaixa(1).Text = "__/__/____" 'Limpa este controle - Data de Nascimento
                MaskCaixa(2).Text = "__/__/____" 'Limpa este controle - Data de Falecimento
                Combo(0).Text = "" 'Limpa este controle - Lugar de Nascimento.
                Combo(9).Text = "" 'Limpa este controle - Nome Indígena.
                Combo(10).Text = "" 'Limpa este controle - Nome Nacional.
                Combo(11).Text = "" 'Limpa este controle - Número da Casa.
                Combo(12).Text = "" 'Limpa este controle - Clã.
                Combo(13).Text = "" 'Limpa este controle - Lugar que Mora.
                Combo(14).Text = "" 'Limpa este controle - Ajudante.
                DBCombo(2).Text = "" 'Limpa este controle - Pesquisador.
                
                SSOption(1).Value = False 'Masculino. Limpa este controle para que o usuário seja obrigado a selecionar o sexo do Ego.
                SSOption(2).Value = False 'Feminino. Limpa este controle para que o usuário seja obrigado a selecionar o sexo do Ego.
                SSOption(3).Enabled = False
                SSOption(4).Enabled = False
                'SSOption(3).Value = True 'Liga a opção de Orientação.
                'SSOption(4).Value = False 'Desliga a opção de Procriação.
                PAC1.TreeView1.Nodes.Clear 'Limpa a vista de qualquer família.
                SSOption(5).Value = True 'Ajusta o estado civil para "solteiro" através deste controle.
                
                SexoEgo = 1000 'Isto significa que nenhum sexo foi selecionado. Pode ser qualquer valor diferente de 0 ou 1.
                 
                 'Muda o texto do botão "Novo" para "Cancelar", pois é neste botão que o usuário deve clicar se ele quer cancelar.
                SSCommand(0).Caption = LoadResString(135 + Lingua)
            
            Else 'Caso o botão "Novo" não esteja com o texto "Novo"...Obs: Com certeza estará com o texto "Cancelar". O banco de dados deve conter algum dado.
                
                DBcp_Ego.Recordset.CancelUpdate 'Cancela a atualização do novo registro.
                DBcp_Ego.ReadOnly = True 'Torna o dbcp_ego não atualizável.
                DBcp_Ego.Refresh 'Reinicializa o dbcp_ego.
                
                EstouCarregando = 0 'Com este valor, o PAC libera as ações do combo_click.
                
                Text(1).Locked = True 'Trava este controle.
                Text(2).Locked = True 'Trava este controle.
                Text(3).Locked = True 'Trava este controle.
                DBCombo(2).Locked = False 'Trava este controle.
                
                Call Desliga_Tab(PAC1.SSTab(1).Tab, PAC1.SSTab(1), True) 'Chama a rotina que Liga-Desliga os tabs
                Call Desliga_Tab(PAC1.SSTab(2).Tab, PAC1.SSTab(2), True) 'Chama a rotina que Liga-Desliga os tabs
                
                SSCommand(1).Enabled = True 'Liga o botão "Editar".
                SSCommand(2).Enabled = False 'Desliga o botão "Gravar".
                SSCommand(0).Caption = LoadResString(131 + Lingua) 'Muda o texto do botão "Novo" para "Novo", já que ele estava com o texto "Cancelar".
                
                SSOption(3).Enabled = True
                SSOption(4).Enabled = True
                SSOption(3).Value = True 'Liga a opção de Orientação.
                SSOption(4).Value = False 'Desliga a opção de Procriação.
                
                'Usa o ID_Ego para colocar este ego selecionado em combo(9). Nome Indígena.
                If Banco = 1 Then Combo(9).ListIndex = DBcp_Ego.Recordset("ID_Ego")
                
            
            End If
        
        Case 1 'Editando um registro(Ego).
            'Se o botão "Editar" estiver com o texto "Editar", então...
            'O motivo disto, é que o botão pode estar com o texto "Cancelar" e aí as ações devem ser outras.
            If SSCommand(1).Caption = LoadResString(136 + Lingua) Then
                
                SSCommand(0).Enabled = False 'Desliga o botão "Novo".
                SSCommand(2).Enabled = True 'Liga o botão "Gravar".
                SSCommand(1).Caption = LoadResString(135 + Lingua) 'Muda o texto do botão "Editar" para "Cancelar".
                
                SSOption(3).Enabled = False
                SSOption(4).Enabled = False
                'SSOption(3).Value = True 'Liga a opção de Orientação.
                'SSOption(4).Value = False 'Desliga a opção de Procriação.
                
                DBcp_Ego.ReadOnly = False 'Torna o dbcp_ego atualizável.
                'dbcp_ego.Refresh
                
                Text(1).Locked = False 'Destrava este controle
                Text(2).Locked = False 'Destrava este controle
                Text(3).Locked = False 'Destrava este controle
                DBCombo(2).Locked = False 'Destrava este controle
                
                Call Desliga_Tab(PAC1.SSTab(1).Tab, PAC1.SSTab(1), False) 'Chama a rotina que Liga-Desliga os tabs
                Call Desliga_Tab(PAC1.SSTab(2).Tab, PAC1.SSTab(2), False) 'Chama a rotina que Liga-Desliga os tabs
                
                'É necessário colocar este valor nesta variável, para evitar que as ações em combo_click
                'sejam acionadas quando o usuário estive simplesmente escolhendo um nome e não fazendo busca de Ego.
                EstouCarregando = 1
                
                DBcp_Ego.Recordset.Edit 'Torna o corrente registro editável.
                
                MaskCaixa(0).SetFocus 'Põe o foco no controle da Data.
            
            Else 'Caso o botão "Editar" não esteja com o texto "Editar"...  Obs: Com certeza estará com o texto "Cancelar".
                
                DBcp_Ego.Recordset.CancelUpdate 'Cancela a edição do corrente registro
                DBcp_Ego.ReadOnly = False 'Torna o dbcp_ego não atualizável.
                Procura_Ego (CStr(DBcp_Ego.Recordset("ID_Ego"))) 'Torna o corrente registro ativo nos controles correspondentes.
                
                SSCommand(0).Enabled = True 'Liga o botão "Novo".
                SSCommand(2).Enabled = False 'Desliga o botão "Gravar".
                SSCommand(1).Caption = LoadResString(136 + Lingua) 'Muda o texto do botão "Editar" para "Editar", já que ele estava com o texto "Cancelar".
                
                SSOption(3).Enabled = True
                SSOption(4).Enabled = True
                SSOption(3).Value = True 'Liga a opção de Orientação.
                SSOption(4).Value = False 'Desliga a opção de Procriação.
                
                Text(1).Locked = True 'Trava este controle
                Text(2).Locked = True 'Trava este controle
                Text(3).Locked = True 'Trava este controle
                DBCombo(2).Locked = True 'Trava este controle
                
                Call Desliga_Tab(PAC1.SSTab(1).Tab, PAC1.SSTab(1), True) 'Chama a rotina que Liga-Desliga os tabs
                Call Desliga_Tab(PAC1.SSTab(2).Tab, PAC1.SSTab(2), True) 'Chama a rotina que Liga-Desliga os tabs
                
                EstouCarregando = 0 'Com este valor, o PAC libera as ações do combo_click.
                
                'DBcp_Ego.Recordset.CancelUpdate 'Cancela a edição do corrente registro
                
                'DBcp_Ego.ReadOnly = False 'Torna o dbcp_ego não atualizável.
                'Procura_Ego (CStr(DBcp_Ego.Recordset("ID_Ego"))) 'Torna o corrente registro ativo nos controles correspondentes.
            
            End If
        
        Case 2 'Gravando um registro(Ego).
  
            If Dados_Válidos = True Then 'Se a função Dados_Válidos retorna o valor True, então...  O processo só continua se os campos estiverem preenchidos corretamente.
                Call Grava_Ego 'Chama a rotina que grava o corrente registro no DB_Ego.
                'CorrenteEgo = DBcp_Ego.Recordset("ID_Ego")
                Combo(9).Clear 'Limpa o combo Nome Indígena.###
                Combo(10).Clear 'Limpa o combo Nome Nacional.###
                Combo(13).Clear 'Limpa o combo Lugar Mora.###
                Combo(11).Clear 'Limpa o combo Número da Casa.###
                Combo(14).Clear 'Limpa o combo Ajudante.###
                Combo(0).Clear 'Limpa o combo Lugar Nascimento.###
                Combo(12).Clear 'Limpa o combo Clã.###
                Call Enche_Combos 'chama a rotina que enche os combos relevantes para Casas-Pessoal.
                EstouCarregando = 0 'Com este valor, o PAC libera as ações do combo_click.
                DBcp_Ego.Recordset.FindFirst "ID_Ego=" & CorrenteEgo
                Dim contando As Integer
                Dim CorrenteItem As Integer
                
                If PAC1.DBcp_Ego.Recordset("Nome_Preferido") = 1 Then
                    For contando = 0 To Combo(9).ListCount
                        Combo(9).Text = Combo(9).List(contando)
                        If Combo(9).Text = DBcp_Ego.Recordset("Nome_Ind") Then
                            CorrenteItem = contando
                            Exit For
                        End If
                    Next contando
                Else
                    For contando = 0 To Combo(10).ListCount
                        Combo(10).Text = Combo(10).List(contando)
                        If Combo(10).Text = DBcp_Ego.Recordset("Nome_Nac") Then
                            CorrenteItem = contando
                            Exit For
                        End If
                    Next contando
                End If
                
                If DBcp_Ego.Recordset("Nome_Ind") <> "" Then
                    Combo(9).ListIndex = CorrenteItem  'CorrenteEgo 'Esta pedindo um indice que nao foi lancado no combo.
                Else
                    Combo(10).ListIndex = CorrenteItem 'CorrenteEgo
                End If
 
                DBcp_Ego.ReadOnly = True 'Torna o dbcp_ego atualizável.
                
                'Se o campo "nome_Ind" do corrente registro não for vazio, então este valor é colocado em Combo(9).text
                If PAC1.DBcp_Ego.Recordset("nome_Ind") <> "" Then Combo(9).Text = PAC1.DBcp_Ego.Recordset("nome_Ind")
                
                'Se o campo "nome_nac" do corrente registro não for vazio, então este valor é colocado em Combo(10).text
                If PAC1.DBcp_Ego.Recordset("nome_Nac") <> "" Then Combo(10).Text = PAC1.DBcp_Ego.Recordset("nome_Nac")
                
                'Atualiza o label Total
                Total.Caption = DBcp_Ego.Recordset("ID_Ego") & " - " & DBcp_Ego.Recordset.RecordCount
                
'                EstouCarregando = 0 'Com este valor, o PAC libera as ações do combo_click.
                
                'Se o botão "Novo" está com o texto "Cancelar" então é colocado o texto "Novo"
                If SSCommand(0).Caption = LoadResString(135 + Lingua) Then SSCommand(0).Caption = LoadResString(131 + Lingua) 'Cancelar 135
                'Se o botão "Editar" está com o texto "Cancelar" então é colocado o texto "Editar"
                If SSCommand(1).Caption = LoadResString(135 + Lingua) Then SSCommand(1).Caption = LoadResString(136 + Lingua) 'Cancelar 135
                
                SSCommand(0).Enabled = True     'Liga o botão "Novo" 131
                SSCommand(1).Enabled = True     'Liga o botão "Editar" 136
                SSCommand(2).Enabled = False    'Liga o botão "Gravar"
                SSOption(3).Enabled = True
                SSOption(4).Enabled = True
                SSOption(3).Value = True 'Liga a opção de Orientação.
                SSOption(4).Value = False 'Desliga a opção de Procriação.
                Call Desliga_Tab(PAC1.SSTab(1).Tab, PAC1.SSTab(1), True) 'Chama a rotina que Liga-Desliga os tabs
                Call Desliga_Tab(PAC1.SSTab(2).Tab, PAC1.SSTab(2), True) 'Chama a rotina que Liga-Desliga os tabs
                DBcp_Ego_Masculino.Refresh
                DBcp_Ego_Feminino.Refresh
                DBcp_Ego_Nomes.Refresh
                Call Familias

            End If

        Case 3 'Imprimindo Ego.
        
        Case 4 'Inserindo um novo filho na lista.
            Dim itmX As ListItem
            Dim QualNome As String
            'Se o DBcombo(13) não está vazio, então...
            If DBCombo(13).Text <> "" Then
                    Debug.Print DBCombo(13).Text 'DBcp_Casais.Recordset("ID_Conj1") 'Val(DBCombo(11).BoundText); Val(DBCombo(13).BoundText)
                    'Debug.Print DBcp_Casais.Recordset("ID_Conj2"); Val(DBCombo(13).BoundText)
                'Se o ID do DBCombo(13).pai for diferente do ID do filho e o ID da mãe for diferente do ID do filho, então...
                '##########
                ID_combo11 = Val(DBCombo(11).BoundText)
                ID_combo12 = Val(DBCombo(12).BoundText)
                If ID_combo11 <> ID_Nome_DBcombo13 And ID_combo12 <> ID_Nome_DBcombo13 Then
                'If DBcp_Ego_Masculino.Recordset("ID_Ego") <> ID_Nome_DBcombo13 And DBcp_Ego_Feminino.Recordset("ID_Ego") <> ID_Nome_DBcombo13 Then
                    
                    
                    'Se o ListView1 não estiver vazio, então...
                    If ListView1.ListItems.Count <> 0 Then
                        For contaitens = 1 To ListView1.ListItems.Count - 1
                            If ListView1.ListItems.Item(contaitens).Tag = ID_Nome_DBcombo13 Then achou = 1
                            Debug.Print ListView1.ListItems.Item(contaitens).Tag; " - "; ID_Nome_DBcombo13
                        Next contaitens
                    Else
                        achou = 0
                    End If
                    
                    'Se o list(0) não estiver vazio, então...
'                    If List(0).ListCount <> 0 Then
'                        For contaitens = 0 To List(0).ListCount - 1
'                            If List(0).ItemData(contaitens) = ID_Nome_DBcombo13 Then achou = 1
'                            Debug.Print List(0).ItemData(contaitens); " - "; ID_Nome_DBcombo13
'                        Next contaitens
'                    Else
'                        achou = 0
'                    End If
                    'Se não achou a corrente pessoal do DBCombo(13) na list(0), então...
                    If achou = 0 Then
'                        List(0).AddItem DBCombo(13).Text
'                        List(0).ItemData(List(0).NewIndex) = ID_Nome_DBcombo13 'Val(DBCombo(13).BoundText)
                    
                        DBcp_Ego.Recordset.FindFirst "ID_Ego=" & ID_Nome_DBcombo13
                        If DBcp_Ego.Recordset.NoMatch = False Then
                            QualNome = IIf(DBcp_Ego.Recordset("Nome_Preferido") = 1, DBcp_Ego.Recordset("Nome_Ind"), DBcp_Ego.Recordset("Nome_Nac"))
                            
                            If PAC1.DBcp_Ego.Recordset("Sexo") = 0 Then
                                Ícone = 8
                            Else
                                Ícone = 4
                            End If
                            If PAC1.DBcp_Ego.Recordset("Data_Falec") <> "" Then Ícone = Ícone - 1
                            
                            Set itmX = ListView1.ListItems.Add(, QualNome, QualNome, , Ícone)
                            ListView1.ListItems.Item(QualNome).Tag = PAC1.DBcp_Ego.Recordset("ID_Ego")
                        End If
                
                    Else
                        Beep
                        Mensagem_Erro = MsgBox(LoadResString(288 + Lingua), 48, LoadResString(282 + Lingua)) 'Atenção!
                        DBCombo(13).SetFocus
                    End If
                Else
                    Beep
                    Mensagem_Erro = MsgBox(LoadResString(287 + Lingua), 48, LoadResString(282 + Lingua)) 'Atenção!
                    DBCombo(13).SetFocus
                End If
            Else
                Beep
            End If
        
        Case 5 'Apagando filhos da lista.
            'Dim QualNome As String
            If ListView1.SelectedItem <> "" Then
                'QualNome = IIf(DBcp_Ego.Recordset("Nome_Preferido") = 1, DBcp_Ego.Recordset("Nome_Ind"), DBcp_Ego.Recordset("Nome_Nac"))
                it(contador) = ListView1.SelectedItem.Tag 'Guarda na matris it o índice do ego. ID_Nome_DBcombo13
                contador = contador + 1
                ListView1.ListItems.Remove ListView1.SelectedItem.Index    'List(0).RemoveItem List(0).ListIndex
            End If
            
'            If List(0).ListIndex <> -1 Then
'                it(contador) = List(0).ItemData(List(0).ListIndex)
'                contador = contador + 1
'                List(0).RemoveItem List(0).ListIndex
'            End If
        
        
        Case 6 'Gravando a familia no DBcp_Casais e o ID_Casal com cada filho em DBcp_Ego
            If DBCombo(10).Text <> "" Then 'Se o campo estiver vazio, então...
    
                    If IsNumeric(Text(5).Text) = True Then
                        Dim CasalAtual As Integer
                        If DBcp_Casais.Recordset.EditMode = dbEditAdd Then
                            CasalAtual = DBcp_Casais.Recordset.RecordCount + 1
                            DBcp_Casais.Recordset("ID_Casal") = DBcp_Casais.Recordset.RecordCount + 1
                        Else
                            CasalAtual = DBcp_Casais.Recordset("ID_Casal")
                        End If
                        DBcp_Casais.Recordset("Ajudante") = Combo(15).Text
                        
                        'Tirando os egos da familia
                        If contador <> 0 Then
                            For contando = 0 To contador - 1
                                DBcp_Ego.Recordset.FindFirst "ID_Ego =" & it(contando)
                                If DBcp_Ego.Recordset.NoMatch = False Then
                                    DBcp_Ego.Recordset.Edit
                                    DBcp_Ego.Recordset("ID_Pais") = 0
                                    DBcp_Ego.Recordset.Update
                                End If
                            Next contando
                        End If
                        
                        'Incluindo os egos na familia
                        For contaitem = 1 To ListView1.ListItems.Count
                            Set ListView1.SelectedItem = ListView1.ListItems(contaitem)
                            DBcp_Ego.Recordset.FindFirst "ID_Ego =" & ListView1.SelectedItem.Tag  '(0).ItemData(contaitem)
                            If DBcp_Ego.Recordset.NoMatch = False Then
                                'Dim NomePreferido As Byte
                                DBcp_Ego.Recordset.Edit
                                'NomePreferido = IIf(DBcp_Ego.Recordset("Nome_Ind") = List(0).List(contaitem), 1, 2) ' 1 = Nome_Ind    2 = Nome_Nac
                                DBcp_Ego.Recordset("ID_Pais") = CasalAtual
                                'DBcp_Ego.Recordset("Nome_Preferido") = NomePreferido
                                DBcp_Ego.Recordset.Update
                            End If
                        Next contaitem
                        'Registra o estado civil do casal através desta variável, que é alterada
                        'quando o usuário clica no símbolo de relacionamento do casal.
                        DBcp_Casais.Recordset("Civil") = Civil
                        DBcp_Casais.Recordset.Update
                        BancoCP_CASAIS = 1 'Indica que a tabela CASAIS no DBCP contém algum dado.
                        
                        Call Familias
                        EstouCarregando = 0
                        DBcp_Casais.Recordset.FindFirst "ID_Casal=" & CasalAtual
                        'Se o botão "Novo" está com o texto "Cancelar" então é colocado o texto "Novo"
                        If SSCommand(32).Caption = LoadResString(135 + Lingua) Then SSCommand(32).Caption = LoadResString(131 + Lingua) 'Cancelar 135
                        'Se o botão "Editar" está com o texto "Cancelar" então é colocado o texto "Editar"
                        If SSCommand(33).Caption = LoadResString(135 + Lingua) Then SSCommand(33).Caption = LoadResString(136 + Lingua) 'Cancelar 135
                        
                        TreeView2.Enabled = True
                        SSCommand(32).Enabled = True     'Liga o botão "Novo" 131
                        SSCommand(33).Enabled = True     'Liga o botão "Editar" 136
                        SSCommand(6).Enabled = False    'Liga o botão "Gravar"
                        Call Desliga_Tab(PAC1.SSTab(1).Tab, PAC1.SSTab(1), True) 'Chama a rotina que Liga-Desliga os tabs
                        Call Desliga_Tab(PAC1.SSTab(2).Tab, PAC1.SSTab(2), True) 'Chama a rotina que Liga-Desliga os tabs
                        Call Liga_Desliga_Controle(0, 1) 'Desliga os controles.
                    Else
                        Beep
                        Mensagem_Erro = MsgBox(LoadResString(280 + Lingua) & " " & UCase(LoadResString(57 + Lingua)) & " " & LoadResString(286 + Lingua), 48, LoadResString(282 + Lingua)) 'Atenção!
                        Text(5).SetFocus
                    End If
            Else
                'Uma mensagem é dada ao usuário depedendo de "campo" ele não preencheu.
                'Todo o texto é tirado do PAC1.res, pois depende da língua em uso na interface.
                Beep
                Mensagem_Erro = MsgBox(LoadResString(280 + Lingua) & Chr(32) & Chr(34) & UCase(Label(17).Caption) & Chr(34) & Chr(32) & LoadResString(281 + Lingua), 48, LoadResString(281 + Lingua)) 'Atenção!
            End If
        
        Case 7 'Apaga a família selecionada.
            DBcp_Ego.Recordset.FindFirst "ID_Pais =" & PAC1.DBcp_Casais.Recordset("ID_Casal")
            Do While DBcp_Ego.Recordset.NoMatch = False
                DBcp_Ego.Recordset.Edit
                DBcp_Ego.Recordset("ID_Pais") = 0
                DBcp_Ego.Recordset.Update
                DBcp_Ego.Recordset.FindNext "ID_Pais =" & PAC1.DBcp_Casais.Recordset("ID_Casal")
            Loop
            ListView1.ListItems.Clear
            PAC1.DBcp_Casais.Recordset.Delete
            TreeView2.Nodes.Remove TreeView2.SelectedItem.Index    'List(0).RemoveItem List(0).ListIndex

        Case 8
            DBcp_Ego_Masculino.Refresh
        
        Case 9
            Dim MeuCritério As String
            Dim EstadoCivil As String
            
            'CASA MORA
            If Combo(16).Text <> "" Then
                MeuCritério = MeuCritério & "Casa_Mora=" & "'" & Val(Combo(16).Text) & "'" & " and "
            End If
            
            'LUGAR QUE MORA
            If Combo(17).Text <> "" Then
                MeuCritério = MeuCritério & "Lugar_Mora=" & Combo(17).ItemData(Combo(17).ListIndex) & " and "
            End If
            
            'NOME INDÍGENA
            If Combo(18).Text <> "" Then
                MeuCritério = MeuCritério & "Nome_Ind like " & "'*" & NovaString(Combo(18).Text) & "*'" & " and "
            End If
            
            'NOME NACIONAL
            If Combo(19).Text <> "" Then
                MeuCritério = MeuCritério & "Nome_Nac like " & "'*" & Combo(19).Text & "*'" & " and "
            End If
            
            'LUGAR DE NASCIMENTO
            If Combo(20).Text <> "" Then
                MeuCritério = MeuCritério & "Lugar_Nasc=" & Combo(20).ItemData(Combo(20).ListIndex) & " and "
            End If
            
            'CLÃ
            If Combo(21).Text <> "" Then
                MeuCritério = MeuCritério & "Clã like " & "'*" & Combo(21).Text & "*'" & " and "
            End If
            
            'SEXO
            If SSCheck(1).Value = True And SSCheck(2).Value = True Then
                MeuCritério = MeuCritério & "(Sexo= 0" & " or " & "Sexo= 1)" & " and "
            Else
                'MASCULINO
                If SSCheck(1).Value = True Then
                    MeuCritério = MeuCritério & "Sexo= 0" & " and "
                End If
                'FEMININO
                If SSCheck(2).Value = True Then
                    MeuCritério = MeuCritério & "Sexo= 1" & " and "
                End If
            End If
            
            'ESTADO CIVIL
            If SSCheck(5).Value = True Then
                Estado_civil = Estado_civil & "Civil=2" & " or "
            End If
            If SSCheck(6).Value = True Then
                Estado_civil = Estado_civil & "Civil=3" & " or "
            End If
            If SSCheck(7).Value = True Then
                Estado_civil = Estado_civil & "Civil=4" & " or "
            End If
            If SSCheck(8).Value = True Then
                Estado_civil = Estado_civil & "Civil=5" & " or "
            End If
            If SSCheck(9).Value = True Then
                Estado_civil = Estado_civil & "Civil=6" & " or "
            End If
                        
            If Estado_civil <> "" Then
                Estado_civil = Left(Estado_civil, Len(Estado_civil) - 4)
                MeuCritério = MeuCritério & "(" & Estado_civil & ")" & " and "
            End If
            
            'FAMÍLIA DE PROCRIAÇÃO/ORIENTAÇÃO
            If SSCheck(10).Value = True And SSCheck(11).Value = True Then
                MeuCritério = MeuCritério & "(ID_Pais<> 0" & " or " & "Civil<>2)" & " and "
            Else
                'ORIENTAÇÃO
                If SSCheck(10).Value = True Then
                    MeuCritério = MeuCritério & "ID_Pais <> 0" & " and "
                End If
                'PROCRIAÇÃO
                If SSCheck(11).Value = True Then
                    MeuCritério = MeuCritério & "Civil <> 2" & " and "
                End If
            End If

            'OBSERVAÇÃO
            If Text(6).Text <> "" Then
                MeuCritério = MeuCritério & "Obs like " & "'*" & NovaString(Text(6).Text) & "*'" & " and "
            End If
            
            'DATAS DE NASCIMENTO E/OU FALECIMENTO POR PERIODO
            'NASCIMENTO
            If SSCheck(3).Value = True Then
                Debug.Print InStr(1, MaskCaixa(4).Text, "_")
                If InStr(1, MaskCaixa(4).Text, "_") = 0 And InStr(1, MaskCaixa(5).Text, "_") <> 0 Then ' instr
                    MeuCritério = MeuCritério & "(Data_Nasc =  #" & Format(MaskCaixa(4).Text, "mm - dd - yyyy") & "#) and "
                Else
                    If InStr(1, MaskCaixa(4).Text, "_") = 0 And InStr(1, MaskCaixa(5).Text, "_") = 0 Then
                        MeuCritério = MeuCritério & "(Data_Nasc Between  #" & Format(MaskCaixa(4).Text, "mm - dd - yyyy") & "# and #" & Format(MaskCaixa(5).Text, "mm - dd - yyyy") & "#) and "
                    End If
                End If
            End If
            'FALECIMENTO
            If SSCheck(4).Value = True Then
                 If InStr(1, MaskCaixa(6).Text, "_") = 0 And InStr(1, MaskCaixa(7).Text, "_") <> 0 Then ' instr
                    MeuCritério = MeuCritério & "(Data_falec =  #" & Format(MaskCaixa(6).Text, "mm - dd - yyyy") & "#) and "
                Else
                    If InStr(1, MaskCaixa(6).Text, "_") = 0 And InStr(1, MaskCaixa(7).Text, "_") = 0 Then
                        MeuCritério = MeuCritério & "(Data_falec Between  #" & Format(MaskCaixa(6).Text, "mm - dd - yyyy") & "# and #" & Format(MaskCaixa(7).Text, "mm - dd - yyyy") & "#) and "
                    End If
                End If
            End If
            
            
            
            
            If MeuCritério <> "" Then
                MeuCritério = Left(MeuCritério, Len(MeuCritério) - 5)
                
                List(4).Clear
                
                EstouCarregando = 1
                DBcp_Ego.Refresh
                
                DBcp_Ego.Recordset.FindFirst MeuCritério
                Do Until DBcp_Ego.Recordset.NoMatch = True
                    If DBcp_Ego.Recordset.NoMatch = False Then
                        QualNome = IIf(DBcp_Ego.Recordset("Nome_Preferido") = 1, DBcp_Ego.Recordset("Nome_Ind"), DBcp_Ego.Recordset("Nome_Nac"))
                        List(4).AddItem QualNome '& "  (" & DBcp_Ego.Recordset("ID_Ego") & ")"
                        List(4).ItemData(List(4).NewIndex) = DBcp_Ego.Recordset("ID_Ego")
                    End If
                    DBcp_Ego.Recordset.FindNext MeuCritério
                Loop
                EstouCarregando = 0
            Else
                List(4).Clear
                Beep
            End If
        
        Case 11 'Acrescenta um novo Termo Técnico na lista padrão.
            Novo_Termo.Show
            
        Case 12 'Imprime a lista ou o mapa de terminologia
            Dim Na_Coluna As Integer
            Dim Na_Linha As Integer
            Dim Tipo_de_Termo As String
            Dim Sexo_do_Ego As Integer
            Dim Termo_Lingua As String
            If Combo(2).ListIndex = 0 Then Tipo_de_Termo = "Referência" Else Tipo_de_Termo = "Tratamento"
            If SSOption(12).Value = True Then 'Se a escolha masculino for selecionada...
                Sexo_do_Ego = 47 'Masculino
            ElseIf SSOption(13).Value = True Then 'Se a escolha feminino for selecionada...
                Sexo_do_Ego = 48 'Feminino
            End If

            
            If SSOption(15).Value = True Then 'Se a vista do mapa estiver selecionada, então o mapa será impresso.
                Clipboard.SetData Mapa_Termo.Image, 2 'Copia o mapa como bitmap para a memória
                Mapa_Termo.Picture = Clipboard.GetData(2) 'Coloca o mapa da memória na propriedade picture
                Printer.Orientation = 2 'Orientação da página como paisagem
                Printer.CurrentY = 0 'Posição para o título
                Printer.FontSize = 12
                Printer.Print "MAPA DE TERMINOLOGIA - " & Format(Date, "Long Date")
                Printer.FontSize = 10
                Printer.Print
                Printer.PaintPicture ImageList1.ListImages(Sexo_do_Ego).Picture, Printer.CurrentX, Printer.CurrentY, 400, 400        'O Ego masculino
                Printer.CurrentX = 500 'Ajusta a coluna.
                Printer.Print "Termo de " & Tipo_de_Termo
                Printer.Print
                Printer.PaintPicture Mapa_Termo.Picture, 0, 1000, Printer.Width - 800, (Printer.Height / 2) 'Manda o bitmap para a impressora com as dimenções ajustadas
                Printer.EndDoc 'somente depois deste comando a impressora inicia o impressão
            ElseIf SSOption(14).Value = True Then 'Se a vista da lista estiver selecionada, então a lista será impressa.
                Termo_Lingua = IIf(Lingua = 0, "Termo_Tec", "Termo_Tec_IN")
                DrawWidth = 1
                Na_Coluna = Printer.Width / 2
                Printer.Orientation = 1 'Orientação da página como retrato (página em pé)
                Printer.CurrentY = 0 'Posição para o título
                Printer.FontSize = 12
                Printer.Print "LISTA DE TERMINOLOGIA - " & Format(Date, "Long Date")
                Printer.FontSize = 10
                Printer.Print
                Printer.PaintPicture ImageList1.ListImages(Sexo_do_Ego).Picture, Printer.CurrentX, Printer.CurrentY, 400, 400        'O Ego masculino
                Printer.CurrentX = 500
                Printer.Print "Termo de " & Tipo_de_Termo
                Printer.Print
                Na_Linha = Printer.CurrentY
                Printer.Print "Termo Técnico"
                Printer.CurrentY = Na_Linha
                Printer.CurrentX = Na_Coluna
                Printer.Print "Termo Indígena"
                Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY), QBColor(0)
                Printer.FontSize = 8
                DBpa_Termos.Recordset.MoveFirst
                Do While DBpa_Termos.Recordset.EOF = False 'Continua o loop até o fim do DBpa
                'For vai = 1 To 3
                    Printer.CurrentX = 0
                    Na_Linha = Printer.CurrentY
                    Printer.Print DBpa_Termos.Recordset(Termo_Lingua)
                    Printer.CurrentY = Na_Linha
                    Printer.CurrentX = Na_Coluna
                    If IsNull(DBpa_Termos.Recordset("Termo_Ind")) = False Then
                        Printer.Print DBpa_Termos.Recordset("Termo_Ind")
                    Else
                        Printer.Print
                    End If
                    DBpa_Termos.Recordset.MoveNext
                    Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY), QBColor(0)
                'Next vai
                Loop
                DBpa_Termos.Recordset.MoveFirst
                Printer.EndDoc 'somente depois deste comando a impressora inicia o impressão
                
            End If
           
        Case 31 'Gravando a corrente configuração.
            
            If Combo(7).ListIndex = 0 Then 'Se o combo "Idioma" for igual a 0 (Português), então...
                ConstIdioma = 0 'Põe o valor 0=Português nesta variável.
            Else 'Caso o combo "Idioma" for diferente de 0...
                ConstIdioma = 2300 'Põe o valor 2300=Inglês nesta variável.
            End If
            
            Combo(7).Clear 'Limpa o combo "Idioma", pois ele vai ser peenchido com o novo idioma selecionado.
            'Grava no Pac.ini o valor do novo idioma selecionado.
            a = WritePrivateProfileString("IDIOMA", "Idioma", ConstIdioma, "pac.ini")
            Lingua = ConstIdioma 'Muda esta variável pública para o valor do novo idioma selecionado.
            Idioma (ConstIdioma) 'Muda toda a interface com o novo idioma selecionado.
        Case 32 'Acrescenta um novo registro sobre casais.
            
            'Se o texto do botão = "Novo", então...
            If SSCommand(32).Caption = LoadResString(131 + Lingua) Then
                Call Desliga_Tab(PAC1.SSTab(1).Tab, PAC1.SSTab(1), False) 'Chama a rotina que Liga-Desliga os tabs
                Call Desliga_Tab(PAC1.SSTab(2).Tab, PAC1.SSTab(2), False) 'Chama a rotina que Liga-Desliga os tabs
                
                DBcombo11_clicou = 0
                DBcombo12_clicou = 0
                DBcombo13_clicou = 0
                EstouCarregando = 1
                DBcp_Casais.Recordset.AddNew
                MaskCaixa(3).Text = Date
                DBCombo(13).Text = ""
                Combo(15).Text = ""
                ListView1.ListItems.Clear
                AniButton.Value = 2 'Mostra o símbolo de casamento
                Civil = 3 'Ajusta esta variável para indicar casamento.
                
                Call Liga_Desliga_Controle(1, 0) 'Liga os controles.
                TreeView2.Enabled = False
                'Muda o texto do botão "Novo" para "Cancelar", pois é neste botão que o usuário deve clicar se ele quer cancelar.
                SSCommand(33).Enabled = False 'Desliga o botão "Editar".
                SSCommand(6).Enabled = True 'Liga o botão "Gravar".
                SSCommand(32).Caption = LoadResString(135 + Lingua) 'Devolve o texto "Novo" para este botão.
            Else
                Call Desliga_Tab(PAC1.SSTab(1).Tab, PAC1.SSTab(1), True) 'Chama a rotina que Liga-Desliga os tabs
                Call Desliga_Tab(PAC1.SSTab(2).Tab, PAC1.SSTab(2), True) 'Chama a rotina que Liga-Desliga os tabs
                                
                Call Liga_Desliga_Controle(0, 1) 'Desliga os controles.
                
                TreeView2.Enabled = True
                SSCommand(33).Enabled = True 'Liga o botão "Editar".
                SSCommand(6).Enabled = False 'Desliga o botão "Gravar".
                SSCommand(32).Caption = LoadResString(131 + Lingua) 'Muda o texto do botão "Novo" para "Novo", já que ele estava com o texto "Cancelar".
                EstouCarregando = 0
                DBcp_Casais.Recordset.CancelUpdate
                If BancoCP_CASAIS = 1 Then 'Se existe algum dado na tabela CASAIS do DBCP, então...
                    Combo(15).Text = DBcp_Casais.Recordset("Ajudante") 'Ajusta o nome do Ajudante do corrente registro.
                    Call DBcp_Casais_Reposition
                End If
            End If
        
        Case 33 'Editando o corrente registro sobre o casal.
            
            'Se o texto do botão = "Editar", então...
            If SSCommand(33).Caption = LoadResString(136 + Lingua) Then
                Call Desliga_Tab(PAC1.SSTab(1).Tab, PAC1.SSTab(1), False) 'Chama a rotina que Liga-Desliga os tabs
                Call Desliga_Tab(PAC1.SSTab(2).Tab, PAC1.SSTab(2), False) 'Chama a rotina que Liga-Desliga os tabs
                
                Call Liga_Desliga_Controle(1, 0) 'Liga os controles.
                
                Civil = DBcp_Casais.Recordset("civil")
                'If List(0).ListCount = 0 Then
                If ListView1.ListItems.Count = 0 Then
                    DBcp_Ego_Nomes.RecordSource = "select Ego.ID_Ego, Ego.Nome_Ind,Ego.Nome_Nac,Ego.Nome_Preferido,Ego.sexo,Ego.Data_Falec from EGO where Nome_Nac<>'' order by Ego.Nome_Nac"
                    DBcp_Ego_Nomes.Refresh
                    'DBcp_Ego_Nomes.Recordset.MoveFirst
                    DBCombo(13).ListField = "Nome_ind"
                    DBCombo(13).Text = "" 'DBcp_Ego_Nomes.Recordset("Nome_ind")
                    Nome_Nac(2).Enabled = True
                    Nome_Nac(2).Value = False
                End If
                
                TreeView2.Enabled = False
                SSCommand(32).Enabled = False 'Desliga o botão "Novo".
                SSCommand(6).Enabled = True 'Liga o botão "Gravar".
                SSCommand(33).Caption = LoadResString(135 + Lingua) 'Devolve o texto "Editar" para este botão.
                
                'If List(0).ListCount <> 0 Then ReDim it(List(0).ListCount - 1)
                If ListView1.ListItems.Count <> 0 Then ReDim it(ListView1.ListItems.Count)
                contador = 1
                DBcp_Casais.Recordset.Edit
            Else
                Call Desliga_Tab(PAC1.SSTab(1).Tab, PAC1.SSTab(1), True) 'Chama a rotina que Liga-Desliga os tabs
                Call Desliga_Tab(PAC1.SSTab(2).Tab, PAC1.SSTab(2), True) 'Chama a rotina que Liga-Desliga os tabs
                
                Call Liga_Desliga_Controle(0, 1) 'Desliga os controles.
                
                TreeView2.Enabled = True
                SSCommand(32).Enabled = True 'Liga o botão "Novo".
                SSCommand(6).Enabled = False 'Desliga o botão "Gravar".
                SSCommand(33).Caption = LoadResString(136 + Lingua) 'Muda o texto do botão "Editar" para "Editar", já que ele estava com o texto "Cancelar".
                DBcp_Casais.Recordset.CancelUpdate
                Call DBcp_Casais_Reposition
            End If
            contador = 0
            
        
    End Select
End Sub

Private Sub SSOption_Click(Index As Integer, Value As Integer)
    

'Os casos do index=0,1,2,3,4,5 e 6 associa na variável correspondente o valor para o sexo ou o estado civil do Ego dependendo do SSoption que foi clicado.
    Dente = 800 '800 é valor do dente no procedimento das linhas
    PGerDesc = 2640
    SGerDesc = 3440
    GerEgo = 1840
    PGerAsc = 1040
    SGerAsc = 240
    espaco = 30000 / 14 '30000 era a largura inicial do mapa_termo.
    outro = (espaco / 3) + Dente

    Select Case Index
        Case 1
            SexoEgo = 0     'Masculino
        Case 2
            SexoEgo = 1     'Feminino
        Case 3
            Call Orientação
        Case 4
            Call Procriação
        Case 5
            CivilEgo = 2    'Solteiro
        Case 6
            CivilEgo = 3    'Casado
        Case 7
            CivilEgo = 4    'Viúvo
        Case 8
            CivilEgo = 5    'Separado
        Case 9
            CivilEgo = 6    'União Irregular
            
        Case 16
            'Se o tab Parentesco/Planejamento estiver ativo, o ssption de Parentesco/Geral será ativado.
            'É por este ssption(Parentesco/Geral ) que a rotina de listar termos é ativada.
            If SSTab(3).Tab = 1 Then
                SSOption(12).Value = True
                List(5).Clear
                List(6).Clear
                Mapa_Parente.Cls
            End If
        Case 17
            'Se o tab Parentesco/Planejamento estiver ativo, o ssption de Parentesco/Geral será ativado.
            'É por este ssption(Parentesco/Geral ) que a rotina de listar termos é ativada.
            If SSTab(3).Tab = 1 Then
                SSOption(13).Value = True
                List(5).Clear
                List(6).Clear
                Mapa_Parente.Cls
            End If
        Case "12", "13"
            If Index = 12 Then 'Aqui é escolhido qual ego vai estar em foco.
                Mapa_Termo.Line (15070 + Dente, 1650)-(15070 + Dente, 1650 + 230), QBColor(15)
                Mapa_Termo.PaintPicture ImageList1.ListImages(13).Picture, 14671 + Dente, 1840 'O Ego masculino
                Mapa_Termo.Line (14780 + Dente, 1640)-(14780 + Dente, 1640 + 230), QBColor(0)
                SSOption(16).Value = True 'Liga este option no tab Patentesco/Planejamento
            Else
                Mapa_Termo.Line (14780 + Dente, 1650)-(14780 + Dente, 1650 + 230), QBColor(15)
                Mapa_Termo.PaintPicture ImageList1.ListImages(17).Picture, 14671 + Dente, 1840 'O Ego feminino
                Mapa_Termo.Line (15070 + Dente, 1640)-(15070 + Dente, 1640 + 230), QBColor(0)
                SSOption(17).Value = True 'Liga este option no tab Patentesco/Planejamento
            End If
            Mapa_Termo.Line (14910 + Dente, PGerDesc - 200)-(14910 + Dente, PGerDesc - 700), QBColor(0)
            Call Lista_Termos 'Chama o procedimento que enche a lista com os parâmetros selecionados.
            DBCombo(0).Text = "" 'Limpa o combo(0) no tab Parestesco/Planejamento
            If SSOption(15).Value = True Then Call SSOption_Click(15, True)
        Case 14 'A vista da lista é selecionada. Torna os controles ligados à vista de lista de termos visíveis.
            DBGrid(1).Visible = True
            SSPanel(0).Visible = False
            HScroll1.Visible = False
            Call Lista_Termos 'Chama o procedimento que enche a lista com os parâmetros selecionados.
        Case 15 'A vista do mapa é selecionada. Torna os controles ligados à vista de mapa de termos visíveis.
            'MaCs - Talvez vc deve fazer um SQL aqui tb para os parâmetros requeridos tal como a lista.
            DBGrid(1).Visible = False
            SSPanel(0).Visible = True
            HScroll1.Visible = True
            HScroll1.Max = Mapa_Termo.Width - SSPanel(0).Width
            Mapa_Termo.Cls 'Limpa o mapa
            
        '1ª geração descendente
            
            For passo = 1 To 14
                Select Case passo
                    Case "1", "2", "3", "4", "11", "12", "13", "14"
                        Mapa_Termo.PaintPicture ImageList1.ListImages(45).Picture, outro, PGerDesc
                        Mapa_Termo.PaintPicture ImageList1.ListImages(46).Picture, outro + 288, PGerDesc
        '1ª geração ascendente
                        If passo = 1 Or passo = 3 Or passo = 11 Or passo = 13 Then
                            Mapa_Termo.PaintPicture ImageList1.ListImages(9).Picture, outro + 1100, PGerAsc
                            If passo = 1 Or passo = 11 Then
                                HouM = 384 'Mulher
                                ponto1 = outro + 1100 + HouM
                                If passo = 11 Then ponto2a = ponto1
                            Else
                                HouM = 108 'Homem
                                ponto2 = outro + 1100 + HouM
                                If passo = 3 Then ponto1a = ponto2
                            End If
                            If ponto1 > 0 And ponto2 > 0 Then
                                ponto1 = 0: ponto2 = 0
                            End If
                        End If
        'Geração do Ego
                        Mapa_Termo.PaintPicture ImageList1.ListImages(9).Picture, outro, GerEgo
        
                    Case "5", "6", "7", "8", "9", "10"
                        Mapa_Termo.PaintPicture ImageList1.ListImages(9).Picture, outro, PGerDesc
        '2ª geração descendente
                        Mapa_Termo.PaintPicture ImageList1.ListImages(45).Picture, outro, SGerDesc
                        Mapa_Termo.PaintPicture ImageList1.ListImages(46).Picture, outro + 288, SGerDesc
        'Geração do Ego
                        If passo = 5 Or passo = 9 Then
                            Mapa_Termo.PaintPicture ImageList1.ListImages(9).Picture, outro + 1100, GerEgo
                            '2ª geração ascendente
                            Mapa_Termo.PaintPicture ImageList1.ListImages(9).Picture, outro + 1100, SGerAsc
                        ElseIf passo = 7 Then
                            'O Ego
                            If SSOption(12).Value = True Then
                                Mapa_Termo.Line (15070 + Dente, 1650)-(15070 + Dente, 1650 + 230), QBColor(15)
                                Mapa_Termo.PaintPicture ImageList1.ListImages(13).Picture, 14671 + Dente, 1840 'O Ego masculino
                                Mapa_Termo.Line (14780 + Dente, 1640)-(14780 + Dente, 1640 + 230), QBColor(0)
                                
                            Else
                                Mapa_Termo.Line (14780 + Dente, 1650)-(14780 + Dente, 1650 + 230), QBColor(15)
                                Mapa_Termo.PaintPicture ImageList1.ListImages(17).Picture, 14671 + Dente, 1840 'O Ego feminino
                                'Call Termos_No_Mapa(63, Int(GerEgo), 400, 14910, Int(Dente))
                                Mapa_Termo.Line (15070 + Dente, 1640)-(15070 + Dente, 1640 + 230), QBColor(0)
                            End If
                            '1ª geração ascendente
                            Mapa_Termo.PaintPicture ImageList1.ListImages(9).Picture, outro + 1100, PGerAsc
                            
                            HouM = 108 'Homem
                            ponto2 = outro + 1100 + HouM
                        End If
                End Select
                outro = outro + espaco
            Next
            Call Linhas_Geração 'Chama o procedimento que desenha as linhas das gerações.
        Case 18
        Case 19
            Mapa_Parente.Height = SSPanel(2).Height
            Mapa_Parente.Width = SSPanel(2).Width
            Mapa_Parente.Left = 0
            Mapa_Parente.Top = 0
            HScroll_Parente.Max = Mapa_Parente.Width - SSPanel(2).Width
            VScroll_Parente.Max = Mapa_Parente.Height - SSPanel(2).Height
    End Select
    

End Sub





Private Sub Desliga_Tab(Tab_Ativo As Integer, Tab_control As Control, True_or_False As Integer)
'Esta rotina liga ou desliga os tabs necessários do SStab indicado

    Dim Índice_do_Tab As Integer 'Contará os índices dos tabs.
    
    For Índice_do_Tab = 0 To Tab_control.Tabs - 1 'Conta os índices dos tabs do corrente SStab.
        'Se o valor de Índice_do_Tab é diferente a Tab_Ativo, então...  Obs: A razão deste teste é para que o tab ativo não seja desligado.
        If Tab_Ativo <> Índice_do_Tab Then Tab_control.TabEnabled(Índice_do_Tab) = True_or_False 'Desliga o Tab (Índice_do_Tab) do SStab indicado.
    Next Índice_do_Tab

End Sub

Private Sub SSOption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If SSOption(2).Value = False Then SSOption(1).Value = True
End Sub

Private Sub SSOption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If SSOption(2).Value = True Then SSOption(1).Value = True
End Sub

Private Sub SSTab_Click(Index As Integer, PreviousTab As Integer)
    
    'Casas/Pessoal
    If SSTab(1).Tab = 0 Then
        'Ego
        If SSTab(2).Tab = 0 Then
            If Combo(9).Text <> "" Then
                Combo_Click (9)
            Else
                Combo_Click (10)
            End If
        End If
        
        'Família Nuclear
        If SSTab(2).Tab = 1 And Banco = 1 Then
            Call DBcp_Casais_Reposition
        End If
     
        'Análise Geral
        If SSTab(2).Tab = 2 Then
            If EgoAtualizado = 1 Then
                EstouCarregando = 1
                    Call Enche_Combos_Busca(DBcp_Ego, Combo(16), "Casa_Mora")
                    Call Enche_Combos_Busca(DBcp_Apoio_Lugar, Combo(17), "Apoio")
                    Call Enche_Combos_Busca(DBcp_Ego, Combo(18), "Nome_Ind")
                    Call Enche_Combos_Busca(DBcp_Ego, Combo(19), "Nome_Nac")
                    Call Enche_Combos_Busca(DBcp_Apoio_Lugar, Combo(20), "Apoio")
                    Call Enche_Combos_Busca(DBcp_Ego, Combo(21), "Clã")
                EstouCarregando = 0
                EgoAtualizado = 0
            End If
        End If
    End If
'Parentesco - Geral
    'If SSTab(1).Tab = 1 Then
        'object.Add ColIndex
        'DBGrid(1).Columns.Add 3
        'DBGrid(1).Column(3).Visible = True
        'DBGrid(1).Refresh
    'End If

End Sub


Public Sub Liga_Desliga_Controle(Enabliavel As Integer, Lockiavel As Integer)
    'Liga ou desliga os controles da página para incluir os novos dados.
    SSCommand(4).Enabled = Enabliavel
    SSCommand(5).Enabled = Enabliavel
    MaskCaixa(3).Enabled = Enabliavel
    Combo(15).Enabled = Enabliavel
    DBCombo(10).Enabled = Enabliavel
    DBCombo(11).Enabled = Enabliavel
    DBCombo(12).Enabled = Enabliavel
    DBCombo(13).Enabled = Enabliavel
    Text(0).Locked = Lockiavel
    Text(4).Locked = Lockiavel
    Text(5).Locked = Lockiavel
End Sub



Public Sub Enche_Combos_Busca(Banco As Control, Combo As Control, Campo As String)
        achou = 0
        Banco.Refresh
        Do Until Banco.Recordset.EOF
            If Banco.Recordset(Campo) <> "" Then
                For contaitens = 0 To Combo.ListCount - 1
                    If CStr(Combo.List(contaitens)) = CStr(Banco.Recordset(Campo)) Then achou = 1
                Next contaitens
                 'Se não achou a corrente pessoal do Banco no Combo, então...
                If achou = 0 Then
                    Combo.AddItem Banco.Recordset(Campo)
                    If Campo = "Apoio" Then
                        Combo.ItemData(Combo.NewIndex) = Banco.Recordset("ID_Apoio")
                    End If
                End If
            End If
            Banco.Recordset.MoveNext
            achou = 0
        Loop
End Sub

Private Sub SSTab_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub SSTab_KeyPress(Index As Integer, KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 13, 14 'Confirmação de termos de parentesco
            'Se a tecla pressionada for diferente do Backspace(8), então...
            If KeyAscii <> 8 Then
                'Somente as teclas numéricas são aceitáveis aqui.
                If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
            End If
    End Select
End Sub


Private Sub Text_LostFocus(Index As Integer)
    Select Case Index
        Case 13 'Confirmação de perguntas ampliadas
            'Verifica se o text box está vazio ou se o valor é menor que 2. Para confirmação, só é aceitável mais que 1.
            If Text(13).Text = "" Or Val(Text(13).Text) < 2 Then Text(13).Text = "2"
            'Grava o novo valor no PAC.INI no diretório do Windows.
            a = WritePrivateProfileString("CONFIRMAÇÃO", "Pergunta Ampliada", Text(13).Text, "pac.ini")
        
        Case 14 'Confirmação de termos de parentesco
            'Verifica se o text box está vazio ou se o valor é menor que 2. Para confirmação, só é aceitável mais que 1.
            If Text(14).Text = "" Or Val(Text(14).Text) < 2 Then Text(14).Text = "2"
            'Grava o novo valor no PAC.INI no diretório do Windows.
            a = WritePrivateProfileString("CONFIRMAÇÃO", "Parentesco", Text(14).Text, "pac.ini")
        Case 15 'Tribo
            a = WritePrivateProfileString("INFORMAÇÃO", "Tribo", Text(15).Text, "pac.ini")
        Case 16 'Aldeia
            a = WritePrivateProfileString("INFORMAÇÃO", "Aldeia", Text(16).Text, "pac.ini")
        Case 17 'Localização
            a = WritePrivateProfileString("INFORMAÇÃO", "Localização", Text(17).Text, "pac.ini")
    End Select

End Sub

Private Sub TreeView2_NodeClick(ByVal Node As ComctlLib.Node)
    Dim QualImagem As Integer
    Dim Índice As Integer
    Índice = Val(Node.Key)
    PAC1.DBcp_Casais.Recordset.FindFirst "ID_Casal =" & Índice
    'AniButton.Value = Switch(PAC1.DBcp_Casais.Recordset("civil") = 6, 1, PAC1.DBcp_Casais.Recordset("civil") = 3, 2, PAC1.DBcp_Casais.Recordset("civil") = 5, 3)
    Call DBCombo_Click(13, 2)
    'QualImagem = Switch(Node.Image = "B1a", 1, Node.Image = "B1b", 2, Node.Image = "B1c", 3, Node.Image = "B1d", 4, _
                        Node.Image = "C1a", 5, Node.Image = "C1b", 6, Node.Image = "C1c", 7, Node.Image = "C1d", 8, _
                        Node.Image = "D1a", 9, Node.Image = "D1b", 10, Node.Image = "D1c", 11, Node.Image = "D1d", 12)
End Sub


Public Function Ajustar_Pessoa(clicou As Integer, BancoDados As Control, Qualcombo As Control, Campo As String, ÍndiceAnterior As Integer, CaixaPicture As Control) As Integer
    If clicou = 1 Or PAC1.DBcp_Casais.Recordset.EditMode = dbEditNone Then
        BancoDados.Recordset.FindFirst "ID_Ego=" & ÍndiceAnterior
        If BancoDados.Recordset.NoMatch = False Then
            Qualcombo.Text = BancoDados.Recordset(Campo)
            Ajustar_Pessoa = BancoDados.Recordset("ID_Ego")
            'Mostra o ícone apropriado para o ego selecionado.
            Call Qual_Icon(BancoDados, CaixaPicture)
        Else
            Qualcombo.Text = ""
            Ajustar_Pessoa = ÍndiceAnterior
            'Não mostra nada, poís nenhum ego está selecionado.
            CaixaPicture.Picture = ImageList2.ListImages(16).Picture
        End If
    Else
        Qualcombo.Text = ""
        'Não mostra nada, poís nenhum ego está selecionado.
        CaixaPicture.Picture = ImageList2.ListImages(16).Picture
    End If
End Function

Public Sub Familias()
    Dim Chave As String
    Dim nodX As Node    ' Declare Node variable.
    'Dim Índice As String
    'contando = 1
    PAC1.TreeView2.Nodes.Clear
    EstouCarregando = 1
    PAC1.DBcp_Casais.Refresh
    PAC1.DBcp_Ego_Nomes.RecordSource = "select * from EGO"

    PAC1.DBcp_Ego_Nomes.Refresh 'PAC1.DBcp_Casais.Recordset.MoveFirst
    Do Until PAC1.DBcp_Casais.Recordset.EOF = True
        'PAC1.DBcp_Ego_Nomes.Refresh
        'PAC1.DBcp_Casais.Recordset.FindFirst "ID_Casal=" & PAC1.DBcp_Ego.Recordset("ID_Pais")
        PAC1.DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & PAC1.DBcp_Casais.Recordset("ID_Conj1")
        b = PAC1.DBcp_Ego_Nomes.Recordset.NoMatch
        'a = PAC1.DBcp_Ego_Nomes.Recordset.RecordCount
        Debug.Print PAC1.DBcp_Ego_Nomes.Recordset("ID_Ego")
        QualNomePai = IIf(PAC1.DBcp_Ego_Nomes.Recordset("Nome_Preferido") = 1, PAC1.DBcp_Ego_Nomes.Recordset("Nome_Ind"), PAC1.DBcp_Ego_Nomes.Recordset("Nome_Nac"))
        
        If PAC1.DBcp_Ego_Nomes.Recordset("Sexo") = 0 Then
            Chave = ImagemCasal("Orientação")
        End If
        
        
        PAC1.DBcp_Ego_Nomes.Recordset.FindFirst "ID_Ego=" & PAC1.DBcp_Casais.Recordset("ID_Conj2")
        QualNomeMãe = IIf(PAC1.DBcp_Ego_Nomes.Recordset("Nome_Preferido") = 1, PAC1.DBcp_Ego_Nomes.Recordset("Nome_Ind"), PAC1.DBcp_Ego_Nomes.Recordset("Nome_Nac"))
        
        If PAC1.DBcp_Ego_Nomes.Recordset("Sexo") = 1 Then
            Chave = ImagemCasal("Orientação")
        End If
       
        
        ' Add Node objects.
        ' First node with 'Root' as text.
        Set nodX = PAC1.TreeView2.Nodes.Add(, , Str(PAC1.DBcp_Casais.Recordset("ID_Casal")) & "Índice", QualNomePai & "  &  " & QualNomeMãe, Chave)
        'contando = contando + 1
        PAC1.DBcp_Casais.Recordset.MoveNext
    Loop
    EstouCarregando = 0
    'PAC1.DBcp_Casais.Recordset.MoveFirst
End Sub

Public Sub Qual_Icon(BancoDados As Control, CaixaPicture As Control)
    Dim Figura As Integer
    'Seleciona pelo sexo do ego.
    Select Case BancoDados.Recordset("sexo")
        Case 0 'Caso o sexo seja masculino(0)
            'Se o campo data_falec é nulo(vazio), isto indica que o ego está vivo,
            'então será mostrada a figura número 10. Caso contrário o ego está morto e
            'será mostrado a figura número 9.
            Figura = IIf(IsNull(BancoDados.Recordset("data_falec")) = True, 10, 9)
        Case 1 'Caso o sexo seja feminino(1)
            'Se o campo data_falec é nulo(vazio), isto indica que o ego está vivo,
            'então será mostrada a figura número 12. Caso contrário o ego está morto e
            'será mostrado a figura número 11.
            Figura = IIf(IsNull(BancoDados.Recordset("data_falec")) = True, 12, 11)
    End Select
    'O picturebox carrega a figura que está armazenada no imagelist2. Isto é feito através
    'da variável Figura que é um número inteiro.
    CaixaPicture.Picture = ImageList2.ListImages(Figura).Picture

End Sub

Public Function Dados_Válidos()
    Dim Resultado As Integer 'Receberá o valor do teste de campos válidos ou não para gravação.
    'Testa se os campos (Data, Data de Nascimento, Data de Falecimento, Pesquisador, Casa N°, Lugar que Mora,
    'Nome Indígena, Nome Nacional e Sexo) são válidos. Caso algum não seja válido,
    'o campo errado recebe o foco e o processo de gravação é interrompido.
    
    Dados_Válidos = True
    
'Data
    'Se o campo de data estiver vazio vai aparecer uma mensagem de alerta.
    If PAC1.MaskCaixa(0).ClipText = "" Then
        Mensagem_Erro = MsgBox(LoadResString(280 + Lingua) & Chr(32) & Chr(34) & UCase(Label(1).Caption) & Chr(34) & Chr(32) & LoadResString(281 + Lingua), 48, LoadResString(281 + Lingua)) 'Atenção!
        PAC1.MaskCaixa(0).SetFocus
        Dados_Válidos = False
        Exit Function
    'Verifica se o dado é uma data válida.
    ElseIf IsDate(PAC1.MaskCaixa(0)) = False Then
        PAC1.MaskCaixa(0).SetFocus
        Dados_Válidos = False
        Exit Function
    End If

'Data de nascimento
    If PAC1.MaskCaixa(1).ClipText <> "" Then
        If IsDate(PAC1.MaskCaixa(1)) = False Then
            PAC1.MaskCaixa(1).SetFocus
            Dados_Válidos = False
            Exit Function
        End If
    End If
    
'Data de falecimento
    If PAC1.MaskCaixa(2).ClipText <> "" Then
        If IsDate(PAC1.MaskCaixa(2)) = False Then
            PAC1.MaskCaixa(2).SetFocus
            Dados_Válidos = False
            Exit Function
        End If
    End If

'Pesquisador
    'A função Testa_Ego é invocada para verificar se o campo foi preenchido e devolve
    'o valor True(se foi preenchido) ou o valor False(se não foi preenchido).
    Resultado = Testa_Ego(PAC1.DBCombo(2), Chr(32) & Chr(34) & UCase(Label(3).Caption) & Chr(34) & Chr(32)) 'Pesquisador
    'Se o resultado for False então o controle em questão recebe o foco e o processo vai para a posição "pula:".
    If Resultado = False Then PAC1.DBCombo(2).SetFocus: Dados_Válidos = Resultado: Exit Function

'Casa N°
    'Resultado = Testa_Ego(PAC1.Combo(11), Chr(32) & Chr(34) & UCase(Label(6).Caption) & Chr(34) & Chr(32)) 'Casa n°
    'If Resultado = False Then PAC1.Combo(11).SetFocus: Dados_Válidos = Resultado: Exit Function

'Lugar que mora
    Resultado = Testa_Ego(PAC1.Combo(13), Chr(32) & Chr(34) & UCase(Label(7).Caption) & Chr(34) & Chr(32)) 'Lugar que mora
    If Resultado = False Then Combo(13).SetFocus: Dados_Válidos = Resultado: Exit Function

'Nome
    'Se o combo "Nome Indígena" está vazio, então...
    If PAC1.Combo(9).Text = "" Then
        'Testa se o combo "Nome Nacional" está preenchido, pois um destes dois combo precisa, ou os dois, precisam estar preenchido para que o registro seja gravado.
        Resultado = Testa_Ego(PAC1.Combo(10), Chr(32) & Chr(34) & UCase(Label(8).Caption) & Chr(34) & Chr(32) & LoadResString(285 + Lingua) & Chr(32) & Chr(34) & UCase(Label(9).Caption) & Chr(34) & Chr(32)) 'Nome indígena ou o campo Nome Nacional
        
        If Resultado = False Then 'Se o resultado for false, então...
            Combo(9).SetFocus 'Põe o foco no combo "Nome Indígena".
            Dados_Válidos = Resultado: Exit Function 'O processo vai para a posição "Pula:", interrompendo a gravação.
        End If
    
    End If

'Sexo
    If SexoEgo = 1000 Then 'Se SexoEgo=1000(Isto significa que nenhum sexo foi selecionado. Pode ser qualquer valor diferente de 0 ou 1), então...
        'Menssagem avisando que deve ser selecionado o sexo do Ego.
        Menssagem_Erro = MsgBox(LoadResString(284 + Lingua), 48, LoadResString(282 + Lingua)) 'Você tem que indicar o SEXO do Ego.
        SSOption(1).SetFocus 'Põe a opção do sexo masculino em foco.
        Resultado = False 'Põe o valor do resultado em False, pois o campo está inválido para a gravação.
        Dados_Válidos = Resultado
    End If

End Function


Public Sub Lista_Termos()
    'Este processo Preenche a lista de termos de parentesco baseado no tipo de termo
    'e no sexo do ego.
    'DBGrid(1).Columns(0).Caption = "teste"
    Dim Qual_Tipo As Integer 'Variável usada para o tipo de termo.
    Dim Qual_Sexo As Integer 'Variável usada para o sexo do ego.
    Qual_Tipo = Combo(2).ListIndex '0=Referência   1=Tratamento
    If SSOption(12).Value = True Then 'Se a escolha masculino for selecionada...
        Qual_Sexo = 0 'Masculino
    ElseIf SSOption(13).Value = True Then 'Se a escolha feminino for selecionada...
        Qual_Sexo = 1 'Feminino
    End If
    'O SQL abaixo executa o filtro usando as variáveis.
    DBpa_Termos.RecordSource = "SELECT Consulta_Geral.ID_Termo,Consulta_Geral.ID_Termo_Tec,Consulta_Geral.Termo_Tec,Consulta_Geral.Termo_Tec_IN, Consulta_Geral.Termo_Ind ,Consulta_Geral.Trilha, Consulta_Geral.ID_Tipo, Consulta_Geral.Sexo_Ego, Consulta_Geral.Obs From Consulta_Geral WHERE (((Consulta_Geral.ID_Tipo)=" & Qual_Tipo & ")) and (((Consulta_Geral.sexo_ego)=" & Qual_Sexo & "));"
    DBpa_Termos.Refresh 'O data control é reiniciado com a nova seleção.
    DBpa_Termos.Recordset.MoveLast 'Vai até o último registro para pode saber quantos registros existem.
    Label(33).Caption = DBpa_Termos.Recordset.RecordCount 'Mostra o número total de registros no Label(33). O número de registros coincide com o número de rows no dbgrid.
    DBCombo(0).ListField = IIf(Lingua = 0, "Termo_Tec", "Termo_Tec_IN")
    DBGrid(1).Columns(0).DataField = IIf(Lingua = 0, "Termo_Tec", "Termo_Tec_IN")
    DBpa_Termos.Recordset.MoveFirst 'Volta para o 1º registro no data control, o que força ir para o 1º row no dbgrid.
End Sub

Private Sub Linhas_Geração()
    'Mapa_Termo.Line (ponto1 + Dente, PGerAsc - 200)-(ponto2 + Dente, PGerAsc - 200), QBColor(0)
    Dim Termo As String
    PGerDesc = 2440
    SGerDesc = 3240
    GerEgo = 1640
    Dente = 800 'Controla o espaço para a margem esquerda do picture box
    PGerAsc = 840
    SGerAsc = 40
    
'Verifica qual é o tipo de tratamento e o sexo do ego para selecionar corretamente
'o Índice do Termo de parentesco disponível no DBpa_Termo, pois este acabou de ser
'filtrado com novos parâmetros.
    If Combo(2).ListIndex = 1 Then 'And SSOption(12).Value = True Then
        Indice_tipo = 153
    Else
        Indice_tipo = 0
    End If
    If SSOption(12).Value = True Then
        Indice_sexo = 76
    Else
        Indice_sexo = 0
    End If
    Indice_filtro = Indice_sexo + Indice_tipo
    
'Segunda Geração Ascendente
    'Os Termos
    Call Termos_No_Mapa(14 + Indice_filtro, Int(SGerAsc), 100, 10630, Int(Dente)) 'pai do pai
    Call Termos_No_Mapa(13 + Indice_filtro, Int(SGerAsc), 400, 10630, Int(Dente)) 'mãe do pai
    Call Termos_No_Mapa(12 + Indice_filtro, Int(SGerAsc), 100, 19200, Int(Dente)) 'pai da mãe
    Call Termos_No_Mapa(11 + Indice_filtro, Int(SGerAsc), 400, 19200, Int(Dente)) 'mãe da mãe
    
'Primeira Geração Ascendente
    'Os Termos
    Call Termos_No_Mapa(10 + Indice_filtro, Int(PGerAsc), 100, 2050, Int(Dente)) 'esposo da irmã do pai
    Call Termos_No_Mapa(9 + Indice_filtro, Int(PGerAsc), 400, 2050, Int(Dente)) 'irmã do pai
    Call Termos_No_Mapa(7 + Indice_filtro, Int(PGerAsc), 100, 6340, Int(Dente)) 'irmão do pai
    Call Termos_No_Mapa(8 + Indice_filtro, Int(PGerAsc), 400, 6340, Int(Dente)) 'esposa do irmão do pai
    Call Termos_No_Mapa(6 + Indice_filtro, Int(PGerAsc), 100, 14920, Int(Dente)) 'pai
    Call Termos_No_Mapa(1 + Indice_filtro, Int(PGerAsc), 400, 14920, Int(Dente)) 'mãe
    Call Termos_No_Mapa(3 + Indice_filtro, Int(PGerAsc), 100, 23480, Int(Dente)) 'esposo da irmã da mãe
    Call Termos_No_Mapa(2 + Indice_filtro, Int(PGerAsc), 400, 23480, Int(Dente)) 'irmã da mãe
    Call Termos_No_Mapa(4 + Indice_filtro, Int(PGerAsc), 100, 27770, Int(Dente)) 'irmão da mãe
    Call Termos_No_Mapa(5 + Indice_filtro, Int(PGerAsc), 400, 27770, Int(Dente)) 'esposa do irmão da mãe
    
    'As Linhas
    Mapa_Termo.Line (2210 + Dente, PGerAsc)-(14770 + Dente, PGerAsc), QBColor(0)
    Mapa_Termo.Line (2210 + Dente, PGerAsc)-(2210 + Dente, PGerAsc + 250), QBColor(0)   'Irmã do pai
    Mapa_Termo.Line (6200 + Dente, PGerAsc)-(6200 + Dente, PGerAsc + 250), QBColor(0)  'Irmão do pai
    Mapa_Termo.Line (10630 + Dente, PGerAsc)-(10630 + Dente, PGerAsc - 500), QBColor(0)  'Os pais do pai
    Mapa_Termo.Line (14770 + Dente, PGerAsc)-(14770 + Dente, PGerAsc + 250), QBColor(0)  'Pai
    
    Mapa_Termo.Line (15070 + Dente, PGerAsc)-(27620 + Dente, PGerAsc), QBColor(0)
    Mapa_Termo.Line (15070 + Dente, PGerAsc)-(15070 + Dente, PGerAsc + 250), QBColor(0) 'Mãe
    Mapa_Termo.Line (19200 + Dente, PGerAsc)-(19200 + Dente, PGerAsc - 500), QBColor(0) 'Os pais da mãe
    Mapa_Termo.Line (23640 + Dente, PGerAsc)-(23640 + Dente, PGerAsc + 250), QBColor(0) 'Irmã da mãe
    Mapa_Termo.Line (27620 + Dente, PGerAsc)-(27620 + Dente, PGerAsc + 250), QBColor(0) 'Irmão da mãe

'Geração do Ego
    'Os Termos
    If SSOption(12).Value = True Then
        Call Termos_No_Mapa(64 + Indice_filtro, Int(GerEgo), 400, 14910, Int(Dente))
    Else
        Call Termos_No_Mapa(63 + Indice_filtro, Int(GerEgo), 100, 14910, Int(Dente))
    End If
    Call Termos_No_Mapa(31 + Indice_filtro, Int(GerEgo), 100, 965, Int(Dente))
    Call Termos_No_Mapa(32 + Indice_filtro, Int(GerEgo), 400, 965, Int(Dente)) 'Esposa do filho da irmã do pai
    Call Termos_No_Mapa(33 + Indice_filtro, Int(GerEgo), 400, 3095, Int(Dente))
    Call Termos_No_Mapa(34 + Indice_filtro, Int(GerEgo), 100, 3095, Int(Dente)) 'Esposo da filha da irmã do pai
    Call Termos_No_Mapa(27 + Indice_filtro, Int(GerEgo), 100, 5245, Int(Dente))
    Call Termos_No_Mapa(28 + Indice_filtro, Int(GerEgo), 400, 5245, Int(Dente))
    Call Termos_No_Mapa(29 + Indice_filtro, Int(GerEgo), 400, 7385, Int(Dente))
    Call Termos_No_Mapa(30 + Indice_filtro, Int(GerEgo), 100, 7385, Int(Dente))
    Call Termos_No_Mapa(15 + Indice_filtro, Int(GerEgo), 100, 10630, Int(Dente))
    Call Termos_No_Mapa(16 + Indice_filtro, Int(GerEgo), 400, 10630, Int(Dente))
    Call Termos_No_Mapa(17 + Indice_filtro, Int(GerEgo), 400, 19200, Int(Dente))
    Call Termos_No_Mapa(18 + Indice_filtro, Int(GerEgo), 100, 19200, Int(Dente))
    Call Termos_No_Mapa(19 + Indice_filtro, Int(GerEgo), 100, 22380, Int(Dente))
    Call Termos_No_Mapa(20 + Indice_filtro, Int(GerEgo), 400, 22380, Int(Dente))
    Call Termos_No_Mapa(21 + Indice_filtro, Int(GerEgo), 400, 24530, Int(Dente))
    Call Termos_No_Mapa(22 + Indice_filtro, Int(GerEgo), 100, 24530, Int(Dente))
    Call Termos_No_Mapa(23 + Indice_filtro, Int(GerEgo), 100, 26665, Int(Dente))
    Call Termos_No_Mapa(24 + Indice_filtro, Int(GerEgo), 400, 26665, Int(Dente))
    Call Termos_No_Mapa(26 + Indice_filtro, Int(GerEgo), 100, 28810, Int(Dente))
    Call Termos_No_Mapa(25 + Indice_filtro, Int(GerEgo), 400, 28810, Int(Dente))
    
    'As Linhas
    Mapa_Termo.Line (830 + Dente, GerEgo)-(3240 + Dente, GerEgo), QBColor(0)
    Mapa_Termo.Line (830 + Dente, GerEgo)-(830 + Dente, GerEgo + 250), QBColor(0) 'Filho da irmã do pai
    
    Mapa_Termo.Line (3240 + Dente, GerEgo)-(3240 + Dente, GerEgo + 230), QBColor(0) 'Filha da irmã do pai
    Mapa_Termo.Line (2050 + Dente, GerEgo)-(2050 + Dente, GerEgo - 500), QBColor(0)
    
    Mapa_Termo.Line (5110 + Dente, GerEgo)-(7540 + Dente, GerEgo), QBColor(0)
    Mapa_Termo.Line (5110 + Dente, GerEgo)-(5110 + Dente, GerEgo + 250), QBColor(0) 'Filho do irmão do pai
    Mapa_Termo.Line (7540 + Dente, GerEgo)-(7540 + Dente, GerEgo + 230), QBColor(0) 'Filha do irmão do pai
    Mapa_Termo.Line (6340 + Dente, GerEgo)-(6340 + Dente, GerEgo - 500), QBColor(0)
    
    Mapa_Termo.Line (10500 + Dente, GerEgo)-(19360 + Dente, GerEgo), QBColor(0)
    Mapa_Termo.Line (10500 + Dente, GerEgo)-(10500 + Dente, GerEgo + 250), QBColor(0) 'Irmão
    Mapa_Termo.Line (19360 + Dente, GerEgo)-(19360 + Dente, GerEgo + 230), QBColor(0) 'Irmã
    Mapa_Termo.Line (14920 + Dente, GerEgo)-(14920 + Dente, GerEgo - 500), QBColor(0)
    
    Mapa_Termo.Line (22250 + Dente, GerEgo)-(24670 + Dente, GerEgo), QBColor(0)
    Mapa_Termo.Line (22250 + Dente, GerEgo)-(22250 + Dente, GerEgo + 250), QBColor(0) 'Filho da irmã da mãe
    Mapa_Termo.Line (24670 + Dente, GerEgo)-(24670 + Dente, GerEgo + 230), QBColor(0) 'Filha da irmã da mãe
    Mapa_Termo.Line (23480 + Dente, GerEgo)-(23480 + Dente, GerEgo - 500), QBColor(0)
    
    Mapa_Termo.Line (26530 + Dente, GerEgo)-(28950 + Dente, GerEgo), QBColor(0)
    Mapa_Termo.Line (26530 + Dente, GerEgo)-(26530 + Dente, GerEgo + 250), QBColor(0) 'Filho do irmão da mãe
    Mapa_Termo.Line (28950 + Dente, GerEgo)-(28950 + Dente, GerEgo + 230), QBColor(0) 'Filha do irmão da mãe
    Mapa_Termo.Line (27770 + Dente, GerEgo)-(27770 + Dente, GerEgo - 500), QBColor(0)

'Primeira Geração Descendente
    'Os Termos
    Call Termos_No_Mapa(55 + Indice_filtro, Int(PGerDesc), 100, 965, Int(Dente))
    Call Termos_No_Mapa(56 + Indice_filtro, Int(PGerDesc), 400, 965, Int(Dente))
    Call Termos_No_Mapa(57 + Indice_filtro, Int(PGerDesc), 100, 3095, Int(Dente))
    Call Termos_No_Mapa(58 + Indice_filtro, Int(PGerDesc), 400, 3095, Int(Dente))
    Call Termos_No_Mapa(59 + Indice_filtro, Int(PGerDesc), 100, 5245, Int(Dente))
    Call Termos_No_Mapa(60 + Indice_filtro, Int(PGerDesc), 400, 5245, Int(Dente))
    Call Termos_No_Mapa(61 + Indice_filtro, Int(PGerDesc), 100, 7385, Int(Dente))
    Call Termos_No_Mapa(62 + Indice_filtro, Int(PGerDesc), 400, 7385, Int(Dente))
    Call Termos_No_Mapa(39 + Indice_filtro, Int(PGerDesc), 100, 9530, Int(Dente))
    Call Termos_No_Mapa(40 + Indice_filtro, Int(PGerDesc), 400, 9530, Int(Dente))
    Call Termos_No_Mapa(42 + Indice_filtro, Int(PGerDesc), 100, 11660, Int(Dente))
    Call Termos_No_Mapa(41 + Indice_filtro, Int(PGerDesc), 400, 11660, Int(Dente))
    Call Termos_No_Mapa(35 + Indice_filtro, Int(PGerDesc), 100, 13810, Int(Dente))
    Call Termos_No_Mapa(36 + Indice_filtro, Int(PGerDesc), 400, 13810, Int(Dente))
    Call Termos_No_Mapa(38 + Indice_filtro, Int(PGerDesc), 100, 15960, Int(Dente))
    Call Termos_No_Mapa(37 + Indice_filtro, Int(PGerDesc), 400, 15960, Int(Dente))
    Call Termos_No_Mapa(43 + Indice_filtro, Int(PGerDesc), 100, 18100, Int(Dente))
    Call Termos_No_Mapa(44 + Indice_filtro, Int(PGerDesc), 400, 18100, Int(Dente)) 'esposa do filho da irmã
    Call Termos_No_Mapa(46 + Indice_filtro, Int(PGerDesc), 100, 20240, Int(Dente))
    Call Termos_No_Mapa(45 + Indice_filtro, Int(PGerDesc), 400, 20240, Int(Dente)) 'filha da irmã
    Call Termos_No_Mapa(47 + Indice_filtro, Int(PGerDesc), 100, 22380, Int(Dente))
    Call Termos_No_Mapa(48 + Indice_filtro, Int(PGerDesc), 400, 22380, Int(Dente))
    Call Termos_No_Mapa(49 + Indice_filtro, Int(PGerDesc), 100, 24530, Int(Dente))
    Call Termos_No_Mapa(50 + Indice_filtro, Int(PGerDesc), 400, 24530, Int(Dente)) 'filha da filha da irmã da mãe
    Call Termos_No_Mapa(51 + Indice_filtro, Int(PGerDesc), 100, 26665, Int(Dente))
    Call Termos_No_Mapa(52 + Indice_filtro, Int(PGerDesc), 400, 26665, Int(Dente))
    Call Termos_No_Mapa(53 + Indice_filtro, Int(PGerDesc), 100, 28810, Int(Dente))
    Call Termos_No_Mapa(54 + Indice_filtro, Int(PGerDesc), 400, 28810, Int(Dente)) 'filha da filha do irmão da mãe
    
    'As Linhas
    Mapa_Termo.Line (830 + Dente, PGerDesc)-(1100 + Dente, PGerDesc), QBColor(0)
    Mapa_Termo.Line (830 + Dente, PGerDesc)-(830 + Dente, PGerDesc + 250), QBColor(0) 'Filho do filho da irmã do pai
    Mapa_Termo.Line (1100 + Dente, PGerDesc)-(1100 + Dente, PGerDesc + 230), QBColor(0) 'Filha do filho da irmã do pai
    Mapa_Termo.Line (965 + Dente, PGerDesc)-(965 + Dente, PGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (2950 + Dente, PGerDesc)-(3240 + Dente, PGerDesc), QBColor(0)
    Mapa_Termo.Line (2950 + Dente, PGerDesc)-(2950 + Dente, PGerDesc + 250), QBColor(0) 'Filho da filha da irmã do pai
    Mapa_Termo.Line (3240 + Dente, PGerDesc)-(3240 + Dente, PGerDesc + 230), QBColor(0) 'Filha da filha da irmã do pai
    Mapa_Termo.Line (3095 + Dente, PGerDesc)-(3095 + Dente, PGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (5110 + Dente, PGerDesc)-(5380 + Dente, PGerDesc), QBColor(0)
    Mapa_Termo.Line (5110 + Dente, PGerDesc)-(5110 + Dente, PGerDesc + 250), QBColor(0) 'Filho do filho do irmão do pai
    Mapa_Termo.Line (5380 + Dente, PGerDesc)-(5380 + Dente, PGerDesc + 230), QBColor(0) 'Filha do filho do irmão do pai
    Mapa_Termo.Line (5245 + Dente, PGerDesc)-(5245 + Dente, PGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (7250 + Dente, PGerDesc)-(7530 + Dente, PGerDesc), QBColor(0)
    Mapa_Termo.Line (7250 + Dente, PGerDesc)-(7250 + Dente, PGerDesc + 250), QBColor(0) 'Filho da filha do irmão do pai
    Mapa_Termo.Line (7530 + Dente, PGerDesc)-(7530 + Dente, PGerDesc + 230), QBColor(0) 'Filha da filha do irmão do pai
    Mapa_Termo.Line (7385 + Dente, PGerDesc)-(7385 + Dente, PGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (9400 + Dente, PGerDesc)-(11810 + Dente, PGerDesc), QBColor(0)
    Mapa_Termo.Line (9400 + Dente, PGerDesc)-(9400 + Dente, PGerDesc + 250), QBColor(0) 'Filho do irmão
    Mapa_Termo.Line (11810 + Dente, PGerDesc)-(11810 + Dente, PGerDesc + 230), QBColor(0) 'Filha do irmão
    Mapa_Termo.Line (10630 + Dente, PGerDesc)-(10630 + Dente, PGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (13670 + Dente, PGerDesc)-(16120 + Dente, PGerDesc), QBColor(0)
    Mapa_Termo.Line (13670 + Dente, PGerDesc)-(13670 + Dente, PGerDesc + 250), QBColor(0) 'Filho
    Mapa_Termo.Line (16120 + Dente, PGerDesc)-(16120 + Dente, PGerDesc + 230), QBColor(0) 'Filha
    Mapa_Termo.Line (14910 + Dente, PGerDesc)-(14910 + Dente, PGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (17960 + Dente, PGerDesc)-(20400 + Dente, PGerDesc), QBColor(0)
    Mapa_Termo.Line (17960 + Dente, PGerDesc)-(17960 + Dente, PGerDesc + 250), QBColor(0) 'Filho da irmã
    Mapa_Termo.Line (20400 + Dente, PGerDesc)-(20400 + Dente, PGerDesc + 230), QBColor(0) 'Filha da irmã
    Mapa_Termo.Line (19200 + Dente, PGerDesc)-(19200 + Dente, PGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (22240 + Dente, PGerDesc)-(22510 + Dente, PGerDesc), QBColor(0)
    Mapa_Termo.Line (22240 + Dente, PGerDesc)-(22240 + Dente, PGerDesc + 250), QBColor(0) 'Filho do filho da irmã da mãe
    Mapa_Termo.Line (22510 + Dente, PGerDesc)-(22510 + Dente, PGerDesc + 230), QBColor(0) 'Filha do filho da irmã da mãe
    Mapa_Termo.Line (22380 + Dente, PGerDesc)-(22380 + Dente, PGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (24400 + Dente, PGerDesc)-(24660 + Dente, PGerDesc), QBColor(0)
    Mapa_Termo.Line (24400 + Dente, PGerDesc)-(24400 + Dente, PGerDesc + 250), QBColor(0) 'Filho da filha da irmã da mãe
    Mapa_Termo.Line (24660 + Dente, PGerDesc)-(24660 + Dente, PGerDesc + 230), QBColor(0) 'Filha da filha da irmã da mãe
    Mapa_Termo.Line (24530 + Dente, PGerDesc)-(24530 + Dente, PGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (26520 + Dente, PGerDesc)-(26810 + Dente, PGerDesc), QBColor(0)
    Mapa_Termo.Line (26520 + Dente, PGerDesc)-(26520 + Dente, PGerDesc + 250), QBColor(0) 'Filho do filho do irmão da mãe
    Mapa_Termo.Line (26810 + Dente, PGerDesc)-(26810 + Dente, PGerDesc + 230), QBColor(0) 'Filha do filho do irmão da mãe
    Mapa_Termo.Line (26665 + Dente, PGerDesc)-(26665 + Dente, PGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (28680 + Dente, PGerDesc)-(28950 + Dente, PGerDesc), QBColor(0)
    Mapa_Termo.Line (28680 + Dente, PGerDesc)-(28680 + Dente, PGerDesc + 250), QBColor(0) 'Filho da filha do irmão da mãe
    Mapa_Termo.Line (28950 + Dente, PGerDesc)-(28950 + Dente, PGerDesc + 230), QBColor(0) 'Filha da filha do irmão da mãe
    Mapa_Termo.Line (28810 + Dente, PGerDesc)-(28810 + Dente, PGerDesc - 500), QBColor(0)

'Segunda Geração Descendente
    'Os Termos
    Call Termos_No_Mapa(65 + Indice_filtro, Int(SGerDesc), 100, 9530, Int(Dente)) 'filho do filho do irmão
    Call Termos_No_Mapa(66 + Indice_filtro, Int(SGerDesc), 400, 9530, Int(Dente)) 'filha do filho do irmão
    Call Termos_No_Mapa(67 + Indice_filtro, Int(SGerDesc), 100, 11660, Int(Dente))
    Call Termos_No_Mapa(68 + Indice_filtro, Int(SGerDesc), 400, 11660, Int(Dente))
    Call Termos_No_Mapa(69 + Indice_filtro, Int(SGerDesc), 100, 13810, Int(Dente))
    Call Termos_No_Mapa(70 + Indice_filtro, Int(SGerDesc), 400, 13810, Int(Dente))
    Call Termos_No_Mapa(71 + Indice_filtro, Int(SGerDesc), 100, 15960, Int(Dente)) 'filho da filha
    Call Termos_No_Mapa(72 + Indice_filtro, Int(SGerDesc), 400, 15960, Int(Dente)) 'filha da filha
    Call Termos_No_Mapa(73 + Indice_filtro, Int(SGerDesc), 100, 18100, Int(Dente))
    Call Termos_No_Mapa(74 + Indice_filtro, Int(SGerDesc), 400, 18100, Int(Dente)) 'filha do filho da irmã
    Call Termos_No_Mapa(75 + Indice_filtro, Int(SGerDesc), 100, 20240, Int(Dente))
    Call Termos_No_Mapa(76 + Indice_filtro, Int(SGerDesc), 400, 20240, Int(Dente)) 'filha da filha da irmã
    
    'As Linhas
    Mapa_Termo.Line (9400 + Dente, SGerDesc)-(9660 + Dente, SGerDesc), QBColor(0)
    Mapa_Termo.Line (9400 + Dente, SGerDesc)-(9400 + Dente, SGerDesc + 250), QBColor(0) 'Filho do filho do irmão
    Mapa_Termo.Line (9660 + Dente, SGerDesc)-(9660 + Dente, SGerDesc + 230), QBColor(0) 'Filha do filho do irmão
    Mapa_Termo.Line (9530 + Dente, SGerDesc)-(9530 + Dente, SGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (11530 + Dente, SGerDesc)-(11810 + Dente, SGerDesc), QBColor(0)
    Mapa_Termo.Line (11530 + Dente, SGerDesc)-(11530 + Dente, SGerDesc + 250), QBColor(0) 'Filho da filha do irmão
    Mapa_Termo.Line (11810 + Dente, SGerDesc)-(11810 + Dente, SGerDesc + 230), QBColor(0) 'Filha da filha do irmão
    Mapa_Termo.Line (11660 + Dente, SGerDesc)-(11660 + Dente, SGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (13680 + Dente, SGerDesc)-(13960 + Dente, SGerDesc), QBColor(0)
    Mapa_Termo.Line (13680 + Dente, SGerDesc)-(13680 + Dente, SGerDesc + 250), QBColor(0) 'Filho do filho
    Mapa_Termo.Line (13960 + Dente, SGerDesc)-(13960 + Dente, SGerDesc + 230), QBColor(0) 'Filha do filho
    Mapa_Termo.Line (13810 + Dente, SGerDesc)-(13810 + Dente, SGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (15830 + Dente, SGerDesc)-(16100 + Dente, SGerDesc), QBColor(0)
    Mapa_Termo.Line (15830 + Dente, SGerDesc)-(15830 + Dente, SGerDesc + 250), QBColor(0) 'Filho da filha
    Mapa_Termo.Line (16100 + Dente, SGerDesc)-(16100 + Dente, SGerDesc + 230), QBColor(0) 'Filha da filha
    Mapa_Termo.Line (15960 + Dente, SGerDesc)-(15960 + Dente, SGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (17960 + Dente, SGerDesc)-(18230 + Dente, SGerDesc), QBColor(0)
    Mapa_Termo.Line (17960 + Dente, SGerDesc)-(17960 + Dente, SGerDesc + 250), QBColor(0) 'Filho do filho da irmã
    Mapa_Termo.Line (18230 + Dente, SGerDesc)-(18230 + Dente, SGerDesc + 230), QBColor(0) 'Filha do filho da irmã
    Mapa_Termo.Line (18100 + Dente, SGerDesc)-(18100 + Dente, SGerDesc - 500), QBColor(0)
    
    Mapa_Termo.Line (20110 + Dente, SGerDesc)-(20380 + Dente, SGerDesc), QBColor(0)
    Mapa_Termo.Line (20110 + Dente, SGerDesc)-(20110 + Dente, SGerDesc + 250), QBColor(0) 'Filho da filha da irmã
    Mapa_Termo.Line (20380 + Dente, SGerDesc)-(20380 + Dente, SGerDesc + 230), QBColor(0) 'Filha da filha da irmã
    Mapa_Termo.Line (20240 + Dente, SGerDesc)-(20240 + Dente, SGerDesc - 500), QBColor(0)
End Sub

Private Sub Termos_No_Mapa(ID_Termo As Integer, Geração As Integer, Distância As Integer, Linha As Integer, Dente As Integer)
    'ID_Termo é o índice do termo a ser pesquisado.
    'Geração é a distância da linha de geração com relação ao topo do picture box.
    'Distância é o espaço onde o termo será escrito abaixo da linha de geração. Se for M o texto é colocado mas para baixo que o do H.
    'Linha é a linha dos pais que será usada como referência para escrever o termo a esquerda se for H e a direita se for M.
    'Dente é o espaço esquerdo de todos os elementos no picture box.
    'Esta rotina procura os termos confirmados e coloca-os no desenho do mapa. Cada termo está
    'relacionado com um índice no mapa.
    Dim Termo As String
    DBpa_Termos.Recordset.FindFirst "ID_Termo =" & ID_Termo
    If DBpa_Termos.Recordset.NoMatch Then
        Termo = ""
    Else
        If IsNull(DBpa_Termos.Recordset("Termo_Ind")) = True Then
            Termo = ""
        Else
            Termo = DBpa_Termos.Recordset("Termo_Ind")
        End If
    End If
    Mapa_Termo.FontSize = 8
    If Distância = 100 Then
        Mapa_Termo.FontItalic = True
        Mapa_Termo.CurrentX = ((Linha + Dente) - 200) - Mapa_Termo.TextWidth(Termo)
        Mapa_Termo.ForeColor = QBColor(9) 'Azul
    Else
        Mapa_Termo.FontItalic = False
        Mapa_Termo.CurrentX = (Linha + Dente) + 100
        Mapa_Termo.ForeColor = QBColor(12) 'Vermelho
    End If
    Mapa_Termo.CurrentY = Geração + Distância '100 H ou 400 M
    Mapa_Termo.Print Termo
End Sub

Public Sub Termo_Relacionado(Sexo As Integer, Termo As String)
    Select Case Sexo
        Case 0 'Masculino
            Mapa_Parente.FontItalic = True
            Mapa_Parente.ForeColor = QBColor(9) 'Azul
            Mapa_Parente.CurrentX = Mapa_Parente.CurrentX - Mapa_Parente.TextWidth(Termo) - 80
            Mapa_Parente.Print Termo
        Case 1 'Feminino
            Mapa_Parente.FontItalic = False
            Mapa_Parente.ForeColor = QBColor(12) 'Azul
            Mapa_Parente.CurrentX = Mapa_Parente.CurrentX + 130
            Mapa_Parente.Print Termo
    End Select
End Sub

Private Sub VScroll_Parente_Change()
    Mapa_Parente.Top = VScroll_Parente.Value * -1
End Sub

Private Sub VScroll_Parente_Scroll()
    Mapa_Parente.Top = VScroll_Parente.Value * -1
End Sub


Public Sub Mapinha(Trilha As String, Nome As String)
        'Dim trilha As String
        Dim Passos As Integer
        Dim TB As String 'Termo Básico
        Dim LV As Integer 'Linha vertical usado para separar as gerações
        Dim LH As Integer 'Linha horizontal usado para separar os indivíduos
        Dim Civil As Integer 'Determina se o ego tem que aparecer casado ou não.
        Dim Lado As Integer 'Indica se é para o lado esquerdo(Espaço para esposo) ou direito(Espaço para esposa)
        Dim Espaço_V As Integer 'Indica o espaço vertical necessário para mostrar todo o gráfico.
        'Dim Espaço_H As Integer 'Indica o espaço horizontal necessário para mostrar todo o gráfico.
        Dim V As Integer 'Espaço vertical
        Dim V_Pais As Integer 'Espaço vertical das gerações de pais
        Dim V_Filhos As Integer 'Espaço vertical das gerações de filhos
        Dim H_irmãos As Integer 'Indica quantos espaços horizontais serão necessários
        Dim Parou As Integer 'Indicará se espaço vertical acima do ego parou de aumentar
        
        Espaço_V = 600 '600 é um espaço padrão para a distância vertical entre uma geração e outra
        LV = Espaço_V / 3 'Este valor é terça parte do valor padrão entre duas gerações
        LH = 1500 'Espaço da linha horizontal que liga os irmãos

        'Trilha = "4667474"  'DBCombo(0).BoundText
        
        For Passos = 1 To Len(Trilha) 'Conta o espaço vertical usado pelas gerações
            TB = Mid(Trilha, Passos, 1) 'Pega cada termo básico individualmente de dentro da string
            If TB = "1" Or TB = "2" Then V_Pais = V_Pais + 1 'Espaço vertical dos pais
            If TB = "5" Or TB = "6" Then V_Filhos = V_Filhos + 1 'Espaço vertical dos filhos
            If TB = "3" Or TB = "4" Then H_irmãos = H_irmãos + 1 'Espaços horizontais dos irmãos
        Next Passos
        'Mapa_Parente.Width = SSPanel(2).Width: Mapa_Parente.Height = SSPanel(2).Height 'Ajusta o tamanho do mapa_parente
        If H_irmãos > 2 Then Mapa_Parente.Width = (LH * H_irmãos) * 2 'Ajusta o espaço horizontal
        If V_Pais >= V_Filhos Then 'Se o espaço vertical dos pais é maior ou igual ao dos filhos...
            V = V_Pais 'O espaço vertical dos pais tem mais importancia
        Else 'Caso contrário...
            V = V_Filhos * -1 'O espaço v dos filhos tem mais importancia
            V = IIf(V_Pais > 0, V * -1, V)  'Se tem alguma geração de pais, o número v deve ser positivo
            Parou = 1 'indica que espaço vertical a cima do ego deve parar
        End If
        If V < 0 Then 'Trata-se dos filhos
            If V < -5 Then
                Espaço_V = (Espaço_V * 2) + (600 * (V * -1))
                Mapa_Parente.Height = Espaço_V + 600
            End If
            Espaço_V = 600
        Else 'Trata-se dos pais
            Espaço_V = (Espaço_V * 2) + (600 * V)
            If V >= 5 Then
                Mapa_Parente.Height = Espaço_V + 600
            End If
            If Parou = 1 Then 'O espaço vertical acima do ego está parado
                V = V - (V_Filhos - V_Pais) 'Para calcular o indice de quanto espaço vertical abaixo do ego deve existir.
                Espaço_V = (600 * 2) + (600 * V) 'O Valor 600 é a distância padrão entre duas gerações.
            End If
        End If
        LV = 200 'Este valor é terça parte do valor padrão entre duas gerações
        LH = 1500 'Espaço horizontal das linhas que unem os irmãos
        
        Lado = IIf(InStr(Trilha, "7") <> 0, -1, 1) 'Se no relacionamento escolhido aparece o termo Esposo, o lado será esquerdo _
                                                    para facilitar o espaço da colocação do nome do esposo no fim do termo.
        'DBpa_Termos.Recordset.Bookmark = DBCombo(0).SelectedItem
        Mapa_Parente.Cls 'Limpa o mapa_parentes
        
        Mapa_Parente.Top = 0: Mapa_Parente.Left = 0 'Ajusta a posição do mapa_patente
        TB = Mid(Trilha, 1, 1) 'Captura o indice do primeiro termo básico nesto termo técnico
        'Debug.Print ImageList1.ListImages(47).Picture.Height
        If SSOption(16).Value = True Then 'Aqui é escolhido qual ego vai estar em foco. Ego Masculino
            Civil = IIf(CInt(TB) > 4, 13, 47) 'Se TB for maior que 4, ou seja, TB é igual a Filho(a) ou Esposo(a) _
                                               Então o ego deve aparecer casado.
            'Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.Width / 2 - (ImageList1.ListImages(Civil).Picture.Width / 4), Mapa_Parente.Height / 2 + Espaço_V 'O Ego masculino
            Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.Width / 2 - (ImageList1.ListImages(Civil).Picture.Width / 4), Espaço_V 'O Ego masculino
        Else 'Ego Feminino
            Civil = IIf(CInt(TB) > 4, 17, 48)
            'Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.Width / 2 - (ImageList1.ListImages(Civil).Picture.Width / 4), Mapa_Parente.Height / 2 + Espaço_V 'O Ego masculino
            Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.Width / 2 - (ImageList1.ListImages(Civil).Picture.Width / 4), Espaço_V 'O Ego masculino
        End If
        If CInt(TB) > 4 Then 'Ego Casado. Ajusta onde vai começar a linha de relacionamento dependendo _
                              gráfico utilizado. O início é diferente entre o ego casado e solteiro
            'Mapa_Parente.CurrentY = Mapa_Parente.Height / 2 + (ImageList1.ListImages(Civil).Picture.Height / 2) + Espaço_V 'Mapa_Parente.CurrentY - ImageList1.ListImages(Civil).Picture.Height
            Mapa_Parente.CurrentY = (ImageList1.ListImages(Civil).Picture.Height / 2) + Espaço_V 'Mapa_Parente.CurrentY - ImageList1.ListImages(Civil).Picture.Height
            Mapa_Parente.CurrentX = Mapa_Parente.Width / 2 + 30
        Else 'Ego Solteiro
            'Mapa_Parente.CurrentY = Mapa_Parente.Height / 2 + Espaço_V ' Mapa_Parente.Height - 170 - ImageList1.ListImages(Civil).Picture.Height
            Mapa_Parente.CurrentY = Espaço_V ' Mapa_Parente.Height - 170 - ImageList1.ListImages(Civil).Picture.Height
            Mapa_Parente.CurrentX = Mapa_Parente.Width / 2
        End If
        For Passos = 1 To Len(Trilha) 'Inicia o passeio pela string para desenhar o mapa.
            TB = Mid(Trilha, Passos, 1)
            Select Case TB
                Case "1" 'Pai
                    Mapa_Parente.Line (Mapa_Parente.CurrentX, Mapa_Parente.CurrentY)-(Mapa_Parente.CurrentX, Mapa_Parente.CurrentY - LV * 3), QBColor(0) 'Desenha a linha que liga com os pais
                    Mapa_Parente.PaintPicture ImageList1.ListImages(9).Picture, Mapa_Parente.CurrentX - (ImageList1.ListImages(9).Picture.Width / 4 + 30), Mapa_Parente.CurrentY 'O Ego masculino
                    Mapa_Parente.CurrentX = Mapa_Parente.CurrentX - ImageList1.ListImages(9).Picture.Width / 4 + 60 'Posiciona em cima do pai
                    If Passos = Len(Trilha) Then 'Se este é o único termo básico no corrente termo técnico
                        Call Termo_Relacionado(0, "Pai=> " & Nome) 'Escreve o nome da pessoa focalizada
                    End If
                Case "2" 'Mãe
                    Mapa_Parente.Line (Mapa_Parente.CurrentX, Mapa_Parente.CurrentY)-(Mapa_Parente.CurrentX, Mapa_Parente.CurrentY - LV * 3), QBColor(0) 'Desenha a linha que liga com os pais
                    Mapa_Parente.PaintPicture ImageList1.ListImages(9).Picture, Mapa_Parente.CurrentX - (ImageList1.ListImages(9).Picture.Width / 4 + 30), Mapa_Parente.CurrentY 'O Ego masculino
                    Mapa_Parente.CurrentX = Mapa_Parente.CurrentX + ImageList1.ListImages(9).Picture.Width / 4 - 60 'Posiciona em cima da Mãe
                    If Passos = Len(Trilha) Then 'Se este é o único termo básico no corrente termo técnico
                        Call Termo_Relacionado(1, Nome & " <=Mãe") 'Escreve o nome da pessoa focalizada
                    End If

                Case "3" 'Irmão
                    Mapa_Parente.Line (Mapa_Parente.CurrentX, Mapa_Parente.CurrentY)-(Mapa_Parente.CurrentX, Mapa_Parente.CurrentY - LV), QBColor(0) 'Linha para cima
                    Mapa_Parente.Line (Mapa_Parente.CurrentX, Mapa_Parente.CurrentY)-(Mapa_Parente.CurrentX + (1000 * Lado), Mapa_Parente.CurrentY), QBColor(0) 'Linha para o lado (esquerdo ou direito depende da variável Lado)
                    Mapa_Parente.Line (Mapa_Parente.CurrentX, Mapa_Parente.CurrentY)-(Mapa_Parente.CurrentX, Mapa_Parente.CurrentY + LV), QBColor(0) 'Linha para baixo
                    If Passos < Len(Trilha) Then 'Passos não pode ser igual ou maior que len(trilha) por causa da função MID da linha abaixo.
                        Civil = IIf(Mid(Trilha, Passos + 1, 1) = 5 Or Mid(Trilha, Passos + 1, 1) = 6, 9, 45) 'Se tem filho (5) ou fiha (6) este "irmão" deve aparecer casado.
                        Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.CurrentX - (ImageList1.ListImages(Civil).Picture.Width / 4) + 110, Mapa_Parente.CurrentY
                        Mapa_Parente.CurrentY = Mapa_Parente.CurrentY + ImageList1.ListImages(Civil).Picture.Height / 2 + 10 'Ajusta abaixo do irmão
                        Mapa_Parente.CurrentX = Mapa_Parente.CurrentX + (ImageList1.ListImages(Civil).Picture.Width / 4) - 75 'Ajusta no meio do casal/ego
                    Else 'Quando o último termo basico é o irmão solteiro
                        Civil = 45
                        Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.CurrentX - (ImageList1.ListImages(Civil).Picture.Width / 4), Mapa_Parente.CurrentY
                        Mapa_Parente.CurrentY = Mapa_Parente.CurrentY + ImageList1.ListImages(Civil).Picture.Height / 2 + 10
                        Mapa_Parente.CurrentX = Mapa_Parente.CurrentX + 130
                    End If
                    If Passos = Len(Trilha) Then 'Chegou ao fim do termo técnico
                        Call Termo_Relacionado(0, Nome) 'Escreve o nome da pessoa focalizada
                    End If

                Case "4" 'Irmã
                    Mapa_Parente.Line (Mapa_Parente.CurrentX, Mapa_Parente.CurrentY)-(Mapa_Parente.CurrentX, Mapa_Parente.CurrentY - LV), QBColor(0)
                    Mapa_Parente.Line (Mapa_Parente.CurrentX, Mapa_Parente.CurrentY)-(Mapa_Parente.CurrentX + (1000 * Lado), Mapa_Parente.CurrentY), QBColor(0)
                    Mapa_Parente.Line (Mapa_Parente.CurrentX, Mapa_Parente.CurrentY)-(Mapa_Parente.CurrentX, Mapa_Parente.CurrentY + LV), QBColor(0)
                    If Passos < Len(Trilha) Then
                        Civil = IIf(Mid(Trilha, Passos + 1, 1) = 5 Or Mid(Trilha, Passos + 1, 1) = 6, 9, 46)
                        Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.CurrentX - (ImageList1.ListImages(Civil).Picture.Width / 2) + 20, Mapa_Parente.CurrentY 'O Ego masculino
                        Mapa_Parente.CurrentY = Mapa_Parente.CurrentY + ImageList1.ListImages(Civil).Picture.Height / 2 + 10
                        Mapa_Parente.CurrentX = Mapa_Parente.CurrentX - (ImageList1.ListImages(Civil).Picture.Width / 4) + 45
                    Else
                        Civil = 46
                        Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.CurrentX - (ImageList1.ListImages(Civil).Picture.Width / 4), Mapa_Parente.CurrentY 'O Ego masculino
                        Mapa_Parente.CurrentY = Mapa_Parente.CurrentY + ImageList1.ListImages(Civil).Picture.Height / 2 + 10
                        Mapa_Parente.CurrentX = Mapa_Parente.CurrentX - 40
                    End If
                    If Passos = Len(Trilha) Then
                        Call Termo_Relacionado(1, Nome) 'Escreve o nome da pessoa focalizada
                    End If

                Case "5" 'Filho
                    Mapa_Parente.Line (Mapa_Parente.CurrentX, Mapa_Parente.CurrentY)-(Mapa_Parente.CurrentX, Mapa_Parente.CurrentY + LV * 2), QBColor(0)
                    If Passos < Len(Trilha) Then
                        Civil = IIf(Mid(Trilha, Passos + 1, 1) = 5 Or Mid(Trilha, Passos + 1, 1) = 6, 9, 45)
                        Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.CurrentX - 90, Mapa_Parente.CurrentY 'O Ego masculino
                        Mapa_Parente.CurrentY = Mapa_Parente.CurrentY + (ImageList1.ListImages(Civil).Picture.Height / 2)
                        Mapa_Parente.CurrentX = Mapa_Parente.CurrentX + (ImageList1.ListImages(Civil).Picture.Width / 4 - 60)
                    Else
                        Civil = 45
                        Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.CurrentX - (ImageList1.ListImages(Civil).Picture.Width / 4), Mapa_Parente.CurrentY 'O Ego masculino
                        Mapa_Parente.CurrentY = Mapa_Parente.CurrentY + (ImageList1.ListImages(Civil).Picture.Height / 2)
                        Mapa_Parente.CurrentX = Mapa_Parente.CurrentX + (ImageList1.ListImages(Civil).Picture.Width / 4)
                    End If
                    If Passos = Len(Trilha) Then
                        Call Termo_Relacionado(0, Nome) 'Escreve o nome da pessoa focalizada
                    End If

                Case "6" 'Filha
                    Mapa_Parente.Line (Mapa_Parente.CurrentX, Mapa_Parente.CurrentY)-(Mapa_Parente.CurrentX, Mapa_Parente.CurrentY + LV * 2), QBColor(0)
                    If Passos < Len(Trilha) Then
                        Civil = IIf(Mid(Trilha, Passos + 1, 1) = 6 Or Mid(Trilha, Passos + 1, 1) = 5, 9, 46)
                        Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.CurrentX - (ImageList1.ListImages(Civil).Picture.Width / 2) + 30, Mapa_Parente.CurrentY 'O Ego masculino
                        Mapa_Parente.CurrentY = Mapa_Parente.CurrentY + (ImageList1.ListImages(Civil).Picture.Height / 2)
                        Mapa_Parente.CurrentX = Mapa_Parente.CurrentX - (ImageList1.ListImages(Civil).Picture.Width / 4 - 45)
                    Else
                        Civil = 46
                        Mapa_Parente.PaintPicture ImageList1.ListImages(Civil).Picture, Mapa_Parente.CurrentX - (ImageList1.ListImages(Civil).Picture.Width / 4), Mapa_Parente.CurrentY 'O Ego masculino
                        Mapa_Parente.CurrentY = Mapa_Parente.CurrentY + (ImageList1.ListImages(Civil).Picture.Height / 2)
                    End If
                    If Passos = Len(Trilha) Then
                        Call Termo_Relacionado(1, Nome) 'Escreve o nome da pessoa focalizada
                    End If

                Case "7" 'Esposo
                    'Se o primeiro TB é esposo e se é a primeira ocorrência e o ego feminino está selecionado, então...
                    If Left(Trilha, 1) = "7" And Passos = 1 And SSOption(17).Value = True Then
                        Mapa_Parente.CurrentY = (ImageList1.ListImages(Civil).Picture.Height / 2) + Espaço_V
                        Mapa_Parente.CurrentX = Mapa_Parente.Width / 2 + 30
                        Mapa_Parente.PaintPicture ImageList1.ListImages(17).Picture, Mapa_Parente.Width / 2 - (ImageList1.ListImages(17).Picture.Width / 4), Espaço_V 'O Ego masculino
                        If Passos = Len(Trilha) Then Call Termo_Relacionado(0, "Esposo") 'Escreve o nome da pessoa focalizada
                        Mapa_Parente.CurrentY = Espaço_V
                        Mapa_Parente.CurrentX = Mapa_Parente.CurrentX - 130

                    ElseIf Passos > 1 Then 'Se não é mais a primeira ocorrência
                        Mapa_Parente.PaintPicture ImageList1.ListImages(9).Picture, Mapa_Parente.CurrentX - (ImageList1.ListImages(9).Picture.Width / 2) + 80, Mapa_Parente.CurrentY - ImageList1.ListImages(9).Picture.Height / 2 'O Ego masculino
                        If Passos = Len(Trilha) Then Call Termo_Relacionado(0, "Esposo") 'Escreve o nome da pessoa focalizada
                        Mapa_Parente.CurrentY = Mapa_Parente.CurrentY - ImageList1.ListImages(9).Picture.Height / 2
                        Mapa_Parente.CurrentX = Mapa_Parente.CurrentX - 245
                    Else 'Caso contrário
                        Beep
                        Exit Sub 'Sai deste processo
                    End If
                Case "8" 'Esposa
                    If Left(Trilha, 1) = "8" And Passos = 1 And SSOption(16).Value = True Then
                        Mapa_Parente.CurrentY = (ImageList1.ListImages(Civil).Picture.Height / 2) + Espaço_V
                        Mapa_Parente.CurrentX = Mapa_Parente.Width / 2 + 30
                        Mapa_Parente.PaintPicture ImageList1.ListImages(13).Picture, Mapa_Parente.Width / 2 - (ImageList1.ListImages(13).Picture.Width / 4), Espaço_V 'O Ego masculino
                        If Passos = Len(Trilha) Then Call Termo_Relacionado(1, Nome) 'Escreve o nome da pessoa focalizada
                        Mapa_Parente.CurrentY = Espaço_V
                        Mapa_Parente.CurrentX = Mapa_Parente.CurrentX + 145
                    ElseIf Passos > 1 Then
                        Mapa_Parente.PaintPicture ImageList1.ListImages(9).Picture, Mapa_Parente.CurrentX - 130, Mapa_Parente.CurrentY - ImageList1.ListImages(9).Picture.Height / 2 'O Ego masculino
                        If Passos = Len(Trilha) Then Call Termo_Relacionado(1, Nome) 'Escreve o nome da pessoa focalizada
                        Mapa_Parente.CurrentY = Mapa_Parente.CurrentY - ImageList1.ListImages(9).Picture.Height / 2
                        Mapa_Parente.CurrentX = Mapa_Parente.CurrentX + 260
                    Else
                        Beep
                        Exit Sub
                    End If
            End Select
        Next Passos
        Debug.Print DBpa_Termos.Recordset("ID_Termo"); DBCombo(0).BoundText; DBCombo(0).Text

End Sub

Public Function Acha_Filhos(MeuCritério As String, Sexo As String, denovo As Integer, Trilha As String, Passos As Integer) As Integer
'Esta função procura por filhos seja qual for a geração solicitada.
'Verifique a descrição da variável Arvore() e Tree_Filhos() abaixo:
'Arvore(Índice do ego_selecionado, quantidade de filho quando o indice for xero _
 exemplo= arvore(41,0)=2 indica que o ego 41 tem dois TB"filho" na lista como segue: _
 arvore(41,1)=34 e arvore(41,2)=25
'Tree_Filhos(Ego_Inicial, Ego_Pai, Quantidade, camada) Quando a quantidade=0, _
 indicará quantos TB estão incluidos da lista de Ego_Inicial. _
 Esta variável tem por finalidade guarda a ligação entre o ego inicial da pesquisa _
 e o ego pai dos filhos que porventura são encontrados. _

    Dim última_posição As Variant
    Dim MeuCritério_Casal As String
    última_posição = DBcp_Ego.Recordset.Bookmark
    MeuCritério_Casal = MeuCritério
    DBcp_Casais.Recordset.FindFirst MeuCritério_Casal
    'Se não encontrar outro ego com ID_PAIS igual, _
     então o processo termina vai para outro ego.
    If DBcp_Casais.Recordset.NoMatch = True Then
        Acha_Filhos = 0
        Exit Function
    End If
    'Este é loop que cicla pelas famílias de orientação do ego considerado.
    Do While DBcp_Casais.Recordset.NoMatch = False
        'Este critério procura egos masculinos que tem os mesmos pais.
        MeuCritério = "ID_Pais =" & DBcp_Casais.Recordset("ID_Casal") _
                      & " and " & Sexo
        'O Critério está pronto para a primeira busca.
        DBcp_Ego.Recordset.FindFirst MeuCritério
        'Se não encontrar um ego com ID_PAIS igual ao solicitado, então _
         o processo termina e volta para pesquisar se o ego considerado tem _
         outra familia de procriação.
        Do While DBcp_Ego.Recordset.NoMatch = False
            n = n + 1 'Contador - Indica o número de filhos encontrados.
            ID_Ego_TB = DBcp_Ego.Recordset("ID_Ego") 'Filho do ego considerado
            'O teste seguinte é necessário para que o ego seja incluído apenas uma vez _
             na lista de egos. Tem que ser quando encontrar o primeiro filho _
             na geração considerada.
            If n = 1 And Passos = Len(Trilha) And Arvore(Ego_Inicial, 0, 0) = "" Then 'Inclui na lista de egos apenas uma vez
                'Inclui o nome do corrente ego na lista de egos.
                List(5).AddItem DB_Temp.Recordset("Nome_Ind")
                'Pega o ID_Ego do corrente ego para uso posterior e associa no itemdata.
                List(5).ItemData(List(5).NewIndex) = Ego_Inicial
            End If
            'Verifica se o corrente ego já tem algum """"filho"""" na sua lista
            If Arvore(Ego_Inicial, 0, 0) <> "" Then 'Se tem filho então...
                'Ao número total do filhos existentes é somado mais um, pois vai _
                 entrar mais um filho na lista.
                n = CInt(Arvore(Ego_Inicial, 0, 0)) + 1
            End If
            'O Filho só é incluido na lista se este estiver na última geração considerada _
             ou se o próximo TB for "8=Esposa" ou "7=Esposo", conforme o caso, pois assim será necessário a lista pronta dos filhos.
            If (denovo = Camada And Passos = Len(Trilha)) Or Mid(Trilha, Passos + 1, 1) = "8" Or Mid(Trilha, Passos + 1, 1) = "7" Then
                Arvore(Ego_Inicial, n, 0) = DBcp_Ego.Recordset("ID_Ego")
                Arvore(Ego_Inicial, n, 1) = DBcp_Ego.Recordset("Nome_Ind")
                'O número total de filhos na lista é atualizado.
                Arvore(Ego_Inicial, 0, 0) = CStr(n)
            End If
            'Está variável guarda o ID_Ego do filho mantendo a ligação entre _
             o ego inicial e o ego pai do filho encontrado.
            Tree_Filhos(DB_Temp.Recordset("id_ego"), DBcp_Casais.Recordset("ID_Conj1"), n, denovo) = CInt(ID_Ego_TB)
            'Aqui a variável guarda os filhos diretos do ego inicial, _
             por isso o ID_Ego do ego inicial aparece também na posição do Ego_Pai.@@@@@
            If denovo = 1 And Passos = 1 And Passos = Len(Trilha) Then Tree_Filhos(DB_Temp.Recordset("id_ego"), DB_Temp.Recordset("id_ego"), n, denovo) = CInt(ID_Ego_TB)
            Proximo_Casal = 0 'Nenhum casal é selecionado.
            ID_Ego_TB = -1 'Nenhum ego relacionado com o TB está selecionado.
            'Procura o próximo filho do ego considerado.
            DBcp_Ego.Recordset.FindNext MeuCritério
        Loop
        'Esta é a soma total de todos o filhos encontrados na geração considerada.
        Total_filhos_Camada = Total_filhos_Camada + CInt(n)
        n = 0 'O contador de filhos é zerado aqui.
        'Procura por outra família de procrição do ego considerado.
        DBcp_Casais.Recordset.FindNext MeuCritério_Casal
    Loop
    'Está variável guarda o número total de filhos mantendo a ligação entre _
     o ego inicial e o ego pai dos filhos encontrados.
    Tree_Filhos(DB_Temp.Recordset("id_ego"), DBcp_Casais.Recordset("ID_Conj1"), 0, denovo) = _
    IIf(Total_filhos_Camada = 0, "", Total_filhos_Camada)
    'Aqui a variável guarda o número total de filhos diretos do ego inicial, _
     por isso o ID_Ego do ego inicial aparece também na posição do Ego_Pai.
    If denovo = 1 And Passos = 1 And Passos = Len(Trilha) Then _
    Tree_Filhos(DB_Temp.Recordset("id_ego"), DB_Temp.Recordset("id_ego"), 0, denovo) = _
    IIf(Total_filhos_Camada = 0, "", Total_filhos_Camada)
    'Volta para a posição do registro em DBcp_Ego que estava antes de entrar nesta função.
    DBcp_Ego.Recordset.Bookmark = última_posição
    'Indica que algum filho foi encontrado em alguma instância. Pode ser usado posteriormente.
    Acha_Filhos = 1
End Function

Public Sub Acha_Esposas(MeuCritério As String, Passos As Integer, Trilha As String)
'Este processo procura por esposas dos ego considerado.
    'Pronto para a primeira busca.
    DBcp_Casais.Recordset.FindFirst MeuCritério
    'Este é o loop que cicla pelas famílias de procriação do ego considerado.
    Do While DBcp_Casais.Recordset.NoMatch = False
        n = n + 1 'Contador - Indica o número de esposas encontradas.
        'Este critério procura pelos dados da esposa no cadastro geral, _
         pois vamos precisar do nome dela para incluir na lista de parentes.
        qwqwq = Left(MeuCritério, 8)
        If Left(MeuCritério, 8) = "ID_Conj1" Then
            Esposa_Critério = "ID_Ego = " & DBcp_Casais.Recordset("ID_Conj2")
        ElseIf Left(MeuCritério, 8) = "ID_Conj2" Then
            Esposa_Critério = "ID_Ego = " & DBcp_Casais.Recordset("ID_Conj1")
        End If
        'Pronto para a busca dos dados da esposa.
        DBcp_Ego.Recordset.FindFirst Esposa_Critério
        'O teste seguinte é necessário para que o ego seja incluído apenas uma vez _
         na lista de egos. Tem que ser quando encontrar a primeira esposa.
        If n = 1 And Passos = Len(Trilha) And Arvore(Ego_Inicial, 0, 0) = "" Then
            'Inclui o nome do corrente ego na lista de egos.
            List(5).AddItem DB_Temp.Recordset("Nome_Ind")
            'Pega o ID_Ego do corrente ego para uso posterior e associa no itemdata.
            List(5).ItemData(List(5).NewIndex) = DB_Temp.Recordset("id_ego")
        End If
        'Verifica se o corrente ego já tem algum """"esposa"""" na sua lista. Podemos estar _
         tratando aqui da esposa do irmão do ego por exemplo.
        If Arvore(Ego_Inicial, 0, 0) <> "" Then 'Se tem esposa então...
            'Ao número total das esposas existentes é somado mais um, pois vai _
             entrar mais uma esposa na lista.
            n = CInt(Arvore(Ego_Inicial, 0, 0)) + 1
        End If
        'Guarda o ID_Ego da esposa nesta variável para uso posterior.
        Arvore(Ego_Inicial, n, 0) = DBcp_Ego.Recordset("ID_Ego")
        'Guarda o Nome da esposa nesta variável para uso posterior.
        Arvore(Ego_Inicial, n, 1) = DBcp_Ego.Recordset("Nome_Ind")
        
        'Procura pelo próximo casamento do corrente ego.
        DBcp_Casais.Recordset.FindNext MeuCritério
        'Grava aqui o número total de esposas encontradas.
        Arvore(Ego_Inicial, 0, 0) = n
    Loop
    n = 0 'Zera o contator

End Sub

Public Function TB_Pais(MaisUm As Integer, Proximo_Casal As Integer, Qual_Conjuge As String, Passos As Integer, Esposa_Outra_Camada As Integer, Trilha As String, Nova_Trilha As String) As Integer
    DBcp_Ego.Recordset.FindFirst "ID_Ego= " & Procurar_De_Quem(MaisUm)
    'Para evitar uma variável nula usei a variável Zero como teste, pois _
     o DB_Temp.Recordset("ID_Pais") pode se nulo(Isto significa que o pai _
     do ego conderado não foi cadastrado) e isto casaria um erro.
    Zero = IIf(IsNull(DBcp_Ego.Recordset("ID_Pais")) = True, 0, DBcp_Ego.Recordset("ID_Pais"))
    'Pega o ID do casal para ser pesquisado.
    Proximo_Casal = IIf(Proximo_Casal <> 0, Proximo_Casal, Zero)
    If Proximo_Casal <> 0 Then
        MeuCritério = "ID_Casal = " & Proximo_Casal
        DBcp_Casais.Recordset.FindFirst MeuCritério 'Procura pelo casal no DB.
        MeuCritério = "ID_Ego=" & DBcp_Casais.Recordset(Qual_Conjuge)
        DBcp_Ego.Recordset.FindFirst MeuCritério 'Procura o homem no DB
        ID_Ego_TB = DBcp_Ego.Recordset("ID_Ego") 'Recebe o ID do homem que responde pelo TB.
        If Passos = Len(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha)) Then 'Se for o último TB pesquizado
            'Inclui o nome do corrente ego na lista de egos, pois ele tem um pai.
            List(5).AddItem DB_Temp.Recordset("Nome_Ind")
            'Pega o ID_Ego do corrente ego para uso posterior e associa no itemdata.
            List(5).ItemData(List(5).NewIndex) = DB_Temp.Recordset("ID_Ego")
            'Guarda o ID_Ego do Pai nesta variável para uso posterior.
            Arvore(Ego_Inicial, 1, 0) = DBcp_Ego.Recordset("ID_Ego") 'teste
            'Guarda o Nome do Pai nesta variável para uso posterior.
            Arvore(Ego_Inicial, 1, 1) = DBcp_Ego.Recordset("Nome_Ind") 'teste
            Arvore(Ego_Inicial, 0, 0) = 1
            Proximo_Casal = 0 'Nenhum casal selecionado.
            ID_Ego_TB = -1 'A variável é zerada, para uso posterior.
        End If
        'Caso o ID_PAIS está vazio o processo termina aqui com este ego, pois ele _
         não tem pai cadastrado
        If IsNull(DBcp_Ego.Recordset("ID_Pais")) = True Then
            Proximo_Casal = 0 'Nenhum casal selecionado.
            ID_Ego_TB = -1 'A variável é zerada, para uso posterior.
            Exit Function 'For 'Sair do for-next que seleciona o TB e vai em para outro ego.
        Else 'Caso tem um pai....
            'Seleciona o corrente casal, pois este casal pode continuar a ser _
             pesquisado no próximo for-next.
            Proximo_Casal = DBcp_Ego.Recordset("ID_Pais")
        End If
        TB_Pais = Proximo_Casal
    End If
    
End Function

Public Function TB_Irmãos(MaisUm As Integer, Proximo_Casal As Integer, ID_Ego_TB As Integer, Sexo As String, Passos As Integer, Esposa_Outra_Camada As Integer, Trilha As String, Nova_Trilha As String) As Integer
    'Entra aqui se estamos procurando parente da esposa.
    If Procurar_De_Quem(MaisUm) <> DB_Temp.Recordset("ID_Ego") Then
        DBcp_Ego.Recordset.FindFirst "ID_Ego= " & Procurar_De_Quem(MaisUm)
        'Para evitar uma variável nula usei a variável Zero como teste, pois _
         o DB_Temp.Recordset("ID_Pais") pode se nulo(Isto significa que o pai _
         do ego conderado não foi cadastrado) e isto casaria um erro.
        Zero = IIf(IsNull(DBcp_Ego.Recordset("ID_Pais")) = True, 0, DBcp_Ego.Recordset("ID_Pais"))
    'Aqui os parentes do ego inicial é procurado.
    Else
        Zero = IIf(IsNull(DB_Temp.Recordset("ID_Pais")) = True, 0, DB_Temp.Recordset("ID_Pais"))
    End If
     'denovo = 0 'zera esta variável caso seja necessário entrar em filhos depois daqui.
     'Pega o ID do casal para ser pesquisado, precisa achar os pai do _
      corrente ego para achar os seus irmãos.
     Proximo_Casal = IIf(Proximo_Casal <> 0, Proximo_Casal, Zero)
    'Proximo_Casal = 0 significa que o corrente ego não tem pais cadastrados
    If Proximo_Casal <> 0 Then
        'Este critério procura os egos masculinos diferente do corrente ego _
         com os mesmos os pais.
        MeuCritério = "ID_Pais = " & Proximo_Casal & " and " _
                      & "ID_Ego <>" & IIf(ID_Ego_TB = -1, CStr(Procurar_De_Quem(MaisUm)), ID_Ego_TB) & " and " _
                      & Sexo
        DBcp_Ego.Recordset.FindFirst MeuCritério
        'Se não encontrar outro ego com ID_PAIS igual, _
         então o processo termina vai para outro ego.
        If DBcp_Ego.Recordset.NoMatch = True Then Exit Function
        ID_Ego_TB = DBcp_Ego.Recordset("ID_Ego")
        'O loop abaixo vai achar os possíveis irmãos e armazená-los _
         na vareável arvore().
        Qt = 0 'Contador
        Do While DBcp_Ego.Recordset.NoMatch = False 'Se não encontrar outro ego com ID_PAIS igual, então o processo termina vai para outro ego.
            Qt = Qt + 1 'Acrescenta mais um ao contador.
            'Guarda o ID_Ego do Pai nesta variável para uso posterior.
            Arvore(Ego_Inicial, Qt, 0) = DBcp_Ego.Recordset("ID_Ego")
            'Guarda o Nome do Pai nesta variável para uso posterior.
            Arvore(Ego_Inicial, Qt, 1) = DBcp_Ego.Recordset("Nome_Ind")
            DBcp_Ego.Recordset.FindNext MeuCritério 'Procura por outro irmão
        Loop
        'Grava aqui o número total de irmãos encontrados.
        Arvore(Ego_Inicial, 0, 0) = Qt
        'If DB_Temp.Recordset("ID_Ego") = 0 Then Stop
        If Passos = Len(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha)) Then 'Se for o último TB pesquizado
            'Inclui o nome do corrente ego na lista de egos, pois ele tem irmãos.
            List(5).AddItem DB_Temp.Recordset("Nome_Ind")
            'Pega o ID_Ego do corrente ego para uso posterior e associa no itemdata.
            List(5).ItemData(List(5).NewIndex) = Ego_Inicial 'teste ID_Ego_TB 'Pega o ID do parente e associa no itemdata.
            If MaisUm = Procurar_De_Quem(0) Then
                Esposa_Outra_Camada = 0
                'Arvore(Ego_Inicial, 0, 0) = ""
            End If
            Proximo_Casal = 0 'Nenhum casal selecionado.
            ID_Ego_TB = -1 'A variável é zerada, para uso posterior.
        End If
        'Caso o ID_PAIS está vazio o processo termina aqui com este ego pois ele _
         não tem irmãos associados e cadastrados no DB
        If IsNull(DBcp_Ego.Recordset("ID_Pais")) = True Then
            Proximo_Casal = 0 'Nenhum casal selecionado.
            ID_Ego_TB = -1 'A variável é zerada, para uso posterior.
            Exit Function 'Sair do for-next que seleciona o TB e vai em para outro ego.
        End If
    Else 'O corrente ego não tem pais cadastrados, por isso... _
          ...vai sair deste processo e procurar outro ego.
        Exit Function
    End If
    TB_Irmãos = Proximo_Casal
End Function

Public Function TB_Prole(TB As String, denovo As Integer, Sexo As String, MaisUm As Integer, Proximo_Casal As Integer, ID_Ego_TB As Integer, Passos As Integer, Esposa_Outra_Camada As Integer, Trilha As String, Nova_Trilha As String) As Integer
    
    Dim MeuCritério As String
    Dim Irmão_Critério As String
    denovo = denovo + 1 'Conta quantas vezes entrou aqui com o mesmo ego. Volta para 1 quando mudar o ego
    'Ajusta para saber até que camada vai. Camada não muda até mudar o Termo.
    If denovo > Camada Then Camada = denovo
    'Descobrir QUEM será pesquisado
    If Passos = 1 Then 'Significa que estamos procurando os filhos do ego em DB_Temp
        'Descobre se o ego é H ou M e associa o ID_Conj correto.
        Qual_Conj = IIf(DB_Temp.Recordset("sexo") = 0, "ID_Conj1", "ID_Conj2")
        MeuCritério = Qual_Conj & "= " & CStr(DB_Temp.Recordset("ID_ego"))
        'Chama a função que vai procurar os filhos do corrente ego.
        Call Acha_Filhos(MeuCritério, Sexo, denovo, IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha), Passos)
    Else 'Estamos procurando os filhos de outros que não do próprio ego...
         'Pode ser filhos de outro filho ou filho de irmão(ã)#####
        If denovo > 1 Then 'Já passou aqui com o mesmo ego, portanto estamos procurando filho de filho..
            'Ego_Inicial = DB_Temp.Recordset("id_ego") 'Mantendo a ligação com o ego inicial
            'É necessário passar por todos os egos para saber quem são os filhos da camada anterior.
            'Estes egos filhos serão agora os pais (Ego_Pai)
            For Ego_Pai = 0 To DBcp_Ego.Recordset.RecordCount - 1
                n = 1 'Inicia a variável que dará o número de filhos no final
                Conta_Pai = Ego_Pai
                'Com cada ego_pai considerado, nós procuramos se há filhos.
                If Tree_Filhos(Ego_Inicial, Conta_Pai, 0, denovo - 1) <> "" Then
                    'Com cada ego_pai considerado, nós procuramos se há filhos.
                    'Lembra do formato do Tree_filhos para o início de cada camada(Ego_Inicial, ego_pai, 1, camada) _
                     Isto indica se há filhos e quantos filhos naquela camada.
                     If Ego_Inicial = 18 Then Stop
                     Pula = 0
                    For conta_filho = 1 To CInt(Tree_Filhos(Ego_Inicial, Conta_Pai, 0, denovo - 1))
                        'Da 2ª camada em diante o pai sempre é diferente o ego inicial, _
                         pois não existe um pai sendo pai dele mesmo.
                        If Ego_Pai = Ego_Inicial And denovo > 2 Then Exit For  'Se prosseguir causa erro.
                        'Evita chamar a função Achar_Filhos sem motivo, pois o Tree_Filhos estará vazio.
                        If Tree_Filhos(Ego_Inicial, Conta_Pai, conta_filho, denovo - 1) = "" Then Exit For
                        Novo_Pai = CInt(Tree_Filhos(Ego_Inicial, Conta_Pai, conta_filho, denovo - 1))
                        'É melhor deixa a busca tanto no lado masculino como feminino para facilitar _
                         o código aqui.
                        MeuCritério = "ID_Conj1=" & CStr(Novo_Pai) & "or ID_Conj2=" & CStr(Novo_Pai) 'O novo_pai vai ser procurado no DB_Casais
                        'Caso estejamos procurando filho da filha ou filha do filho é necessário _
                         fazer a troca da variável Sexo, senão vai procurar filha entre os filhos _
                         e filho entre as filhas.
                        If TB = "5" And Sexo = "Sexo=1" Then
                            Sexo = "Sexo=0"
                        ElseIf TB = "6" And Sexo = "Sexo=0" Then
                            Sexo = "Sexo=1"
                        End If
                        Call Acha_Filhos(MeuCritério, Sexo, denovo, IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha), Passos) ' = 0 Then Exit For
                    Next conta_filho
                End If
            Next Ego_Pai
        'Primeira vez que entra aqui, mas já passou por outro TB, _
         então procuramos filho do(a) irmão(ã)
        ElseIf denovo = 1 Then
            'If DB_Temp.Recordset("id_ego") = 18 Then Stop
            If Mid(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha), Passos - 1, 1) = "3" Then
                'Caso estejamos procurando filho do irmão (do, da)...
                MeuCritério = "ID_Conj1= "
            ElseIf Mid(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha), Passos - 1, 1) = "4" Then
                'Caso estejamos procurando fiho da irmã (do, da)...
                MeuCritério = "ID_Conj2= "
            End If
            'Já sabemos quantos irmãos(ãs) o ego em db_temp tem, pois _
             o processo já passou pela seção de busca de irmãos(ãs). _
             Carreguei os irmãos nas variáveis Cada_Irmão() e esvaziei _
             as variáveis Arvore() por segurança, pois esta variável _
             será manipulada pela função Acha_filhos()
            If Arvore(Ego_Inicial, 0, 0) <> "" Then
                qt_irmão = CInt(Arvore(Ego_Inicial, 0, 0))
                ReDim Cada_Irmão(qt_irmão) As Integer
                For Sequ = 1 To qt_irmão
                    Cada_Irmão(Sequ) = CInt(Arvore(Ego_Inicial, Sequ, 0))
                    Arvore(Ego_Inicial, Sequ, 0) = ""
                    Arvore(Ego_Inicial, Sequ, 1) = ""
                Next Sequ
                Arvore(Ego_Inicial, 0, 0) = ""
                For Qt = 1 To qt_irmão
                    Irmão_Critério = MeuCritério & Cada_Irmão(Qt)
                    'Esta função será chamada com cada irmão selecionado.
                    Call Acha_Filhos(Irmão_Critério, Sexo, denovo, IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha), Passos)
                Next Qt
            End If
        End If
    End If
    TB_Prole = denovo
End Function

Public Function TB_Conjuges(Conjuge_Ego As String, Esposa_Outra_Camada As Integer, Trilha As String, Nova_Trilha As String, Passos As Integer) As Integer
    Dim MeuCritério As String
    Dim Irmão_Critério As String
    'Só entra aqui se está procurando a esposa do ego inicial ou se
    If Procurar_De_Quem(1) = DB_Temp.Recordset("ID_Ego") Or Passos > 2 Then
        'Caso a busca seja pela(s) esposa(s) do corrente ego...
        If Len(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha)) = 1 Then
            'Este critério procura pelas uniões feitas pelo corrente ego. _
             Observe que ele pode ter ou teve várias famílias de procriação. _
             As uniões desfeitas por separação não entram aqui.
            MeuCritério = Conjuge_Ego & DB_Temp.Recordset("ID_Ego") _
                           & " and Civil <> 5"
            Call Acha_Esposas(MeuCritério, Passos, IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha))
        'Vai entrar aqui quando procura a(s) esposa(s) que não _
         sejam do próprio ego.
        Else
            'Esta variável pega o termo básico anterior ao tb esposa. Estamos _
             procurando esposa(s) do pai, do irmão e do filho.
            TB_Anterior = Mid(IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha), Passos - 1, 1)
            Select Case TB_Anterior
                Case "1", "2" 'Pai ou Mãe
                    'Lembre que o ID do pais está guardado nesta variável: Arvore(ego_inicial, 1, 0)
                    'Este critério procura pelas uniões feitas pelo pai do ego. _
                     Estamos procurando por aquelas esposas do pai que não seja _
                     a própria mãe do ego. As uniões desfeitas pelo pai não importam aqui.
                    MeuCritério = Conjuge_Ego & Arvore(Ego_Inicial, 1, 0) & _
                                   " and ID_Casal<> " & DB_Temp.Recordset("ID_Pais") & _
                                   " and Civil<> 5"
                    Call Acha_Esposas(MeuCritério, Passos, IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha))
                Case "3", "4" 'Irmão ou Irmã
                    'Já sabemos quantos irmãos o ego em db_temp tem, pois _
                     o processo já passou pela seção de busca de irmãos. _
                     Carreguei os irmãos nas variáveis Cada_Irmão() e esvaziei _
                     as variáveis Arvore() por segurança, pois esta variável _
                     será manipulada pela função Acha_Esposas()
                    If Arvore(Ego_Inicial, 0, 0) <> "" Then
                        qt_irmão = CInt(Arvore(Ego_Inicial, 0, 0))
                        ReDim Cada_Irmão(qt_irmão) As Integer
                        For Sequ = 1 To qt_irmão
                            Cada_Irmão(Sequ) = CInt(Arvore(Ego_Inicial, Sequ, 0))
                            Arvore(Ego_Inicial, Sequ, 0) = ""
                            Arvore(Ego_Inicial, Sequ, 1) = ""
                        Next Sequ
                        Arvore(Ego_Inicial, 0, 0) = ""
                        For Qt = 1 To qt_irmão
                            'Este critério procura as esposas dos irmãos atuais ou que já morreram. _
                             As esposas de casamentos desfeitos por separação não contam.
                            Irmão_Critério = Conjuge_Ego & Cada_Irmão(Qt) & " and Civil<> 5"
                            'Esta função será chamada com cada irmão selecionado.
                            Call Acha_Esposas(Irmão_Critério, Passos, IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha))
                        Next Qt
                    End If
                Case "5", "6" 'Filho ou Filha
                    'Já sabemos quantos filhos o ego em db_temp tem, pois _
                     o processo já passou pela seção de busca de filhos. _
                     Carreguei os filhos nas variáveis Cada_Irmão() e esvaziei _
                     as variáveis Arvore() por segurança, pois esta variável _
                     será manipulada pela função Acha_Esposas()
                    If Arvore(Ego_Inicial, 0, 0) <> "" Then
                        qt_irmão = CInt(Arvore(Ego_Inicial, 0, 0))
                        ReDim Cada_Irmão(qt_irmão) As Integer
                        For Sequ = 1 To qt_irmão
                            Cada_Irmão(Sequ) = CInt(Arvore(Ego_Inicial, Sequ, 0))
                            Arvore(Ego_Inicial, Sequ, 0) = ""
                            Arvore(Ego_Inicial, Sequ, 1) = ""
                        Next Sequ
                        Arvore(Ego_Inicial, 0, 0) = ""
                        For Qt = 1 To qt_irmão
                            'Este critério procura as esposas dos irmãos atuais ou que já morreram. _
                             As esposas de casamentos desfeitos por separação não contam.
                            Irmão_Critério = Conjuge_Ego & Cada_Irmão(Qt) & " and Civil<> 5"
                            'Esta função será chamada com cada irmão selecionado.
                            Call Acha_Esposas(Irmão_Critério, Passos, IIf(Esposa_Outra_Camada = 0, Trilha, Nova_Trilha))
                        Next Qt
                    End If
            End Select
            
        End If
        'Caso nenhuma esposa seja encontrada, então o For-Next é interrompido _
         para que outro ego inicial seja escolhido para outra busca.
        If Arvore(Ego_Inicial, 0, 0) = "" Then Exit Function
        'Caso estamos procurando parentes de esposas do irmão, pai ou filho. Isto _
         siginifica que o termo tecnico não é composto apenas de Esposa, nem o primeiro TB e _
         nem o último TB é Esposa. Left(Trilha, 1) <> "8")
         'If DB_Temp.Recordset("ID_Ego") = 0 Then Stop
         qqqq = Arvore(41, 0, 0)
         Sequ = 0
         posi = IIf(InStr(Trilha, "8") = 2, 2, 3)
         Sequ = InStr(posi, Trilha, 8)
        If Len(Trilha) > 2 And Sequ <> 0 And Right(Trilha, 1) <> "8" Then
            Nova_Trilha = Mid(Trilha, Passos + 1, Len(Trilha) - (Passos))
            Esposa_Outra_Camada = 1
            pulaUma = 1
            TB_Conjuges = pulaUma
            Exit Function
        
        End If
    End If
    TB_Conjuges = 0
End Function
