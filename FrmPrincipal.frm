VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPrincipal 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7680
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4590
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cadastro"
      TabPicture(0)   =   "FrmPrincipal.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrmBotoes"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TxtCodigo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtDescricaoCompleta"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TxtEAN"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "FrmPrincipal.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFlexGrid"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   15
         Top             =   1560
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5106
         _Version        =   393216
      End
      Begin VB.TextBox TxtEAN 
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox TxtDescricaoCompleta 
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   6015
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filtros"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   7
         Top             =   360
         Width           =   7455
         Begin VB.CommandButton CmdPesquisar 
            Caption         =   "Pesquisar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5880
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame FrmBotoes 
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   3240
         Width           =   7335
         Begin VB.CommandButton CmdSair 
            Caption         =   "Sair"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   5880
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton CmdExcluir 
            Caption         =   "Excluir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4440
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton CmdEditar 
            Caption         =   "Editar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3000
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton CmdIncluir 
            Caption         =   "Incluir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1560
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton CmdNovo 
            Caption         =   "Novo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Código de barras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Descrição completa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub LimparCampos()
    TxtCodigo.Text = ""
    TxtDescricaoCompleta.Text = ""
    TxtEAN.Text = ""
End Sub

Private Sub CmdNovo_Click()
    LimparCampos
End Sub

Sub PreencherFlexGrid()
    MSFlexGrid.Cols = 4
    MSFlexGrid.ColWidth(0) = 200  'Selecao
    MSFlexGrid.ColWidth(1) = 1000 'Codigo
    MSFlexGrid.ColWidth(2) = 5000 'Descricao
    MSFlexGrid.ColWidth(3) = 3000 'Ean
    
    MSFlexGrid.TextMatrix(0, 0) = "*"
    MSFlexGrid.TextMatrix(0, 1) = "Codigo"
    MSFlexGrid.TextMatrix(0, 2) = "Descrição Completa"
    MSFlexGrid.TextMatrix(0, 3) = "Código de Barras"
    
    'teste
    MSFlexGrid.TextMatrix(1, 0) = ""
    MSFlexGrid.TextMatrix(1, 1) = "1"
    MSFlexGrid.TextMatrix(1, 2) = "COCA COLA 2L"
    MSFlexGrid.TextMatrix(1, 3) = "7894900011517"
    
End Sub

Private Sub CmdPesquisar_Click()
    PreencherFlexGrid
End Sub
