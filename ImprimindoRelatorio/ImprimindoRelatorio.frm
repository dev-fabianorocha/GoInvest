VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   6840
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "ImprimindoRelatorio.frx":0000
      Height          =   2655
      Left            =   240
      OleObjectBlob   =   "ImprimindoRelatorio.frx":0014
      TabIndex        =   12
      Top             =   3120
      Width           =   7695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\BASESQL\teste.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CLIENTES"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      DataField       =   "CLI_RG"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      DataField       =   "CLI_CPF"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "CLI_ENDERECO"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1440
      Width           =   6015
   End
   Begin VB.TextBox Text3 
      DataField       =   "CLI_CADASTRO"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "CLI_NOME"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   720
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      DataField       =   "CLI_CODIGO"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "RG"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "CPF"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Endereço"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Data de Nascimento"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Nome"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Banco As Database
Dim Tabela As Recordset
Private Sub Command1_Click()
   
   FrmRelCli.Show
   
   
End Sub
Public Sub Cabecalho()
   Printer.Font = "Arial"
   Printer.FontBold = True
   Printer.FontSize = 11

   Printer.Print Tab(35); "Relatório de Clientes"
   Printer.Print
   Printer.Print Tab(5); "Código";
   Printer.Print Tab(15); "Nome";
   Printer.Print Tab(40); "Endereço";
   Printer.Print Tab(60); "Data de Nascimento";
   Printer.Print Tab(85); "RG";
   Printer.FontBold = False
End Sub
