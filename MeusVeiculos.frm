VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{7F2E2FCC-611E-11D1-890F-F48FE777651C}#106.0#0"; "Controles.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MeusVeiculos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meus Veiculos"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10770
   ForeColor       =   &H00404000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin Controles.ctlBotao Recarregar 
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Recarrega"
      Formato         =   35
   End
   Begin Controles.ctlBotao Voltar 
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Voltar"
      Formato         =   13
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5175
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9128
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      FlatScrollBar   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Placa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Modelo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Marca"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Lugares"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Descrição"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "url_imagem"
         Text            =   "Url"
         Object.Width           =   2540
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _Version        =   65536
      _ExtentX        =   18018
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Listagem de veiculos"
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MeusVeiculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim Rec As New ADODB.Recordset

Private Lista As ListView

Private Sub listagem()
    Con.Open "DSN=MSHOP30;Database=marcos;uid=sa;pwd=DBd4t43xp0rt@;"
        
    Rec.Open "Select * from veiculo ", Con, adOpenStatic, adLockReadOnly
    
    Do Until Rec.EOF
       Set lastItems = ListView1.ListItems.Add(, , Rec!id)
       lastItems.SubItems(1) = Rec!placa
       lastItems.SubItems(2) = Rec!modelo
       lastItems.SubItems(3) = Rec!marca
       lastItems.SubItems(4) = Rec!quantidade_lugares
       lastItems.SubItems(5) = Rec!descricao
       lastItems.SubItems(6) = Rec!url_imagem
       Rec.MoveNext
    Loop
    
    Rec.Close
    Con.Close
End Sub

Private Sub Recarregar_Click()
    ListView1.ListItems.Clear
    Call listagem
End Sub

Private Sub Voltar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call listagem
End Sub
