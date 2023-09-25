VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{7F2E2FCC-611E-11D1-890F-F48FE777651C}#106.0#0"; "Controles.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form CadastrarVeiculo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10155
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel StatusSite 
      Height          =   615
      Left            =   7560
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "STATUS NÃO INFORMADO"
      ForeColor       =   8388608
      BackColor       =   15790320
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
   Begin Controles.ctlBotao RemoverSite 
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   8400
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      Caption         =   "Remover Site"
      Formato         =   25
   End
   Begin Controles.ctlBotao PublicaSite 
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   7800
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      Caption         =   "Publicar Site"
      Formato         =   8
   End
   Begin Threed.SSCheck DuplicarCadastro 
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   9600
      Visible         =   0   'False
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Duplicar Cadastro"
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
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3735
      Left            =   7560
      TabIndex        =   12
      Top             =   1560
      Width           =   4575
      ExtentX         =   8070
      ExtentY         =   6588
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin Controles.ctlTexto url 
      Height          =   675
      Left            =   360
      TabIndex        =   11
      Top             =   3600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1191
      Caption         =   "Imagem"
      BackColor       =   16777215
      ForeColor       =   0
   End
   Begin Controles.ctlBotao listagem 
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   7800
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      Caption         =   "Listagem"
      Formato         =   33
   End
   Begin RichTextLib.RichTextBox descricao 
      Height          =   2295
      Left            =   360
      TabIndex        =   8
      Top             =   4800
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4048
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"CadastroVeiculo.frx":0000
   End
   Begin Controles.ctlNumero id 
      Height          =   675
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1191
      Caption         =   "ID"
      BackColor       =   16777215
      ForeColor       =   0
      Enabled         =   0   'False
      Locked          =   -1  'True
      MaxLength       =   15
   End
   Begin Controles.ctlBotao Pesquisar 
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   1200
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   873
      Caption         =   "Pesquisar"
      Formato         =   4
   End
   Begin Controles.ctlBotao Excluir 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   9000
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      Caption         =   "Excluir"
      Formato         =   0
   End
   Begin Controles.ctlBotao Gravar 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   8400
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      Caption         =   "Salvar"
   End
   Begin Controles.ctlNumero quantidade_lugares 
      Height          =   675
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   1191
      Caption         =   "Qtd Lugares"
      BackColor       =   16777215
      ForeColor       =   0
      MaxLength       =   15
   End
   Begin Controles.ctlPlaca placa 
      Height          =   675
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1191
      Caption         =   "Placa"
      BackColor       =   16777215
      ForeColor       =   0
      MaxLength       =   10
   End
   Begin Controles.ctlTexto marca 
      Height          =   675
      Left            =   3960
      TabIndex        =   1
      Top             =   2760
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   1191
      Caption         =   "Marca"
      BackColor       =   16777215
      ForeColor       =   0
   End
   Begin Controles.ctlTexto modelo 
      Height          =   675
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   1191
      Caption         =   "Modelo"
      BackColor       =   16777215
      ForeColor       =   0
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   240
      Width           =   11895
      _Version        =   65536
      _ExtentX        =   20981
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "Cadastrar Veiculos"
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
   Begin VB.Label FotoVeiculo 
      Caption         =   "Foto Veiculo"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   20
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Ações Cadastrais"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label LabelSiteAcoes 
      Caption         =   "Ações Site:"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Descrição 
      Caption         =   "Descrição"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   975
   End
End
Attribute VB_Name = "CadastrarVeiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim Rec As New ADODB.Recordset

Private Sub Clear()
    id.Text = ""
    placa.Text = ""
    modelo.Text = ""
    marca.Text = ""
    quantidade_lugares.Text = ""
    descricao.Text = ""
    url.Text = ""
    
    WebBrowser1.Navigate "https://e7.pngegg.com/pngimages/829/733/png-clipart-logo-brand-product-trademark-font-not-found-logo-brand.png"
    
    DuplicarCadastro.Visible = False
    Excluir.Visible = False
    PublicaSite.Visible = False
    RemoverSite.Visible = False
    LabelSiteAcoes.Visible = False
    StatusSite.Visible = False
End Sub

Private Sub listagem_Click()
    MeusVeiculos.Show
End Sub

Private Sub Pesquisar_Click()
    If placa.Text = "" Then
         MsgBox "Digite a Placa do Veiculo para realizar pesquisar!"
    Else
        Con.Open "DSN=MSHOP30;Database=marcos;uid=sa;pwd=DBd4t43xp0rt@;"
        
        Rec.Open "Select * from veiculo where placa='" & placa.Text & "'", Con, adOpenStatic, adLockReadOnly
        
        Call Clear
        
        If Rec.RecordCount > 0 Then
            id.Text = Rec.Fields!id
            placa.Text = Rec.Fields!placa
            modelo.Text = Rec.Fields!modelo
            marca.Text = Rec.Fields!marca
            quantidade_lugares.Text = Rec.Fields!quantidade_lugares
            descricao.Text = Rec.Fields!descricao
            url.Text = Rec.Fields!url_imagem
            
            WebBrowser1.Navigate Rec.Fields!url_imagem
            Excluir.Visible = True
            DuplicarCadastro.Visible = True
            LabelSiteAcoes.Visible = True
            StatusSite.Visible = True
            
            If Rec.Fields!is_site = 0 Or Rec.Fields!is_site = Null Then
                PublicaSite.Visible = True
                StatusSite.Caption = "NÃO FOI PUBLICADO NO SITE"
            Else
                RemoverSite.Visible = True
                StatusSite.Caption = "PUBLICADO NO SITE"
            End If
            
        Else
            MsgBox "Cadastro não encontrado!"
            
            If id.Text <> "" Then
                Call Clear
            End If
        End If
            
        Rec.Close
        Con.Close
    End If
End Sub

Private Sub Excluir_Click()
    If id.Text = "" Then
         MsgBox "Selecione um veiculo para realizar a exclusão!"
    Else
        Con.Open "DSN=MSHOP30;Database=marcos;uid=sa;pwd=DBd4t43xp0rt@;"
     
        Rec.Open "SELECT * FROM veiculo WHERE id = " & id.Text & "", Con, adOpenKeyset, adLockOptimistic
        
        If Not Rec.EOF Then
            MsgBox "Cadastro excluido com  sucesso!"
            
            Rec.Delete
            Rec.Close
            Con.Close
            
            Call Clear
        End If
    End If
End Sub

Private Sub Gravar_Click()
    Dim strsql As String

    If placa.Text = "" Then
        MsgBox "Digite a placa do Veiculo!"
    ElseIf modelo.Text = "" Then
        MsgBox "Digite o modelo do veiculo!"
    ElseIf marca.Text = "" Then
        MsgBox "Digite a marca do veiculo!"
    ElseIf quantidade_lugares.Text = "" Then
        MsgBox "Digite a quantidade de lugares do veiculo!"
    ElseIf descricao.Text = "" Then
        MsgBox "Digite uma descrição do veiculo!"
    ElseIf url.Text = "" Then
        MsgBox "Informe uma url de imagem do veiculo!"
    Else
        Con.Open "DSN=MSHOP30;Database=marcos;uid=sa;pwd=DBd4t43xp0rt@;"
        
        If (id.Text <> "" And DuplicarCadastro.Value = False) Then
            strsql = "Update dbo.veiculo set url_imagem = '" & url.Text & "', placa = '" & placa.Text & "', modelo = '" & modelo.Text & "', marca = '" & marca.Text & "', descricao = '" & descricao.Text & "', quantidade_lugares = '" & quantidade_lugares.Text & "' where id = '" & id.Text & "'"
        
            MsgBox "Cadastro alterado com  sucesso!", vbOKCancel, "Realizado!"
        Else
            strsql = "INSERT INTO dbo.veiculo(url_imagem, placa,modelo,marca,quantidade_lugares, descricao)VALUES('" & url.Text & "', '" & placa.Text & "', '" & modelo.Text & "','" & marca.Text & "', '" & quantidade_lugares.Text & "', '" & descricao.Text & "')"
            
            If DuplicarCadastro.Value = True Then
                MsgBox "Cadastro duplicado com sucesso!"
            Else
                MsgBox "Cadastro registrado com  sucesso!"
            End If
            
        End If
        
        Con.BeginTrans
        Con.Execute strsql
        Con.CommitTrans
        Con.Close
        
        Call Clear
    End If
End Sub


Function StatusVeiculoSite(is_site As Integer, id As Integer) As String
    Con.Open "DSN=MSHOP30;Database=marcos;uid=sa;pwd=DBd4t43xp0rt@"
    Con.BeginTrans
    Con.Execute "Update dbo.veiculo set is_site = " & is_site & "  where id = " & id & ""
    Con.CommitTrans
    Con.Close
End Function

Private Sub PublicaSite_Click()
  Dim result As String
  result = APISITE.RequestSite(id.Text, placa.Text, descricao.Text, url.Text, marca.Text, modelo.Text, quantidade_lugares.Text, 1)
  result = StatusVeiculoSite(1, id.Text)
  Call Clear
End Sub

Private Sub RemoverSite_Click()
    Dim result As String
    result = APISITE.RequestSite(id.Text, "", "", "", "", "", "", 0)
    result = StatusVeiculoSite(0, id.Text)
    Call Clear
End Sub
