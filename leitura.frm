VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{D39AF3D5-81AD-4B7D-9C62-36BFAF895414}#1.0#0"; "systemup00.ocx"
Begin VB.Form fo_ler_arq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leitura do Arquivo de Retorno do Escritório/Finasa"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14685
   Icon            =   "leitura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   14685
   Begin MSComDlg.CommonDialog cx_dialogo 
      Left            =   90
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SystemUp00.CaixaOption op_tipo 
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   510
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   344
      Largura         =   1860
      xColunas        =   4
      xLinhas         =   1
      Nomes           =   "Escritório Itaú;Escritório Unibanco;Escritório Santander;Finasa"
      Valores         =   "1;2;3;4"
      Valor           =   "1"
   End
   Begin SystemUp00.CaixaCombo bo_arquivo 
      Height          =   270
      Left            =   780
      TabIndex        =   1
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   476
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaixaAlta       =   -1  'True
      Obrigatorio     =   -1  'True
      Selecionado     =   -1  'True
      Tamanho         =   ""
      Limpar          =   0   'False
   End
   Begin MSGrid.Grid gr_cons 
      Height          =   7155
      Left            =   60
      TabIndex        =   10
      Top             =   1200
      Width           =   14550
      _Version        =   65536
      _ExtentX        =   25665
      _ExtentY        =   12621
      _StockProps     =   77
      BackColor       =   16777215
      FixedCols       =   0
      HighLight       =   0   'False
   End
   Begin Threed.SSCommand bt_LerArquivo 
      Height          =   315
      Left            =   11070
      TabIndex        =   5
      ToolTipText     =   "Lê o Arquivo"
      Top             =   120
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "&Ler Arquivo"
      ForeColor       =   12582912
      BevelWidth      =   0
   End
   Begin Threed.SSCommand bt_cancelar 
      Height          =   315
      Left            =   13470
      TabIndex        =   9
      ToolTipText     =   "Sair da Tela"
      Top             =   840
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "&Sair"
      ForeColor       =   0
      BevelWidth      =   0
   End
   Begin Threed.SSCommand bt_imprimir 
      Height          =   315
      Left            =   12270
      TabIndex        =   6
      ToolTipText     =   "Impressão do Relatório"
      Top             =   120
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "&Imprimir"
      ForeColor       =   32768
      BevelWidth      =   0
   End
   Begin Threed.SSCheck op_inclui_fase 
      Height          =   195
      Left            =   8250
      TabIndex        =   3
      Top             =   180
      Width           =   2385
      _Version        =   65536
      _ExtentX        =   4207
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Não incluir as fases na leitura"
      ForeColor       =   12582912
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
   Begin Threed.SSCommand bt_baixar 
      Height          =   315
      Left            =   13470
      TabIndex        =   7
      ToolTipText     =   "Baixa em lote"
      Top             =   120
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "&Baixa Lote"
      ForeColor       =   255
      Enabled         =   0   'False
      BevelWidth      =   0
   End
   Begin Threed.SSCheck op_itau_cnab400 
      Height          =   195
      Left            =   8250
      TabIndex        =   4
      Top             =   480
      Width           =   2385
      _Version        =   65536
      _ExtentX        =   4207
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Gerar Arquivo Cnab400 Itaú"
      ForeColor       =   12582912
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
   Begin Threed.SSCommand bt_cnab400 
      Height          =   315
      Left            =   11070
      TabIndex        =   8
      ToolTipText     =   "Gerar arquivo Cnab400 para registro dos boletos no banco ITAÚ"
      Top             =   480
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "Arq cnab400"
      ForeColor       =   255
      Enabled         =   0   'False
      BevelWidth      =   0
   End
   Begin Threed.SSCommand bt_json 
      Height          =   315
      Left            =   12270
      TabIndex        =   11
      ToolTipText     =   "Registrar os boletos ITAÚ diretamente no banco via WebService"
      Top             =   480
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "Reg webServ"
      ForeColor       =   255
      Enabled         =   0   'False
      BevelWidth      =   0
   End
   Begin Threed.SSCheck op_itau_webservice 
      Height          =   195
      Left            =   8250
      TabIndex        =   12
      Top             =   780
      Width           =   2385
      _Version        =   65536
      _ExtentX        =   4207
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Registro Itaú no WebService"
      ForeColor       =   12582912
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivo"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "fo_ler_arq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'definição do registro para leitura
Dim fl_foco As Integer
Dim tpArquivo As String
Dim Tabela As String
Dim Tabela2 As String
Dim registro As String
Dim numarq As Integer
Dim fim As Integer
Dim contador As Long

Dim r_prefcentr As String
Dim r_cgccli As String
Dim r_nomecli As String
Dim r_endcli As String
Dim r_baicli As String
Dim r_cidcli As String
Dim r_estcli As String
Dim r_cepcli As Long
Dim r_dddcli As String
Dim r_fonecli As String
Dim r_fonecli2 As String
Dim r_faxcli As String
Dim r_ctacli As String
Dim r_contato As String
Dim r_cgcdev As String
Dim r_nomedev As String
Dim r_enddev As String
Dim r_baidev As String
Dim r_ciddev As String
Dim r_estdev As String
Dim r_cepdev As Long
Dim r_ddddev As String
Dim r_fonedev As String
Dim r_fonedev2 As String
Dim r_faxdev As String
Dim r_observ As String
Dim r_especie As String
Dim r_titulo As String
Dim r_parcela As String
Dim r_dataemi As Variant
Dim r_datavct As Variant
Dim r_valortit As Currency
Dim r_valorpro As Currency
Dim r_valdivida As Currency
Dim r_valvencid As Currency
Dim r_valminimo As Currency
Dim r_cdidentbanco As String
Dim r_tipofia As String
Dim r_nomefia As String
Dim r_endfia As String
Dim r_baifia As String
Dim r_cidfia As String
Dim r_estadofia As String
Dim r_estfia As String
Dim r_cepfia As Long
Dim r_dddfia As String
Dim r_fonefia As String
Dim r_fonefia2 As String
Dim r_faxfia As String

'definição de outras variáveis
Dim bd_auxParcela As New ADODB.Recordset
Dim mySQL As String
Dim cd_filial As Integer
Dim parcela As String
Dim ErroBaixaParcela As String
Dim ErroBaixaDesconto As String
Dim auxTotErros As Currency

Sub Limpa_Grid()
gr_cons.Rows = 2
pp.LimpaGrid gr_cons
End Sub

Sub Formata_Grid()
'define tamanho do grid
gr_cons.Rows = 2
gr_cons.Cols = 10

'define largura das colunas
gr_cons.ColWidth(0) = 1 '1100
gr_cons.ColWidth(1) = 1200
gr_cons.ColWidth(2) = 1200
gr_cons.ColWidth(3) = 1200
gr_cons.ColWidth(4) = 1200
gr_cons.ColWidth(5) = 1200
gr_cons.ColWidth(6) = 1950
gr_cons.ColWidth(7) = 1
gr_cons.ColWidth(8) = 5250
gr_cons.ColWidth(9) = 900

'define posicionamento do conteudo
gr_cons.ColAlignment(0) = 1
gr_cons.ColAlignment(1) = 1
gr_cons.ColAlignment(2) = 1
gr_cons.ColAlignment(3) = 1
gr_cons.ColAlignment(4) = 2
gr_cons.ColAlignment(5) = 2
gr_cons.ColAlignment(6) = 0
gr_cons.ColAlignment(7) = 0
gr_cons.ColAlignment(8) = 0
gr_cons.ColAlignment(9) = 0

gr_cons.FixedAlignment(0) = 2
gr_cons.FixedAlignment(1) = 2
gr_cons.FixedAlignment(2) = 2
gr_cons.FixedAlignment(3) = 2
gr_cons.FixedAlignment(4) = 2
gr_cons.FixedAlignment(5) = 2
gr_cons.FixedAlignment(6) = 2
gr_cons.FixedAlignment(7) = 2
gr_cons.FixedAlignment(8) = 2
gr_cons.FixedAlignment(9) = 2

'define titulo das colunas
pp.PoeGrid gr_cons, 0, 0, "Cod. Boleto"
pp.PoeGrid gr_cons, 1, 0, "Boleto"
pp.PoeGrid gr_cons, 2, 0, "Vl.Boleto"
pp.PoeGrid gr_cons, 3, 0, "Vl.Pago"
pp.PoeGrid gr_cons, 4, 0, "Dt.Vcto."
pp.PoeGrid gr_cons, 5, 0, "Dt.Pgto."
pp.PoeGrid gr_cons, 6, 0, "Status"
pp.PoeGrid gr_cons, 7, 0, "Cliente"
pp.PoeGrid gr_cons, 8, 0, "Devedor"
pp.PoeGrid gr_cons, 9, 0, "Cheque?"
End Sub

Private Sub bo_arquivo_ClickBt()
pp.relogio 1
cx_dialogo.filename = ""
cx_dialogo.FilterIndex = 1
cx_dialogo.Filter = "Arquivo Escritório Itaú (*.ret)|*.ret|Arquivo Escritório Unibanco (*.00)|*.00;*.01;*.02;*.03;*.04;*.05;*.06|Arquivo Finasa (*.ret)|*.ret|Arquivo Escritório Santander (*.txt)|*.txt|"
cx_dialogo.Flags = &H280900
cx_dialogo.Action = 1
If Trim$(cx_dialogo.filename) <> "" Then
   bo_arquivo.text = cx_dialogo.filename
End If
pp.relogio 0
End Sub

Private Sub bo_arquivo_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub bt_cnab400_Click()
Dim bd_aux As New ADODB.Recordset
Dim bd_end As New ADODB.Recordset
Dim Reg As String
Dim vlAux As Currency
Dim numarq As Integer
Dim contador As Integer
Dim nomeArquivo As String
Dim cdSeq As Double
Dim endereco As String
Dim cep As Double

'NOVA FORMA DE REGISTRAR OS BOLETOS NO SITE DO BANCO

If MsgPergunta("Você confirma a geração do arquivo CNAB400 para registro dos boletos no Itaú?") = 7 Then
   Exit Sub
End If

On Error GoTo Erros
relogio 1

Reg = ""
vlAux = 0

cn.BeginTrans

nomeArquivo = "ITAU_REM_" & Format(Now, "yyyyMMdd_hhmm") & ".TXT"

ChDir glb_diretorio

If Right(glb_diretorio, 1) = "\" Then
   nomeArquivo = glb_diretorio & nomeArquivo
Else
   nomeArquivo = glb_diretorio & "\" & nomeArquivo
End If

numarq = FreeFile
Open nomeArquivo For Output As numarq

'HEADER
Reg = 0
Reg = Reg & 1
Reg = Reg & "REMESSA"
Reg = Reg & "01"
Reg = Reg & "COBRANCA       "
Reg = Reg & "8842"  'Agência
Reg = Reg & "00"    'Complemento
Reg = Reg & "01585" 'Conta
Reg = Reg & "4"     'Dígito
Reg = Reg & Space(8)
Reg = Reg & Format("SCHULZE ADVOGADOS ASSOCIADOS", "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
Reg = Reg & "341"
Reg = Reg & Format("BANCO ITAU SA", "!@@@@@@@@@@@@@@@")
Reg = Reg & Format(Now, "ddMMyy")
Reg = Reg & Space(294)
Reg = Reg & "000001" 'sequencial - nr da linha no arquivo
Print #numarq, Reg

'REGISTRO DETALHE
mySQL = "SELECT cd_processo, cd_cliente, cd_devedor, cd_boleto, dt_cadastro, dt_limite, vl_boleto, vl_desconto, cd_nosso_numero,"
mySQL = mySQL & " Cobranca.dbo.MostraCliente(cd_cliente) AS CLIENTE"
mySQL = mySQL & " FROM BoletoEscritorioItau"
mySQL = mySQL & " WHERE cast(dt_limite as date)>='" & Format(Now, "MM/dd/yyyy") & "'"
mySQL = mySQL & " AND (dt_pagamento='01/01/1900' or dt_pagamento is null)"
mySQL = mySQL & " AND (fl_envio_cnab=0 or fl_envio_cnab IS NULL)"
If bd_aux.State = adStateOpen Then bd_aux.Close
Set bd_aux = cn.Execute(mySQL, 3)
contador = 1
While Not bd_aux.EOF
  
   contador = contador + 1
   
   Reg = "1"
   Reg = Reg & "02"    'indicará o cnpj do beneficiário
   Reg = Reg & "81144396000142" 'nr cnpj do beneficiário
   Reg = Reg & "8842"  'Agência
   Reg = Reg & "00"    'Complemento
   Reg = Reg & "01585" 'Conta
   Reg = Reg & "4"     'Dígito
   
   'Instrução/Alegação cancelada. Se no campo CÓD. DE OCORRÊNCIA for informado os valores:
   '35 – Cancelamento de Instrução ou 38 – beneficiário não concorda com alegação do pagador (posição 109 e 110)
   'Deverá ser informado o cód de Instrução/Alegação cancelada, senão deverá preencher com zeros
   Reg = Reg & "0000"
   Reg = Reg & Space(25) 'Texto livre para uso da empresa
   Reg = Reg & Mid(Trim("" & bd_aux("cd_nosso_numero")), 5, 8) 'Nosso número - identificação do título no banco
   Reg = Reg & "0000000000000" 'quantidade de moeda - preencher com zeros quando a moeda for o real
   Reg = Reg & "109"     'carteira - DIRETA ELETRÔNICA SEM EMISSÃO – SIMPLES
   Reg = Reg & Space(21) 'Uso do banco
   Reg = Reg & "I"       'código da carteira
   Reg = Reg & "01"      'código de ocorrência - vamos fazer só 01-remessa
   Reg = Reg & Format(CDbl(0 & bd_aux("cd_boleto")), "0000000000") 'documento, código do boleto
   Reg = Reg & Format(MsDt("" & bd_aux("dt_limite")), "ddMMyy") 'data do vencimento "ddMMaa"
   vlAux = CCur(0 & pp.MsVl("" & bd_aux("vl_boleto")))
   Reg = Reg & Format(Left$(Fix(vlAux) & Right(Format(CCur(vlAux) - Fix(vlAux), "#,##0.00"), 2), 13), "0000000000000")
   Reg = Reg & "341"     'nr do banco
   Reg = Reg & "00000"   'Agência cobradora - será definido pelo banco
   Reg = Reg & "99"      'espécie do título - 99 diversos
   Reg = Reg & "N"       '(A)ceito, (N)ão aceito
   Reg = Reg & Format(MsDt("" & bd_aux("dt_cadastro")), "ddMMyy") 'data de cadastro "ddMMaa"
   Reg = Reg & "05"      'instruções de cobrança 1 - 05 Receber conforme instruções no própio título
   Reg = Reg & "39"      'instruções de cobrança 2 - 39 não receber após o vencimento
   Reg = Reg & "0000000000000" 'mora por dia de atraso
   Reg = Reg & Format(MsDt("" & bd_aux("dt_limite")), "ddMMyy") 'data limite para concessão do desconto "ddMMaa"
   vlAux = CCur(0 & pp.MsVl("" & bd_aux("vl_desconto")))
   Reg = Reg & Format(Left$(Fix(vlAux) & Right(Format(CCur(vlAux) - Fix(vlAux), "#,##0.00"), 2), 13), "0000000000000")
   Reg = Reg & "0000000000000" 'valor do IOF - somente para seguradoras
   Reg = Reg & "0000000000000" 'valor do abatimento
   
   'Identificação do pagador
   mySQL = "SELECT tp_pessoa, no_cliente, nr_cpf, nr_cgc,"
   mySQL = mySQL & " cobranca.dbo.MostraEnderecoCodigo(cd_cliente,1) AS ENDERECO,"
   mySQL = mySQL & " cobranca.dbo.MostraEnderecoBairroCodigo(cd_cliente,1) AS BAIRRO,"
   mySQL = mySQL & " cobranca.dbo.MostraEnderecoCidadeCodigo(cd_cliente,1) AS CIDADE,"
   mySQL = mySQL & " cobranca.dbo.MostraEnderecoestadoCodigo(cd_cliente,1) AS UF,"
   mySQL = mySQL & " cobranca.dbo.MostraEnderecocepcliente(cd_cliente,1) AS CEP"
   mySQL = mySQL & " FROM cliente"
   mySQL = mySQL & " WHERE cd_cliente=" & CDbl(0 & bd_aux("cd_devedor"))
   If bd_end.State = adStateOpen Then bd_end.Close
   Set bd_end = cn.Execute(mySQL, 3)
   If Not bd_end.EOF Then
      If UCase(Trim("" & bd_end("tp_pessoa"))) = "F" Then
         Reg = Reg & "01" '1 = cpf, 2 = cnpj
         Reg = Reg & Format(CDbl(0 & bd_end("nr_cpf")), "00000000000000")
      Else
         Reg = Reg & "02" '1 = cpf, 2 = cnpj
         Reg = Reg & Format(CDbl(0 & bd_end("nr_cgc")), "00000000000000")
      End If
      Reg = Reg & Format(Left$(Trim("" & bd_end("no_cliente")), 40), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
      Reg = Reg & Format(Left(Trim("" & bd_end("ENDERECO")), 40), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
      Reg = Reg & Format(Left(Trim("" & bd_end("BAIRRO")), 12), "!@@@@@@@@@@@@")
      Reg = Reg & Format(CDbl(0 & bd_end("CEP")), "00000000")
      Reg = Reg & Format(Left(Trim("" & bd_end("CIDADE")), 15), "!@@@@@@@@@@@@@@@")
      Reg = Reg & Format(Trim("" & bd_end("UF")), "!@@")
   Else
      MsgProblema "Informações sobre o devedor não encontradas, vai dar erro na montagem do arquivo. Boleto: " & CDbl(0 & bd_aux("cd_boleto"))
   End If
   bd_end.Close
   
   Reg = Reg & Format(Left(Trim("" & bd_aux("CLIENTE")), 30), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") 'Sacador/Avalista
   Reg = Reg & Space(4) 'brancos
   Reg = Reg & Format(MsDt("" & bd_aux("dt_limite")), "ddMMyy") 'data de mora "ddMMaa"
   Reg = Reg & "00"     'prazo para receber após o vencimento
   Reg = Reg & Space(1) 'brancos
   Reg = Reg & Format(contador, "000000") 'sequência do registro no arquivo
   
   Print #numarq, Reg
   
   mySQL = "UPDATE BoletoEscritorioItau SET fl_envio_cnab=1"
   mySQL = mySQL & " WHERE cd_boleto=" & CDbl(0 & bd_aux("cd_boleto"))
   mySQL = mySQL & " AND cd_processo=" & CDbl(0 & bd_aux("cd_processo"))
   cn.Execute mySQL, , adExecuteNoRecords
   
   bd_aux.MoveNext
Wend

'TRAILLER
Reg = "9"
Reg = Reg & Space(393)
Reg = Reg & Format(contador + 1, "000000")

Print #numarq, Reg

Close #numarq
cn.CommitTrans

MsgAviso "Arquivo Gerado com Sucesso:" & vbCrLf & vbCrLf & nomeArquivo & vbCrLf & vbCrLf & "Esse arquivo deve ser enviado ao banco Itaú."

On Error GoTo 0
relogio 0
Exit Sub
Resume

Erros:
erro Err
Close #numarq
If Right(glb_diretorio, 1) = "\" Then
   Kill (glb_diretorio & nomeArquivo)
Else
   Kill (glb_diretorio & "\" & nomeArquivo)
End If
On Error GoTo 0
cn.RollbackTrans
relogio 0
End Sub

Private Sub bt_baixar_Click()
Dim bd_aux As New ADODB.Recordset
Dim bd_aux2 As New ADODB.Recordset
Dim bd_aux3 As New ADODB.Recordset
Dim bd_cobrador As New ADODB.Recordset
Dim cd_criacao As String
Dim dt_criacao As String
Dim cd_alteracao As String
Dim Baixou As String
Dim JaBaixado As String
Dim BaixaManual As String
Dim NaoEncontrou As String
Dim msgEmail As String
Dim auxVlDespesa As String
Dim auxTotManual As Currency
Dim i As Integer

If gr_cons.Rows = 2 Then
   MsgInformacao "Selecione o arquivo e os boletos para baixa..."
   Exit Sub
End If

Baixou = ""
JaBaixado = ""
BaixaManual = ""
NaoEncontrou = ""
ErroBaixaParcela = ""
ErroBaixaDesconto = ""
auxTotErros = 0
cd_filial = 0

If MsgPergunta("Deseja baixar as parcelas ?") = vbNo Then Exit Sub

On Error GoTo Erros
pp.relogio 1

For i = 1 To gr_cons.Rows - 1
   If CDbl(0 & pp.PegaGrid(gr_cons, 0, i)) <> 0 Then
      'Ignora os boletos sem data de pagamento
      If Trim(pp.PegaGrid(gr_cons, 5, i)) = "/  /" Then
         GoTo pula
      End If
      
      'Verifica se o boleto existe
      mySQL = "SELECT * FROM BoletoFinasa WHERE cd_boleto=" & CDbl(0 & pp.PegaGrid(gr_cons, 1, i))
      Set bd_aux = cn.Execute(mySQL, 3)
      If bd_aux.EOF Then
         NaoEncontrou = NaoEncontrou & vbCrLf & "Boleto: " & Trim(pp.PegaGrid(gr_cons, 1, i))
         bd_aux.Close
         GoTo pula
      End If
      bd_aux.Close
      
      'Consulta as parcelas do boleto
      mySQL = "SELECT cd_parcela, cd_processo, cd_titulo, vl_titulo, vl_total, vl_outras,"
      mySQL = mySQL & " vl_multa, vl_juros, vl_honorarios, vl_despesa, vl_protesto"
      mySQL = mySQL & " FROM CalculoBoletoFinasa"
      mySQL = mySQL & " WHERE cd_boleto=" & CDbl(0 & pp.PegaGrid(gr_cons, 1, i))
      mySQL = mySQL & " ORDER BY cd_parcela"
      Set bd_auxParcela = cn.Execute(mySQL, 3)
      If Not bd_auxParcela.EOF Then
         If CCur(0 & bd_auxParcela("vl_protesto")) <> 0 Or CCur(0 & bd_auxParcela("vl_despesa")) <> 0 Or CCur(0 & bd_auxParcela("vl_outras")) <> 0 Then
            auxVlDespesa = " - Despesas: R$ " & CCur(0 & bd_auxParcela("vl_protesto")) + CCur(0 & bd_auxParcela("vl_outras")) + CCur(0 & bd_auxParcela("vl_despesa"))
         Else
            auxVlDespesa = " - Despesas: R$ 0,00"
         End If
         While Not bd_auxParcela.EOF
            mySQL = "SELECT SUM(vl_total) AS VALOR"
            mySQL = mySQL & " FROM CalculoBoletoFinasa"
            mySQL = mySQL & " WHERE cd_boleto=" & CDbl(0 & pp.PegaGrid(gr_cons, 1, i))
            Set bd_aux3 = cn.Execute(mySQL, 3)
            If Format(pp.MsVl("" & bd_aux3("VALOR")), "#,##0.00") = Trim("" & pp.PegaGrid(gr_cons, 2, i)) Then
               parcela = Format(Trim("" & bd_auxParcela("cd_parcela")), "000")
               cd_criacao = ""
               dt_criacao = ""
               cd_alteracao = ""
               cd_filial = 1
               
               mySQL = "SELECT cd_processo, cd_cliente, cd_devedor, cd_criacao, dt_criacao, cd_alteracao, cd_filial"
               mySQL = mySQL & " FROM BoletoFinasa"
               mySQL = mySQL & " WHERE cd_boleto=" & CDbl(0 & pp.PegaGrid(gr_cons, 1, i))
               Set bd_aux = cn.Execute(mySQL, 3)
               If Not bd_aux.EOF Then
                  cd_criacao = Trim("" & bd_aux("cd_criacao"))
                  If Trim(pp.PegaGrid(gr_cons, 5, i)) <> "/  /" Then
                     dt_criacao = pp.GrDtString(glb_tipo_banco, Trim(pp.PegaGrid(gr_cons, 5, i)))
                  Else
                     dt_criacao = pp.GrDtString(glb_tipo_banco, Format(Now, "dd/MM/yyyy"))
                 End If
              
                  cd_alteracao = IIf(Trim("" & bd_aux("cd_alteracao")) = "", Trim("" & bd_aux("cd_criacao")), Trim("" & bd_aux("cd_alteracao")))
                  cd_filial = CDbl(0 & bd_aux("cd_filial"))
              
                  'Consulta o cobrador do contrato no dia de cadastro
                  mySQL = "SELECT " & cn.DefaultDatabase & ".dbo.MostraCobradorFicha(" & CDbl(0 & bd_aux("cd_processo")) & "," & dt_criacao & ") AS Cobrador"
                  Set bd_cobrador = cn.Execute(mySQL, 3)
                  If Not bd_cobrador.EOF Then
                     If Trim("" & bd_cobrador("Cobrador")) <> "" Then
                        cd_criacao = Trim("" & bd_cobrador("Cobrador"))
                        mySQL = "SELECT cd_usuario AS Cobrador FROM Cliente WHERE cd_cliente=" & CDbl(0 & bd_cobrador("Cobrador"))
                        bd_cobrador.Close
                        Set bd_cobrador = cn.Execute(mySQL, 3)
                        If Not bd_cobrador.EOF Then
                           If Trim("" & bd_cobrador("Cobrador")) <> "" Then
                              cd_criacao = Trim("" & bd_cobrador("Cobrador"))
                           End If
                        End If
                     End If
                  End If
                  bd_cobrador.Close
                  mySQL = "SELECT " & cn.DefaultDatabase & ".dbo.MostraCodFilialUsuario('" & cd_criacao & "') AS Filial"
                  Set bd_cobrador = cn.Execute(mySQL, 3)
                  If Not bd_cobrador.EOF Then
                     If Trim("" & bd_cobrador("Filial")) <> "" Then
                        cd_filial = CDbl(0 & bd_cobrador("Filial"))
                     End If
                  End If
                  bd_cobrador.Close
                  
                  mySQL = "SELECT vl_titulo FROM CobrancaParcial"
                  mySQL = mySQL & " WHERE cd_processo=" & CDbl(0 & bd_auxParcela("cd_processo"))
                  mySQL = mySQL & " AND cd_titulo='" & Trim(bd_auxParcela("cd_titulo")) & "'"
                  If Len("" & parcela) = 1 Then
                     mySQL = mySQL & " AND (cd_parcela='" & parcela & "'"
                     mySQL = mySQL & " OR cd_parcela='" & Format(parcela, "00") & "'"
                     mySQL = mySQL & " OR cd_parcela='" & Format(parcela, "000") & "')"
                  ElseIf Len("" & parcela) = 2 Then
                     mySQL = mySQL & " AND (cd_parcela='" & parcela & "'"
                     mySQL = mySQL & " OR cd_parcela='" & Format(parcela, "000") & "')"
                  Else
                     mySQL = mySQL & " AND cd_parcela='" & parcela & "'"
                  End If
                  Set bd_aux2 = cn.Execute(mySQL, 3)
                  If Not bd_aux2.EOF Then
                     JaBaixado = JaBaixado & vbCrLf & "Boleto: " & Trim(pp.PegaGrid(gr_cons, 1, i)) & " - Ficha: " & CDbl(0 & bd_auxParcela("cd_processo")) & " - Título: " & pp.GrTx(bd_auxParcela("cd_titulo")) & " - Parcela: " & parcela & auxVlDespesa
                  Else
                     If Baixar_Parcela(i) = True Then
                        Baixou = Baixou & vbCrLf & "Boleto: " & Trim(pp.PegaGrid(gr_cons, 1, i)) & " - Ficha: " & CDbl(0 & bd_auxParcela("cd_processo")) & " - Título: " & pp.GrTx(bd_auxParcela("cd_titulo")) & " - Parcela: " & parcela & auxVlDespesa
                     End If
                  End If
                  bd_aux2.Close
               End If
               bd_aux.Close
            Else
               BaixaManual = BaixaManual & vbCrLf & "Boleto: " & Trim(pp.PegaGrid(gr_cons, 1, i)) & " - Valor: " & Trim("" & pp.PegaGrid(gr_cons, 2, i)) & " - Ficha: " & CDbl(0 & bd_auxParcela("cd_processo")) & " - Título: " & pp.GrTx(bd_auxParcela("cd_titulo")) & " - Parcela: " & parcela
               auxTotManual = auxTotManual + CDbl(0 & pp.PegaGrid(gr_cons, 2, i))
            End If
            bd_aux3.Close
            bd_auxParcela.MoveNext
         Wend
      Else
         BaixaManual = BaixaManual & vbCrLf & "Boleto: " & Trim(pp.PegaGrid(gr_cons, 1, i)) & " - Valor: " & Trim("" & pp.PegaGrid(gr_cons, 2, i))
         auxTotManual = auxTotManual + CDbl(0 & pp.PegaGrid(gr_cons, 2, i))
      End If
      bd_auxParcela.Close
   End If
pula:
Next i

msgEmail = "Processo de baixa automática para Bradesco/Finasa" & vbCrLf & vbCrLf & "Arquivo: " & bo_arquivo.text & vbCrLf
If Baixou <> "" Then
   msgEmail = msgEmail & vbCrLf & "Boletos baixados:"
   msgEmail = msgEmail & vbCrLf & "-----------------" & Baixou & vbCrLf
End If
If JaBaixado <> "" Then
   msgEmail = msgEmail & vbCrLf & "Boletos que já estavam baixados:"
   msgEmail = msgEmail & vbCrLf & "--------------------------------" & JaBaixado & vbCrLf
End If
If NaoEncontrou <> "" Then
   msgEmail = msgEmail & vbCrLf & "ERRO - Boletos que não foram encontrados na nossa base de dados:"
   msgEmail = msgEmail & vbCrLf & "----------------------------------------------------------------" & NaoEncontrou & vbCrLf
End If
If Baixou = "" And JaBaixado = "" And NaoEncontrou = "" Then
   msgEmail = msgEmail & vbCrLf & "Nenhum boleto baixado neste processo..."
End If

enviaEmail "Baixa Automática Bradesco [ Baixados ]", ("" & strSql(msgEmail, False)), "repasse@schulze.com.br;honorarios@schulze.com.br", , , "TEXT"

If BaixaManual <> "" Or ErroBaixaParcela <> "" Or ErroBaixaDesconto <> "" Then
   msgEmail = "Processo de baixa automática para Bradesco/Finasa" & vbCrLf & vbCrLf & "Arquivo: " & bo_arquivo.text & vbCrLf
   If BaixaManual <> "" Then
      msgEmail = msgEmail & vbCrLf & "Boletos que não foram baixados - Efetue a baixa manualmente:"
      msgEmail = msgEmail & vbCrLf & "------------------------------------------------------------" & BaixaManual & vbCrLf
      msgEmail = msgEmail & "Valor total dos boletos: R$ " & Format(auxTotManual, "#,##0.00") & vbCrLf
   End If
   If ErroBaixaParcela <> "" Or ErroBaixaDesconto <> "" Then
      msgEmail = msgEmail & vbCrLf & "ERRO na baixa das parcelas - Efetue a baixa manualmente:"
      msgEmail = msgEmail & vbCrLf & "--------------------------------------------------------" & ErroBaixaParcela & vbCrLf
      If ErroBaixaDesconto <> "" Then
         msgEmail = msgEmail & vbCrLf & "Boletos com ERRO no valor do desconto:" & ErroBaixaDesconto & vbCrLf
      End If
      msgEmail = msgEmail & vbCrLf & "Valor total dos boletos: R$ " & Format(auxTotErros, "#,##0.00") & vbCrLf
   End If
   
   enviaEmail "Baixa Automática Bradesco [ Não Baixados ]", ("" & strSql(msgEmail, False)), "repasse@schulze.com.br;honorarios@schulze.com.br", , , "TEXT"
End If

MsgAviso "Processo de baixa encerrado !!!"

On Error GoTo 0
pp.relogio 0
Exit Sub

Erros:
erro Err
On Error GoTo 0
pp.relogio 0
End Sub

Function Baixar_Parcela(ByVal plin As Integer) As Boolean
Dim repasse As Currency
Dim captador As Currency
Dim total As Currency
Dim auxDesconto As Currency
Dim auxDescontoRepasse As Currency
Dim auxVLMulta As Currency
Dim bd_aux As New ADODB.Recordset

On Error GoTo Erro_Baixa

Baixar_Parcela = False

auxVLMulta = 0
         
'Confere a data de emissão do título: se >= "05/05/2014" cobrar multa, senão não cobrar multa
mySQL = "SELECT dt_emissao FROM CobrancaTitulo"
mySQL = mySQL & " WHERE cd_processo=" & CDbl(0 & bd_auxParcela("cd_processo"))
mySQL = mySQL & " AND cd_titulo='" & pp.GrTx(bd_auxParcela("cd_titulo")) & "'"
mySQL = mySQL & " AND cd_parcela='" & pp.GrTx("" & parcela) & "'"
mySQL = mySQL & " AND dt_emissao<>'01/01/1900' AND dt_emissao IS NOT NULL"
Set bd_aux = cn.Execute(mySQL, 3)
If Not bd_aux.EOF Then
   If CDate(Format("" & bd_aux("dt_emissao"), "dd/MM/yyyy")) >= CDate("05/05/2014") Then
      auxVLMulta = CCur(0 & bd_auxParcela("vl_multa"))
   Else
      auxVLMulta = 0
   End If
End If

repasse = CCur(0 & bd_auxParcela("vl_total")) - CCur(0 & bd_auxParcela("vl_honorarios")) - CCur(0 & bd_auxParcela("vl_despesa"))
captador = 0
total = CCur(0 & bd_auxParcela("vl_titulo")) + CCur(0 & bd_auxParcela("vl_protesto")) + CCur(0 & bd_auxParcela("vl_outras")) + CCur(0 & bd_auxParcela("vl_juros")) + auxVLMulta + CCur(0 & bd_auxParcela("vl_despesa"))

auxDesconto = CCur(0 & bd_auxParcela("vl_titulo")) + CCur(0 & bd_auxParcela("vl_honorarios")) + CCur(0 & bd_auxParcela("vl_protesto")) + CCur(0 & bd_auxParcela("vl_outras")) + CCur(0 & bd_auxParcela("vl_juros")) + auxVLMulta + CCur(0 & bd_auxParcela("vl_despesa"))
auxDesconto = CCur(0 & auxDesconto) - CCur(0 & bd_auxParcela("vl_total"))

If auxDesconto < 0 Then
   GoTo Erro_Desconto
End If

mySQL = "INSERT INTO CobrancaParcial ("
mySQL = mySQL & "cd_filial,"
mySQL = mySQL & "cd_filial_pgto,"
mySQL = mySQL & "cd_processo,"
mySQL = mySQL & "cd_titulo,"
mySQL = mySQL & "cd_parcela,"
mySQL = mySQL & "dt_pagamento,"
mySQL = mySQL & "dt_prestacao_contas,"
mySQL = mySQL & "vl_pago,"
mySQL = mySQL & "dt_cheque,"
mySQL = mySQL & "nr_cheque,"
mySQL = mySQL & "dt_compensacao,"
mySQL = mySQL & "de_baixa,"
mySQL = mySQL & "ic_direto,"
mySQL = mySQL & "vl_repasse,"
mySQL = mySQL & "vl_retido,"
mySQL = mySQL & "vl_hd,"
mySQL = mySQL & "vl_cobrador,"
mySQL = mySQL & "vl_captador,"
mySQL = mySQL & "vl_multa,"
mySQL = mySQL & "vl_juros,"
mySQL = mySQL & "vl_protesto,"
mySQL = mySQL & "vl_titulo,"
mySQL = mySQL & "vl_credito,"
mySQL = mySQL & "vl_multa_contratual,"
mySQL = mySQL & "vl_comissao_permanencia,"
mySQL = mySQL & "vl_custas,"
mySQL = mySQL & "vl_desconto,"
mySQL = mySQL & "vl_despesas,"
mySQL = mySQL & "vl_adiantado,"
mySQL = mySQL & "cd_banco,"
mySQL = mySQL & "fl_quitado,"
mySQL = mySQL & "cd_criacao,"
mySQL = mySQL & "dt_criacao,"
mySQL = mySQL & "tp_boleto,"
mySQL = mySQL & "cd_boleto,"
mySQL = mySQL & "cd_desconto"
mySQL = mySQL & ") VALUES ("
mySQL = mySQL & glb_filial & ","
mySQL = mySQL & CDbl(0 & cd_filial) & ","
mySQL = mySQL & CDbl(0 & bd_auxParcela("cd_processo")) & ","
mySQL = mySQL & "'" & pp.GrTx(bd_auxParcela("cd_titulo")) & "',"
mySQL = mySQL & "'" & pp.GrTx("" & parcela) & "',"
mySQL = mySQL & pp.GrDtString(glb_tipo_banco, pp.PegaGrid(gr_cons, 5, plin)) & ","
mySQL = mySQL & "'01/01/1900',"
mySQL = mySQL & pp.GrVl(0 & bd_auxParcela("vl_total")) & ","
mySQL = mySQL & "'01/01/1900',"
mySQL = mySQL & "'',"
mySQL = mySQL & "'01/01/1900',"
mySQL = mySQL & "'',"
mySQL = mySQL & "0,"
mySQL = mySQL & pp.GrVl(0 & repasse) & ","
mySQL = mySQL & pp.GrVl(0 & bd_auxParcela("vl_honorarios")) & ","
mySQL = mySQL & "0,"
mySQL = mySQL & "0,"
mySQL = mySQL & pp.GrVl(0 & captador) & ","
mySQL = mySQL & pp.GrVl(auxVLMulta) & ","
'Verificar se o valor do juros for menor que o valor de desconto, colocar zero, senão grava valor do juros menos o desconto
If CCur(0 & bd_auxParcela("vl_juros")) < CCur(0 & auxDesconto) Then
   mySQL = mySQL & "0,"
   auxDescontoRepasse = CCur(0 & auxDesconto) - CCur(0 & bd_auxParcela("vl_juros"))
Else
   mySQL = mySQL & pp.GrVl(CCur(0 & bd_auxParcela("vl_juros")) - CCur(0 & auxDesconto)) & ","
   auxDescontoRepasse = 0
End If
mySQL = mySQL & pp.GrVl(CCur(0 & bd_auxParcela("vl_protesto")) + CCur(0 & bd_auxParcela("vl_outras"))) & ","
mySQL = mySQL & pp.GrVl(0 & bd_auxParcela("vl_titulo")) & ","
mySQL = mySQL & pp.GrVl(0 & total) & ","
mySQL = mySQL & "0,"
mySQL = mySQL & "0,"
mySQL = mySQL & "0,"
mySQL = mySQL & pp.GrVl(0 & auxDescontoRepasse) & ","
mySQL = mySQL & pp.GrVl(0 & bd_auxParcela("vl_despesa")) & ","
mySQL = mySQL & "0,"
mySQL = mySQL & "0,"
mySQL = mySQL & "1,"
mySQL = mySQL & "'" & glb_usuario & "',"
mySQL = mySQL & "GetDate(),"
mySQL = mySQL & "'FINASA',"
mySQL = mySQL & CDbl(0 & pp.PegaGrid(gr_cons, 1, plin)) & ","
mySQL = mySQL & "'')"
cn.Execute mySQL, , adExecuteNoRecords

'coloca indicador de pagamento parcial do processo e atualiza valores
mySQL = "UPDATE CobrancaTitulo SET"
mySQL = mySQL & " dt_baixa=" & pp.GrDtString(glb_tipo_banco, pp.PegaGrid(gr_cons, 5, plin)) & ","
mySQL = mySQL & " dt_inibicao=" & pp.GrDtString(glb_tipo_banco, pp.PegaGrid(gr_cons, 5, plin)) & ","
mySQL = mySQL & " dt_compensacao=" & pp.GrDtString(glb_tipo_banco, pp.PegaGrid(gr_cons, 5, plin)) & ","
mySQL = mySQL & " vl_protesto=" & pp.GrVl(CCur(0 & bd_auxParcela("vl_protesto")) + CCur(0 & bd_auxParcela("vl_outras"))) & ","
mySQL = mySQL & " vl_honorarios_adicionais=0,"
mySQL = mySQL & " vl_pago=vl_pago+" & pp.GrVl(0 & bd_auxParcela("vl_total")) & ","
mySQL = mySQL & " cd_motivo=7,"
mySQL = mySQL & " vl_saldo=0,"
mySQL = mySQL & " cd_alteracao='" & glb_usuario & "',"
mySQL = mySQL & " dt_alteracao=getdate()"
mySQL = mySQL & " WHERE cd_processo=" & CDbl(0 & bd_auxParcela("cd_processo"))
mySQL = mySQL & " AND cd_titulo='" & pp.GrTx(bd_auxParcela("cd_titulo")) & "'"
mySQL = mySQL & " AND cd_parcela='" & pp.GrTx("" & parcela) & "'"
cn.Execute mySQL, , adExecuteNoRecords

'Grava a data de pagamento no boleto finasa
mySQL = "UPDATE BoletoFinasa SET"
mySQL = mySQL & " dt_pagamento=" & pp.GrDtString(glb_tipo_banco, pp.PegaGrid(gr_cons, 5, plin))
mySQL = mySQL & " WHERE cd_boleto=" & CDbl(0 & pp.PegaGrid(gr_cons, 1, plin))
cn.Execute mySQL, , adExecuteNoRecords

Baixar_Parcela = True
On Error GoTo 0
Exit Function

Erro_Desconto:
ErroBaixaDesconto = ErroBaixaDesconto & vbCrLf & "Boleto: " & Trim(pp.PegaGrid(gr_cons, 1, plin)) & " - Valor: " & Trim("" & pp.PegaGrid(gr_cons, 2, plin)) & " - Ficha: " & CDbl(0 & bd_auxParcela("cd_processo")) & " - Título: " & pp.GrTx(bd_auxParcela("cd_titulo")) & " - Parcela: " & parcela & " - Desconto: " & auxDesconto
auxTotErros = auxTotErros + CDbl(0 & pp.PegaGrid(gr_cons, 2, plin))
On Error GoTo 0
Exit Function

Erro_Baixa:
ErroBaixaParcela = ErroBaixaParcela & vbCrLf & "Boleto: " & Trim(pp.PegaGrid(gr_cons, 1, plin)) & " - Valor: " & Trim("" & pp.PegaGrid(gr_cons, 2, plin)) & " - Ficha: " & CDbl(0 & bd_auxParcela("cd_processo")) & " - Título: " & pp.GrTx(bd_auxParcela("cd_titulo")) & " - Parcela: " & parcela
ErroBaixaParcela = ErroBaixaParcela & vbCrLf & Error$(Err)
auxTotErros = auxTotErros + CDbl(0 & pp.PegaGrid(gr_cons, 2, plin))
erro Err
On Error GoTo 0
End Function

Private Sub bt_cancelar_Click()
Unload Me
End Sub

Private Sub bt_imprimir_Click()
Dim ct As Integer
Dim Contpag As Integer
Dim bd_aux As New ADODB.Recordset
Dim i As Integer

fo_printer.Show 1
If glb_imprimir <> 1 Then
   glb_imprimir = 0
   Exit Sub
End If
glb_imprimir = 0

pp.relogio 1
On Error GoTo Erros

ct = 0
Contpag = 0
ConfiguraImpressora
Printer.Orientation = 2

For i = 1 To gr_cons.Rows - 4
   
   'Faz quebra de página e imprime o cabeçalho
   If ct > 180 Or ct = 0 Then
      If ct <> 0 Then
         Printer.NewPage
      End If
      Contpag = Contpag + 1
      Printer.FontSize = 12
      
      ct = 10
      Imprime 5, ct, "" & emp_nome
      Printer.FontSize = 8
      Imprime 263, ct, "Pag.: " & Format(Contpag, "000")
      
      ct = ct + 6
      Imprime 5, ct, "Emissão: " & Now
      Printer.FontSize = 14
      If op_tipo.valor = 1 Then
         Imprime 108, ct, "Arquivo de Retorno do Escritório - Itaú"
      ElseIf op_tipo.valor = 2 Then
         Imprime 108, ct, "Arquivo de Retorno do Escritório - Unibanco"
      ElseIf op_tipo.valor = 3 Then
         Imprime 108, ct, "Arquivo de Retorno do Escritório - Santander"
      Else
         Imprime 108, ct, "Arquivo de Retorno do Finasa"
      End If
      Printer.FontSize = 7
           
      ct = ct + 5
      Imprime 5, ct, "Arquivo: " & Trim$(bo_arquivo.text)
      
      Printer.FontSize = 9
      ct = ct + 6
      Imprime 5, ct, "Status"
      Imprime 30, ct, "Código"
      Imprime 53, ct, "Valor Original"
      Imprime 76, ct, "Valor Cobrado"
      Imprime 100, ct, "Dt.Vcto"
      Imprime 116, ct, "Dt.Crédito"
      Imprime 133, ct, "Cliente"
      Imprime 193, ct, "Devedor"
      Imprime 260, ct, "Cheque"
      
      ct = ct + 4
      Printer.Line (5, ct)-(277, ct)
      ct = ct - 2
   End If
   
   'Imprime linhas de detalhe
   Printer.FontSize = 8
   If Trim$(UCase$(PegaGrid(gr_cons, 6, i))) = "TÍTULO PAGO" Then
      If Trim$(PegaGrid(gr_cons, 1, i)) <> "" Then
         ct = ct + 4
         Imprime 5, ct, Trim$(UCase$(PegaGrid(gr_cons, 6, i)))
         Imprime 30, ct, PegaGrid(gr_cons, 1, i)
         Imprime 133, ct, Mid$(PegaGrid(gr_cons, 7, i), 14, 30)
         Imprime 193, ct, Mid$(PegaGrid(gr_cons, 8, i), 14, 40)
         
         Printer.FontName = "Courier New"
         Imprime 50, ct, Format(Format(PegaGrid(gr_cons, 2, i), "#,##0.00"), "@@@@@@@@@@@@@")
         Imprime 75, ct, Format(Format(PegaGrid(gr_cons, 3, i), "#,##0.00"), "@@@@@@@@@@@@@")
         Printer.FontName = "Arial"
   
         Imprime 100, ct, PegaGrid(gr_cons, 4, i)
         Imprime 116, ct, PegaGrid(gr_cons, 5, i)
         Imprime 260, ct, PegaGrid(gr_cons, 9, i)
      End If
   End If
   Printer.FontName = "Arial"
Next i

'Imprime a linha de totais
If Trim$(UCase$(PegaGrid(gr_cons, 1, gr_cons.Rows - 2))) = "TOTAL" Then
   ct = ct + 6
   Imprime 30, ct, PegaGrid(gr_cons, 1, gr_cons.Rows - 2)
   Printer.FontName = "Courier New"
   Imprime 50, ct, Format(Format(PegaGrid(gr_cons, 2, gr_cons.Rows - 2), "#,##0.00"), "@@@@@@@@@@@@@")
   Imprime 75, ct, Format(Format(PegaGrid(gr_cons, 3, gr_cons.Rows - 2), "#,##0.00"), "@@@@@@@@@@@@@")
   Printer.FontName = "Arial"
End If
Printer.Line (5, ct + 5)-(277, ct + 5)
ct = ct + 5
Printer.FontSize = 7
Imprime 205, ct + 2, "System-Up® - http://www.system-up.com.br"

Printer.EndDoc
On Error GoTo 0
pp.relogio 0
Exit Sub

Erros:
erro Err
Printer.EndDoc
On Error GoTo 0
pp.relogio 0
End Sub

Function GetBase64String(ByRef pArray() As Byte) As String
    Dim pDoc As DOMDocument
    Dim pTxt64 As IXMLDOMElement

    Set pDoc = New DOMDocument
    Set pTxt64 = pDoc.createElement("encode")
    
    pTxt64.dataType = "bin.base64"
    pTxt64.nodeTypedValue = pArray
    GetBase64String = pTxt64.text
End Function

Private Sub bt_json_Click()
Dim pUsuario As String
Dim pSenha As String
Dim pArray() As Byte
Dim pAcesso As String
Dim pURL As String
Dim pStatus As Integer
Dim pRetorno As String
Dim pToken As String

If MsgPergunta("Confirma o registro via WebService dos boletos pendentes no sistema do Itaú?") = 7 Then
   Exit Sub
End If

On Error GoTo Erros
relogio 1

'PEGA O TOKEN (VÁLIDO POR 4 HORAS) VIA JSON

'Usuário e senha fornecidos pelo banco
pUsuario = "FA7091xq5SUF0"
pSenha = "NrCI_Rtz9hiJOwxm7DgnkGeWPO0-F3VdNdtitwy2V_JPu75GWgVcfbORrVAq3FQKDhZ6VP4fjn2wG7cdyO7Bjg2"

'Monta o array base 64 - só funciona se o pUsuario e pSenha não tiverem caracteres acentuados
pArray = StrConv(pUsuario & ":" & pSenha, vbFromUnicode)

'Converte o array para uma string de base 64
pAcesso = GetBase64String(pArray)
'Se tiver alguma quebra de linha retira o caracter
If InStr(pAcesso, vbCr) > 0 Then
   pAcesso = Left(pAcesso, InStr(pAcesso, vbCr) - 1) & Mid(pAcesso, InStr(pAcesso, vbCr) + 1)
End If
If InStr(pAcesso, vbLf) > 0 Then
   pAcesso = Left(pAcesso, InStr(pAcesso, vbLf) - 1) & Mid(pAcesso, InStr(pAcesso, vbLf) + 1)
End If

'url para produção
pURL = "https://autorizador-boletos.itau.com.br"
'url homologação
pURL = "https://oauth.itau.com.br/identity/connect/token"
pRetorno = Ler_Token(pURL, pAcesso)

'No retorno deve estar o token
If InStr(pRetorno, "expires_in") > 3 Then
   pToken = Mid(pRetorno, 1, InStr(pRetorno, "expires_in") - 3)
   pToken = Trim(Mid(pToken, InStr(pToken, "access_token") + 14))
   pToken = Replace(pToken, """", "")
End If

If pToken = "" Then
   relogio 0
   MsgProblema "Erro ao gerar o Token de acesso ao Itaú." & vbCrLf & "Tente enviar novamente, se persistir o erro entre em contato com a T.I."
   Exit Sub
End If

'PROCESSO PARA MONTAR A STRING DE ENVIO VIA JSON

Dim bd_aux As New ADODB.Recordset
Dim bd_end As New ADODB.Recordset
Dim Reg As String
Dim vlAux As Currency
Dim cdSeq As Double
Dim endereco As String
Dim cep As Double
Dim pErro As Boolean
Dim NossoNum As String
Dim auxDig As Integer

Reg = ""
vlAux = 0

'LÊ BOLETOS NÃO ENVIADOS AINDA
mySQL = "SELECT cd_processo, cd_cliente, cd_devedor, cd_boleto, dt_cadastro, dt_limite, vl_boleto, vl_desconto,"
mySQL = mySQL & " cd_nosso_numero, cd_codigo_barras, Cobranca.dbo.MostraCliente(cd_cliente) AS CLIENTE"
mySQL = mySQL & " FROM BoletoEscritorioItau"
mySQL = mySQL & " WHERE CAST(dt_limite AS DATE)>='" & Format(Now, "MM/dd/yyyy") & "'"
mySQL = mySQL & " AND (dt_pagamento='01/01/1900' or dt_pagamento is null)"
mySQL = mySQL & " AND (fl_envio_json=0 or fl_envio_json IS NULL)"
If bd_aux.State = adStateOpen Then bd_aux.Close
Set bd_aux = cn.Execute(mySQL, 3)
While Not bd_aux.EOF
   
   'Recalcula o dígito verificador para o nosso número (é diferente para este envio...)
   'Tira o dígito verificador
   NossoNum = Left(Trim("" & bd_aux("cd_nosso_numero")), Len(Trim("" & bd_aux("cd_nosso_numero"))) - 2)
   NossoNum = Replace(NossoNum, "/", "")
   auxDig = Digito10("8842" & "01585" & NossoNum)
   'Tira a carteira do nossoNumero
   NossoNum = Right(NossoNum, 8)
   
   '1.1 Principal
   Reg = "{"
   Reg = Reg & vbCrLf & "'tipo_ambiente' : 1,"      '1.teste, 2.produção
   Reg = Reg & vbCrLf & "'tipo_registro' : 1,"      '1.registro, 2.alteração, 3.consulta
   Reg = Reg & vbCrLf & "'tipo_cobranca' : 1,"      '1.boletos, 2.débito automático, 3.cartão de crédito, 4.TEF reversa
   Reg = Reg & vbCrLf & "'tipo_produto' : '00006'," 'cliente
   Reg = Reg & vbCrLf & "'subproduto' : '00008',"   'cliente
   Reg = Reg & vbCrLf & "'titulo_aceite' : 'S',"    'S.cobrança, N.proposta (Para carteiras 986 e 885: Utilizar 'N')
   Reg = Reg & vbCrLf & "'indicador_titulo_negociado' : 'S',"   'S.cobrança, N.proposta (Para carteiras 986 e 885: Utilizar 'N')
   '(deve mudar de 175 para 109. Se mudar usar a linha abaixo)
   'Reg = Reg & vbCrLf & "'tipo_carteira_titulo' : '109',"       '109 DIRETA ELETRÔNICA SEM EMISSÃO – SIMPLES (Para carteiras 986 e 885: Utilizar: 885 – Renegociação PF, e 986 - (PFE) Reneg Consórcio)
   Reg = Reg & vbCrLf & "'tipo_carteira_titulo' : '175',"       '175 DIRETA ELETRÔNICA SEM EMISSÃO – SIMPLES (Para carteiras 986 e 885: Utilizar: 885 – Renegociação PF, e 986 - (PFE) Reneg Consórcio)
   Reg = Reg & vbCrLf & "'numero_contrato_proposta' : '" & Format(CDbl(0 & bd_aux("cd_boleto")), "00000000000000000000") & "',"
   Reg = Reg & vbCrLf & "'nosso_numero' : '" & NossoNum & "',"
   Reg = Reg & vbCrLf & "'digito_verificador_nosso_numero' : '" & auxDig & "',"
   If Len("" & bd_aux("cd_codigo_barras")) = 44 Then
      Reg = Reg & vbCrLf & "'codigo_barras' : '" & bd_aux("cd_codigo_barras") & "',"
   Else
      Reg = Reg & vbCrLf & "'codigo_barras' : '00000000000000000000000000000000000000000000',"
   End If
   Reg = Reg & vbCrLf & "'data_vencimento' : '" & Format(MsDt("" & bd_aux("dt_limite")), "yyyy-MM-dd") & "',"        'data do vencimento "yyyy-MM-dd"
   vlAux = CCur(0 & pp.MsVl("" & bd_aux("vl_boleto")))
   Reg = Reg & vbCrLf & "'valor_cobrado' : '" & Format(Left$(Fix(vlAux) & Right(Format(CCur(vlAux) - Fix(vlAux), "#,##0.00"), 2), 17), "00000000000000000") & "',"
   Reg = Reg & vbCrLf & "'especie' : '99',"                         'espécie do título - 99 diversos
   Reg = Reg & vbCrLf & "'data_emissao' : '" & Format(MsDt("" & bd_aux("dt_cadastro")), "yyyy-MM-dd") & "',"         'data de emissão "yyyy-MM-dd"
   Reg = Reg & vbCrLf & "'data_limite_pagamento' : '" & Format(MsDt("" & bd_aux("dt_limite")), "yyyy-MM-dd") & "',"  'data limite para pagamento "yyyy-MM-dd"
   Reg = Reg & vbCrLf & "'tipo_pagamento' : 1,"                     '1.a vista, 3.com data de vencimento determinada
   Reg = Reg & vbCrLf & "'indicador_pagamento_parcial' : 'false',"  'false - não aceita pagamento parcial
   Reg = Reg & vbCrLf & "'instrucao_cobranca_1' : '05',"            'instruções de cobrança 1 - 05 Receber conforme instruções no própio título
   Reg = Reg & vbCrLf & "'instrucao_cobranca_2' : '39',"            'instruções de cobrança 2 - 39 não receber após o vencimento
   vlAux = CCur(0 & pp.MsVl("" & bd_aux("vl_desconto")))
   Reg = Reg & vbCrLf & "'valor_abatimento' : '" & Format(Left$(Fix(vlAux) & Right(Format(CCur(vlAux) - Fix(vlAux), "#,##0.00"), 2), 17), "00000000000000000") & "',"
   Reg = Reg & vbCrLf
   
   '1.2 seção beneficiário
   Reg = Reg & vbCrLf & "'beneficiario' : {"
   Reg = Reg & vbCrLf & "'cpf_cnpj_beneficiario' : '81144396000142',"
   Reg = Reg & vbCrLf & "'nome_beneficiario' : 'SCHULZE ADVOGADOS ASSOCIADOS',"
   Reg = Reg & vbCrLf & "'agencia_beneficiario' : '8842',"
   Reg = Reg & vbCrLf & "'conta_beneficiario' : '0001585',"
   Reg = Reg & vbCrLf & "'digito_verificador_conta_beneficiario' : '4'"
   Reg = Reg & vbCrLf & "}," & vbCrLf
   
   '1.3 seção débito em conta - opcional
   
   '1.4 seção pagador
   mySQL = "SELECT tp_pessoa, no_cliente, nr_cpf, nr_cgc,"
   mySQL = mySQL & " cobranca.dbo.MostraEnderecoCodigo(cd_cliente,1) AS ENDERECO,"
   mySQL = mySQL & " cobranca.dbo.MostraEnderecoBairroCodigo(cd_cliente,1) AS BAIRRO,"
   mySQL = mySQL & " cobranca.dbo.MostraEnderecoCidadeCodigo(cd_cliente,1) AS CIDADE,"
   mySQL = mySQL & " cobranca.dbo.MostraEnderecoestadoCodigo(cd_cliente,1) AS UF,"
   mySQL = mySQL & " cobranca.dbo.MostraEnderecocepcliente(cd_cliente,1) AS CEP"
   mySQL = mySQL & " FROM cliente"
   mySQL = mySQL & " WHERE cd_cliente=" & CDbl(0 & bd_aux("cd_devedor"))
   If bd_end.State = adStateOpen Then bd_end.Close
   Set bd_end = cn.Execute(mySQL, 3)
   If Not bd_end.EOF Then
      Reg = Reg & vbCrLf & "'pagador' : {"
      If UCase(Trim("" & bd_end("tp_pessoa"))) = "F" Then
         Reg = Reg & vbCrLf & "'cpf_cnpj_pagador' : '" & Format(CDbl(0 & bd_end("nr_cpf")), "00000000000") & "',"
      Else
         Reg = Reg & vbCrLf & "'cpf_cnpj_pagador' : '" & Format(CDbl(0 & bd_end("nr_cgc")), "00000000000000") & "',"
      End If
      Reg = Reg & vbCrLf & "'nome_pagador' : '" & Format(Left$(Trim("" & bd_end("no_cliente")), 30), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & "',"
      Reg = Reg & vbCrLf & "'logradouro_pagador' : '" & Format(Left(Trim("" & bd_end("ENDERECO")), 40), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & "',"
      Reg = Reg & vbCrLf & "'bairro_pagador' : '" & Format(Left(Trim("" & bd_end("BAIRRO")), 15), "!@@@@@@@@@@@@@@@") & "',"
      Reg = Reg & vbCrLf & "'cidade_pagador' : '" & Format(Left(Trim("" & bd_end("CIDADE")), 20), "!@@@@@@@@@@@@@@@@@@@@") & "',"
      Reg = Reg & vbCrLf & "'uf_pagador' : '" & Format(Trim("" & bd_end("UF")), "!@@") & "',"
      Reg = Reg & vbCrLf & "'cep_pagador' : '" & Format(CDbl(0 & bd_end("CEP")), "00000000") & "'"
      Reg = Reg & vbCrLf & "}," & vbCrLf
   Else
      MsgProblema "Informações sobre o devedor não encontradas, não será enviado o Boleto: " & CDbl(0 & bd_aux("cd_boleto"))
      bd_end.Close
      GoTo Proximo_Boleto
   End If
   bd_end.Close
   
   '1.5 seção sacador avalista - opcional
   
   '1.6 seção moeda
   Reg = Reg & vbCrLf & "'moeda' : {"
   Reg = Reg & vbCrLf & "'codigo_moeda_cnab' : '9'"
   Reg = Reg & vbCrLf & "}," & vbCrLf
   
   '1.7 seção juros
   Reg = Reg & vbCrLf & "'juros' : {"
   Reg = Reg & vbCrLf & "'tipo_juros' : 5"
   Reg = Reg & vbCrLf & "}," & vbCrLf
   
   '1.8 seção multa
   Reg = Reg & vbCrLf & "'multa' : {"
   Reg = Reg & vbCrLf & "'tipo_multa' : 3"
   Reg = Reg & vbCrLf & "}," & vbCrLf
   
   '1.9 seção desconto
   Reg = Reg & vbCrLf & "'grupo_desconto' : [ {"
   Reg = Reg & vbCrLf & "'data_desconto' : '" & Format(MsDt("" & bd_aux("dt_limite")), "yyyy-MM-dd") & "',"  'data referência para desconto "yyyy-MM-dd"
   vlAux = CCur(0 & pp.MsVl("" & bd_aux("vl_desconto")))
   If vlAux = 0 Then
      Reg = Reg & vbCrLf & "'tipo_desconto' : '0'," 'sem desconto
      Reg = Reg & vbCrLf & "'valor_desconto' : '00000000000000000'"
   Else
      Reg = Reg & vbCrLf & "'tipo_desconto' : '1'," 'valor fixo
      Reg = Reg & vbCrLf & "'valor_desconto' : '" & Format(Left$(Fix(vlAux) & Right(Format(CCur(vlAux) - Fix(vlAux), "#,##0.00"), 2), 17), "00000000000000000") & "'"
   End If
   Reg = Reg & vbCrLf & "} ]," & vbCrLf
   
   '1.10 seção recebimento divergente
   Reg = Reg & vbCrLf & "'recebimento_divergente' : {"
   Reg = Reg & vbCrLf & "'tipo_autorizacao_recebimento' : '3'" '3. o título não deve aceitar pagamentos de valores divergentes ao da cobrança
   Reg = Reg & vbCrLf & "}," & vbCrLf
   
   '1.11 seção rateio
   Reg = Reg & vbCrLf & "'grupo_rateio' : [ ]"
   Reg = Reg & vbCrLf & "}"
   Reg = Reg & vbCrLf
   
   'Trocar aspas simples por aspas duplas
   Reg = Replace(Reg, "'", """")
   
   'REALIZAR O REGISTRO DO BOLETO
   pURL = "http://panserviceshml.bancopan.com.br/PanRegistraBoletoCIP/RegistraBoleto?Dados="
   'homologação - https://autorizador-boletos.itau.com.br
   pURL = "https://gerador-boletos.itau.com.br/router-gateway-app/public/codigo_barras/registro"
   
   pErro = True
   pStatus = 0
   pRetorno = ""
   Registrar_Boleto pURL, pToken, Reg, pStatus, pRetorno
   
   Select Case pStatus
   Case 200
      '(OK) A chamada foi bem sucedida
      pErro = False
   Case 400
      MsgAviso "Erro ao enviar o Boleto: " & CDbl(0 & bd_aux("cd_boleto")) & vbCrLf & "(Bad Request) A requisição é inválida, em geral conteúdo mal formado." & vbCrLf & vbCrLf & pRetorno
   Case 401
      MsgAviso "Erro ao enviar o Boleto: " & CDbl(0 & bd_aux("cd_boleto")) & vbCrLf & "(Unauthorized) O usuário e senha ou token de acesso são inválidos." & vbCrLf & vbCrLf & pRetorno
   Case 403
      MsgAviso "Erro ao enviar o Boleto: " & CDbl(0 & bd_aux("cd_boleto")) & vbCrLf & "(Forbidden) O acesso à API está bloqueado ou o usuário está bloqueado." & vbCrLf & vbCrLf & pRetorno
   Case 404
      MsgAviso "Erro ao enviar o Boleto: " & CDbl(0 & bd_aux("cd_boleto")) & vbCrLf & "(Not Found) O endereço acessado não existe." & vbCrLf & vbCrLf & pRetorno
   Case 422
      MsgAviso "Erro ao enviar o Boleto: " & CDbl(0 & bd_aux("cd_boleto")) & vbCrLf & "(Unprocessable Entity) A Requisição é válida, mas os dados passados não são válidos." & vbCrLf & vbCrLf & pRetorno
   Case 429
      MsgAviso "Erro ao enviar o Boleto: " & CDbl(0 & bd_aux("cd_boleto")) & vbCrLf & "(Too Many Requests) O usuário atingiu o limite de requisições." & vbCrLf & vbCrLf & pRetorno
   Case 500
      MsgAviso "Erro ao enviar o Boleto: " & CDbl(0 & bd_aux("cd_boleto")) & vbCrLf & "(Internal Server Error) Houve um erro interno do servidor ao processar a requisição." & vbCrLf & vbCrLf & pRetorno
   End Select
   
   'se não deu erro, grava o indicador de que registrou o boleto
   If pErro = False Then
      mySQL = "UPDATE BoletoEscritorioItau SET fl_envio_json=1"
      mySQL = mySQL & " WHERE cd_boleto=" & CDbl(0 & bd_aux("cd_boleto"))
      mySQL = mySQL & " AND cd_processo=" & CDbl(0 & bd_aux("cd_processo"))
      cn.Execute mySQL, , adExecuteNoRecords
   End If
   
Proximo_Boleto:
   
   bd_aux.MoveNext
Wend
bd_aux.Close

MsgAviso "Processo encerrado !!!"

On Error GoTo 0
relogio 0
Exit Sub
Resume

Erros:
erro Err
On Error GoTo 0
relogio 0
End Sub

Function Get_Json(pURL As String) As String
   Dim pHTTP As Object
   Set pHTTP = CreateObject("MSXML2.ServerXMLHTTP")
   
   pHTTP.open "GET", pURL, False
   pHTTP.send
   
   Get_Json = pHTTP.responseText
   Set pHTTP = Nothing
End Function

Function Ler_Token(pURL As String, pAcesso As String) As String
   Dim pHTTP As Object
   Set pHTTP = CreateObject("MSXML2.ServerXMLHTTP")
   
   pHTTP.open "POST", pURL, False
   pHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
   pHTTP.setRequestHeader "Authorization", "Basic " & pAcesso
   pHTTP.send ("scope=readonly&grant_type=client_credentials")
   
   Ler_Token = pHTTP.responseText
   Set pHTTP = Nothing
End Function

Sub Registrar_Boleto(ByVal pURL As String, ByVal pToken As String, ByVal pBoleto As String, ByRef pStatus As Integer, ByRef pRetorno As String)
   Dim pHTTP As Object
   Set pHTTP = CreateObject("MSXML2.ServerXMLHTTP")
   
   pHTTP.open "POST", pURL, False
   pHTTP.setRequestHeader "access_token", pToken
   pHTTP.setRequestHeader "Accept", "application/vnd.itau"
   pHTTP.setRequestHeader "identificador", "81144396000142"
   pHTTP.setRequestHeader "itau-chave", "9a6a013b-54df-49a5-bf99-f674761f5775"
   pHTTP.send (pBoleto)
   
   pStatus = pHTTP.Status
   pRetorno = pHTTP.responseText
   
   Set pHTTP = Nothing
End Sub

Private Sub bt_LerArquivo_Click()
Dim vl_total_dividas As Currency
Dim qt_fichas As Long
Dim bd_aux As New ADODB.Recordset
Dim bd_cobrador As New ADODB.Recordset
Dim bd_cli As New ADODB.Recordset
Dim bd_matricula As New ADODB.Recordset
Dim Convenio As Double
Dim Matricula As String
Dim valor1 As Currency
Dim valor2 As Currency
Dim processo As Double, cliente As Double, devedor As Double, ProAnt As Double
Dim nrBoleto As Double
Dim cd_criacao As String
Dim dt_criacao As String
Dim cd_alteracao As String
Dim i As Integer

'Nelson - 26/10/2017 - confere se a data da variável global é igual ao dia de hoje
' a criação das fases está dando erro porque usa a variável global e se o cob foi aberto no dia anterior vai gravar a fase como se fosse criada ontem
If pp.MsDt(data_atual) <> pp.MsDt(Now) Then
   MsgAviso "Erro na data interna do cobSystem." & vbCrLf & "Por favor, feche e abra o cobSystem novamente para atualizar as variáveis internas."
   bo_arquivo.SetFocus
   Exit Sub
End If

If Trim(bo_arquivo.text) = "" Then
   MsgAviso "Informe o arquivo!"
   bo_arquivo.SetFocus
   Exit Sub
End If

If op_inclui_fase.Value = False Then
   If MsgPerguntaNao("Deseja mesmo incluir as fases na leitura?") = 7 Then
      op_inclui_fase.Value = True
      DoEvents
   End If
End If

If op_tipo.valor = "1" Then
   tpArquivo = "ESCRITORIOITAU"
   Tabela = "BoletoEscritorioItau"
   Tabela2 = "CalculoBoletoEscritorioItau"
   LerEscritorioItau
   Exit Sub
ElseIf op_tipo.valor = "2" Then
   tpArquivo = "ESCRITORIO"
   Tabela = "Boleto"
   Tabela2 = "CalculoBoleto"
ElseIf op_tipo.valor = "3" Then
   tpArquivo = "ESCRITORIOSANTANDER"
   Tabela = "BoletoEscritorioSantander"
   Tabela2 = "CalculoBoletoEscritorioSantander"
ElseIf op_tipo.valor = "4" Then
   tpArquivo = "FINASA"
   Tabela = "BoletoFinasa"
   Tabela2 = "CalculoBoletoFinasa"
End If

'importação do arquivo
On Error GoTo Erro_importacao

pp.relogio 1

cn.BeginTrans

i = 0
valor1 = 0
valor2 = 0
ProAnt = 0
nrBoleto = 0
cd_filial = 0

'abrir o arquivo
numarq = FreeFile
Open bo_arquivo.text For Input As numarq

Limpa_Grid
Formata_Grid

'ler os registros
fim = 0
contador = 0
While fim = 0
   Convenio = 0
   On Error Resume Next
   Line Input #numarq, registro
   If Err = 62 Then
      fim = 1
      GoTo fim_leitura
   End If
   On Error GoTo Erro_importacao
   If Not IsNumeric(Left(registro, 1)) Then
      fim = 1
      GoTo fim_leitura
   End If
   
   If Left$(registro, 1) = "1" Then 'Detalhe
      contador = contador + 1
            
      i = i + 1
      gr_cons.Rows = i + 1
      If CCur(0 & Mid$(registro, 254, 13)) = 0 Then
         pp.PoeGrid gr_cons, 5, i, "/  /"
      Else
         pp.PoeGrid gr_cons, 5, i, Format("" & Mid$(Mid$(registro, 111, 10), 1, 2) & "/" & Mid$(Mid$(registro, 111, 10), 3, 2) & "/" & Mid$(Mid$(registro, 111, 10), 5, 2), "dd/MM/yyyy")
         'Só soma os valores quando tem data de pagamento
         valor1 = valor1 + CCur(Format(CCur(Mid$(registro, 153, 11) & "," & Mid$(registro, 164, 2)), "#,##0.00"))
         valor2 = valor2 + CCur(Format(CCur(Mid$(registro, 254, 11) & "," & Mid$(registro, 265, 2)), "#,##0.00"))
      End If
            
      pp.PoeGrid gr_cons, 2, i, Format(CCur(Mid$(registro, 153, 11) & "," & Mid$(registro, 164, 2)), "#,##0.00")
      pp.PoeGrid gr_cons, 3, i, Format(CCur(Mid$(registro, 254, 11) & "," & Mid$(registro, 265, 2)), "#,##0.00")
      
      If tpArquivo = "FINASA" Then
         nrBoleto = CDbl(0 & Mid$(registro, 74, 8))
         pp.PoeGrid gr_cons, 0, i, Mid$(registro, 74, 8)
         pp.PoeGrid gr_cons, 1, i, Mid$(registro, 74, 8)
      ElseIf tpArquivo = "ESCRITORIOSANTANDER" Then
         nrBoleto = CDbl(0 & Mid$(registro, 46, 14))
         pp.PoeGrid gr_cons, 0, i, Mid$(registro, 38, 7)
         pp.PoeGrid gr_cons, 1, i, Mid$(registro, 38, 7)
      Else
         nrBoleto = CDbl(0 & Mid$(registro, 63, 8))
         pp.PoeGrid gr_cons, 0, i, Mid$(registro, 63, 15)
         pp.PoeGrid gr_cons, 1, i, Mid$(registro, 63, 8)
      End If
      
      'Cliente
      Matricula = ""
      mySQL = "SELECT cd_cliente, dt_limite FROM " & Tabela & "  WHERE cd_boleto=" & nrBoleto
      Set bd_aux = cn.Execute(mySQL, 3)
      If Not bd_aux.EOF Then
         Matricula = Mostra_Cliente(CLng(0 & bd_aux("cd_cliente")))
         pp.PoeGrid gr_cons, 4, i, pp.MsDt("" & bd_aux("dt_limite"))
      End If
      bd_aux.Close
      pp.PoeGrid gr_cons, 7, i, Trim(Matricula)
            
      'Devedor
      Matricula = ""
      mySQL = "SELECT cd_processo, cd_cliente, cd_devedor, dt_vencimento, dt_criacao, cd_criacao, cd_alteracao, cd_filial FROM " & Tabela & " WHERE cd_boleto=" & nrBoleto
      Set bd_aux = cn.Execute(mySQL, 3)
      If Not bd_aux.EOF Then
         processo = CDbl(0 & bd_aux("cd_processo"))
         cliente = CDbl(0 & bd_aux("cd_cliente"))
         devedor = CDbl(0 & bd_aux("cd_devedor"))
         Matricula = Mostra_Cliente(CLng(0 & bd_aux("cd_devedor")))
         
         cd_criacao = Trim("" & bd_aux("cd_criacao"))
         'Alterado para data de Pagamento conforme Chamado 28122010102101
         'dt_criacao = pp.GrDtString(glb_tipo_banco, pp.MsDt("" & bd_aux("dt_criacao")))
         If Trim(pp.PegaGrid(gr_cons, 5, i)) <> "/  /" Then
            dt_criacao = pp.GrDtString(glb_tipo_banco, Trim(pp.PegaGrid(gr_cons, 5, i)))
         Else
            dt_criacao = pp.GrDtString(glb_tipo_banco, Format(Now, "dd/MM/yyyy"))
         End If
         
         cd_alteracao = IIf(Trim("" & bd_aux("cd_alteracao")) = "", Trim("" & bd_aux("cd_criacao")), Trim("" & bd_aux("cd_alteracao")))
         cd_filial = CDbl(0 & bd_aux("cd_filial"))
         
         'Consulta o cobrador do contrato no dia de cadastro
         If CDbl(0 & bd_aux("cd_processo")) <> 0 Then
            mySQL = "SELECT " & cn.DefaultDatabase & ".dbo.MostraCobradorFicha(" & CDbl(0 & processo) & "," & dt_criacao & ") AS Cobrador"
            Set bd_cobrador = cn.Execute(mySQL, 3)
            If Not bd_cobrador.EOF Then
               If Trim("" & bd_cobrador("Cobrador")) <> "" Then
                  mySQL = "SELECT cd_usuario AS Cobrador FROM Cliente WHERE cd_cliente=" & CDbl(0 & bd_cobrador("Cobrador"))
                  Set bd_cobrador = cn.Execute(mySQL, 3)
                  If Not bd_cobrador.EOF Then
                     cd_criacao = Trim("" & bd_cobrador("Cobrador"))
                  End If
               End If
            End If
            bd_cobrador.Close
            
            'Pega a filial do Cobrador Foco
            fl_foco = 0
            mySQL = "SELECT cd_cliente FROM Cliente WHERE fl_foco=1 AND cd_usuario='" & cd_criacao & "'"
            Set bd_cobrador = cn.Execute(mySQL, 3)
            If Not bd_cobrador.EOF Then
               fl_foco = 1
            End If
            bd_cobrador.Close
            
            If CInt(0 & fl_foco) <> 0 And CDbl(0 & bd_aux("cd_devedor")) <> 0 Then
               mySQL = "SELECT cd_filialss AS Filial FROM Cliente WHERE cd_cliente=" & CDbl(0 & bd_aux("cd_devedor"))
               Set bd_cobrador = cn.Execute(mySQL, 3)
               If Not bd_cobrador.EOF Then
                  cd_filial = CDbl(0 & bd_cobrador("Filial"))
               End If
               bd_cobrador.Close
            Else
               If Trim(cd_criacao) <> "" Then
                  mySQL = "SELECT " & cn.DefaultDatabase & ".dbo.MostraCodFilialUsuario('" & cd_criacao & "') AS Filial"
                  Set bd_cobrador = cn.Execute(mySQL, 3)
                  If Not bd_cobrador.EOF Then
                     If Trim("" & bd_cobrador("Filial")) <> "" Then
                        cd_filial = CDbl("" & bd_cobrador("Filial"))
                     End If
                   End If
                   bd_cobrador.Close
               End If
            End If
         End If
      Else
         'Nelson - 07/12/2017 - se não encontrar o boleto passa para a próxima linha
         bd_aux.Close
         pp.PoeGrid gr_cons, 8, i, Trim(Matricula)
         pp.PoeGrid gr_cons, 6, i, "Boleto não cadast."
         GoTo fim_leitura
      End If
      bd_aux.Close
      pp.PoeGrid gr_cons, 8, i, Trim(Matricula)
      
      If Trim$(Mid$(registro, 303, 2)) <> "CE" And tpArquivo <> "FINASA" And tpArquivo <> "ESCRITORIOSANTANDER" Then
         pp.PoeGrid gr_cons, 9, i, "SIM"
      End If
      
      Select Case Trim$(Mid$(registro, 109, 2))
      Case "06", "17"
         pp.PoeGrid gr_cons, 6, i, "Título Pago"
         'Nelson - 09/06/2017 - Só grava se o campo 'Incluir Fase' estiver desmarcado
         If op_inclui_fase.Value = False Then
            mySQL = "UPDATE " & Tabela & " SET"
            mySQL = mySQL & " dt_pagamento='" & GrDtString(Format("" & Mid$(Mid$(registro, 111, 10), 1, 2) & "/" & Mid$(Mid$(registro, 111, 10), 3, 2) & "/" & Mid$(Mid$(registro, 111, 10), 5, 2), "dd/MM/yyyy")) & "',"
            mySQL = mySQL & " cd_filial=" & CDbl(0 & cd_filial) & ","
            mySQL = mySQL & " cd_criacao='" & Trim(cd_criacao) & "',"
            mySQL = mySQL & " cd_alteracao='" & Trim(cd_alteracao) & "',"
            mySQL = mySQL & " vl_pago=" & pp.GrVl(Format(CCur(Mid$(registro, 254, 11) & "," & Mid$(registro, 265, 2)), "#,##0.00"))
            mySQL = mySQL & " WHERE cd_boleto=" & CDbl(0 & nrBoleto)
            'mySQL = mySQL & " AND cd_processo=" & CDbl(0 & processo)
            cn.Execute mySQL, , adExecuteNoRecords
            If processo <> 0 Then
               If ProAnt <> processo Then
                  ProAnt = processo
                  If Trim$(Mid$(registro, 303, 2)) <> "CE" And tpArquivo <> "FINASA" And tpArquivo <> "ESCRITORIOSANTANDER" Then
                     IncluiFase processo, 439, ""
                  Else
                     IncluiFase processo, 69, ""
                  End If
               End If
            End If
         End If
      Case Else
         Select Case CDbl(0 & Mid$(registro, 378, 8))
            Case "01"
               pp.PoeGrid gr_cons, 6, i, "Falta de Registro header"
            Case "02"
               pp.PoeGrid gr_cons, 6, i, "Falta de Registro trailer"
            Case "03"
               pp.PoeGrid gr_cons, 6, i, "Grupo empresarial não cadastrado no sistema"
            Case "04"
               pp.PoeGrid gr_cons, 6, i, "Versão do arquivo inválido"
            Case "05"
               pp.PoeGrid gr_cons, 6, i, "Tipo de formulário inválido"
            Case "06"
               pp.PoeGrid gr_cons, 6, i, "Tipo de Crítica de referência inválida"
            Case "07"
               pp.PoeGrid gr_cons, 6, i, "Tipo de postagem inválida"
            Case "08"
               pp.PoeGrid gr_cons, 6, i, "Grupo sem registro detalhe"
            Case "09"
               pp.PoeGrid gr_cons, 6, i, "Número de referência inválido"
            Case "10"
               pp.PoeGrid gr_cons, 6, i, "Data de vencimento inválida"
            Case "11"
               pp.PoeGrid gr_cons, 6, i, "Nome do sacado inválido"
            Case "12"
               pp.PoeGrid gr_cons, 6, i, "Endereço / Cep do sacado inválido"
            Case "13"
               pp.PoeGrid gr_cons, 6, i, "Agência depositária inválida"
            Case "14"
               pp.PoeGrid gr_cons, 6, i, "Código da moeda inválido"
            Case "15"
               pp.PoeGrid gr_cons, 6, i, "Valor do Título inválido"
            Case "16"
               pp.PoeGrid gr_cons, 6, i, "Quantidade de moeda incompatível com código da moeda"
            Case "17"
               pp.PoeGrid gr_cons, 6, i, "Indicador de mensagem inválido"
            Case "18"
               pp.PoeGrid gr_cons, 6, i, "Registro detalhe sem mensagem"
            Case "19"
               pp.PoeGrid gr_cons, 6, i, "Registro mensagem inválido"
            Case "20"
               pp.PoeGrid gr_cons, 6, i, "Registro mensagem sem detalhe"
            Case "21"
               pp.PoeGrid gr_cons, 6, i, "Total de registros difere do apurado"
            Case "22"
               pp.PoeGrid gr_cons, 6, i, "Somatória dos títulos difere do apurado"
            Case "23"
               pp.PoeGrid gr_cons, 6, i, "Remessa desprezada"
            Case "24"
               pp.PoeGrid gr_cons, 6, i, "Código de carteira inválido"
            Case "25"
               pp.PoeGrid gr_cons, 6, i, "Número para código de barras inválido"
            Case "26"
               pp.PoeGrid gr_cons, 6, i, "Sequência do registro inválida"
            Case "27"
               pp.PoeGrid gr_cons, 6, i, "Quantidade do registro total inválida"
            Case "28"
               pp.PoeGrid gr_cons, 6, i, "Valor do registro total inválido"
            Case "29"
               pp.PoeGrid gr_cons, 6, i, "Tipo de registro inválido"
            Case "30"
               pp.PoeGrid gr_cons, 6, i, "Agência cedente inválida"
            Case "31"
               pp.PoeGrid gr_cons, 6, i, "Conta cedente inválida"
            Case "32"
               pp.PoeGrid gr_cons, 6, i, "Valor desconto/juros/multa inválido"
            Case "33"
               pp.PoeGrid gr_cons, 6, i, "Faixa de nosso número inválida"
            Case "34"
               pp.PoeGrid gr_cons, 6, i, "Carnê desprezado - parcelas inconsistentes"
            Case "35"
               pp.PoeGrid gr_cons, 6, i, "Tipo de carnê inválido"
            Case "36"
               pp.PoeGrid gr_cons, 6, i, "Total de parcelas inválida"
            Case "37"
               pp.PoeGrid gr_cons, 6, i, "Total de parcelas não é única no carnê"
            Case "38"
               pp.PoeGrid gr_cons, 6, i, "Número de parcela inválida"
            Case "39"
               pp.PoeGrid gr_cons, 6, i, "Número de parcela fora da sequência"
            Case "42"
               pp.PoeGrid gr_cons, 6, i, "Carnê com a c/c head diferente da c/c detalhe"
            Case "43"
               pp.PoeGrid gr_cons, 6, i, "Número do documento inválido"
            Case "44"
               pp.PoeGrid gr_cons, 6, i, "Agência cedente (detalhe) inválida"
            Case "45"
               pp.PoeGrid gr_cons, 6, i, "Conta cedente (detalhe) inválida"
            Case "46"
               pp.PoeGrid gr_cons, 6, i, "Mensagens inválidas"
            Case "47"
               pp.PoeGrid gr_cons, 6, i, "Cep do sacado não é único no carnê"
            Case "48"
               pp.PoeGrid gr_cons, 6, i, "Agência e conta do cedente não é única no carnê"
            Case "53"
               pp.PoeGrid gr_cons, 6, i, "Parâmetros incremento da referência inválido"
            Case "54"
               pp.PoeGrid gr_cons, 6, i, "Data vencimento incompatível com parâmetro"
            Case "55"
               pp.PoeGrid gr_cons, 6, i, "Valor do título incompatível com parâmetro"
            Case "56"
               pp.PoeGrid gr_cons, 6, i, "Valor do desconto incompatível com parâmetro"
            Case "57"
               pp.PoeGrid gr_cons, 6, i, "Valor do juros incompatível com parâmetro"
            Case "58"
               pp.PoeGrid gr_cons, 6, i, "Valor da multa incompatível com parâmetro"
            Case "59"
               pp.PoeGrid gr_cons, 6, i, "Data do desconto incompatível com parâmetro"
            Case "60"
               pp.PoeGrid gr_cons, 6, i, "Data da multa incompatível com parâmetro"
            End Select
      End Select
   End If
fim_leitura:
Wend

i = i + 2
gr_cons.Rows = i + 1
pp.PoeGrid gr_cons, 1, i, "TOTAL"
pp.PoeGrid gr_cons, 2, i, Format(valor1, "#,##0.00")
pp.PoeGrid gr_cons, 3, i, Format(valor2, "#,##0.00")

i = i + 1
gr_cons.Rows = i + 1

Finalizacao:

If contador = 0 Then
   cn.RollbackTrans
   MsgAviso "Não foram encontrados títulos nesse arquivo!"
Else
   cn.CommitTrans
   If contador = 1 Then
      MsgAviso "Importação efetuada com sucesso! " & vbCrLf & "Foi lido 1 título do arquivo."
   Else
      MsgAviso "Importação efetuada com sucesso! " & vbCrLf & "Foram lidos o total de " & contador & " títulos do arquivo."
   End If
End If

If CCur(0 & valor1) <> CCur(0 & valor2) Then
   MsgAviso "Existem Títulos Pagos com valor menor ou maior que o impresso no boleto!" & vbCrLf & "Valor total do boletos: " & Format(valor1, "R$ #,##0.00") & "." & vbCrLf & "Valor total recebido: " & Format(valor2, "R$ #,##0.00") & "."
Else
   MsgAviso "Valor dos boletos igual ao valor pago." & vbCrLf & "Valor total do boletos: " & Format(valor1, "R$ #,##0.00") & "." & vbCrLf & "Valor total recebido: " & Format(valor2, "R$ #,##0.00") & "."
End If

Close numarq
pp.relogio 0
Exit Sub

Resume
Erro_importacao:
erro Err
cn.RollbackTrans
Close numarq
On Error GoTo 0
pp.relogio 0
End Sub

Sub LerEscritorioItau()
Dim vl_total_dividas As Currency
Dim qt_fichas As Long
Dim bd_aux As New ADODB.Recordset
Dim bd_cobrador As New ADODB.Recordset
Dim bd_cli As New ADODB.Recordset
Dim bd_matricula As New ADODB.Recordset
Dim Convenio As Double
Dim Matricula As String
Dim valor1 As Currency
Dim valor2 As Currency
Dim processo As Double, cliente As Double, devedor As Double, ProAnt As Double
Dim nrBoleto As Double
Dim cd_criacao As String
Dim dt_criacao As String
Dim cd_alteracao As String
Dim i As Integer

If Trim(bo_arquivo.text) = "" Then
   MsgAviso "Informe o arquivo!"
   bo_arquivo.SetFocus
   Exit Sub
End If

'importacao do arquivo
pp.relogio 1
On Error GoTo Erro_importacao
cn.BeginTrans

i = 0
valor1 = 0
valor2 = 0
ProAnt = 0
nrBoleto = 0
cd_filial = 0

'abrir o arquivo
numarq = FreeFile
Open bo_arquivo.text For Input As numarq

Limpa_Grid
Formata_Grid

'ler os registros
fim = 0
contador = 0
While fim = 0
   Convenio = 0
   On Error Resume Next
   Line Input #numarq, registro
   If Err = 62 Then
      fim = 1
      GoTo fim_leitura
   End If
   On Error GoTo Erro_importacao
   If Not IsNumeric(Left(registro, 1)) Then
      fim = 1
      GoTo fim_leitura
   End If
   If Left$(registro, 1) = "1" Then 'Detalhe
      contador = contador + 1
            
      i = i + 1
      gr_cons.Rows = i + 1
      If CCur(0 & Mid$(registro, 254, 13)) = 0 Then
         pp.PoeGrid gr_cons, 5, i, "/  /"
      Else
         pp.PoeGrid gr_cons, 5, i, Format("" & Mid$(Mid$(registro, 111, 10), 1, 2) & "/" & Mid$(Mid$(registro, 111, 10), 3, 2) & "/" & Mid$(Mid$(registro, 111, 10), 5, 2), "dd/MM/yyyy")
      End If
            
      pp.PoeGrid gr_cons, 2, i, Format(CCur(Mid$(registro, 153, 11) & "," & Mid$(registro, 164, 2)), "#,##0.00")
      pp.PoeGrid gr_cons, 3, i, Format(CCur(Mid$(registro, 254, 11) & "," & Mid$(registro, 265, 2)), "#,##0.00")
      
      pp.PoeGrid gr_cons, 3, i, Format(CCur(CCur(0 & pp.PegaGrid(gr_cons, 3, i)) + CCur(Format(CCur(Mid$(registro, 176, 11) & "," & Mid$(registro, 187, 2)), "#,##0.00"))), "#,##0.00")
      
      valor1 = valor1 + CCur(0 & pp.PegaGrid(gr_cons, 2, i))
      valor2 = valor2 + CCur(0 & pp.PegaGrid(gr_cons, 3, i))
      
      nrBoleto = CDbl(0 & Mid$(registro, 63, 8))
      pp.PoeGrid gr_cons, 0, i, Mid$(registro, 83, 11)
      pp.PoeGrid gr_cons, 1, i, Mid$(registro, 63, 8)
      
      'Cliente
      Matricula = ""
      mySQL = "SELECT cd_cliente, dt_limite FROM " & Tabela & "  WHERE cd_boleto=" & nrBoleto
      Set bd_aux = cn.Execute(mySQL, 3)
      If Not bd_aux.EOF Then
         Matricula = Mostra_Cliente(CLng(0 & bd_aux("cd_cliente")))
         pp.PoeGrid gr_cons, 4, i, pp.MsDt("" & bd_aux("dt_limite"))
      End If
      bd_aux.Close
      pp.PoeGrid gr_cons, 7, i, Trim(Matricula)
            
      'Devedor
      Matricula = ""
      mySQL = "SELECT cd_processo, cd_cliente, cd_devedor, dt_vencimento, dt_criacao, cd_criacao, cd_alteracao, cd_filial FROM " & Tabela & " WHERE cd_boleto=" & nrBoleto
      Set bd_aux = cn.Execute(mySQL, 3)
      If Not bd_aux.EOF Then
         processo = CDbl(0 & bd_aux("cd_processo"))
         cliente = CDbl(0 & bd_aux("cd_cliente"))
         devedor = CDbl(0 & bd_aux("cd_devedor"))
         Matricula = Mostra_Cliente(CLng(0 & bd_aux("cd_devedor")))
         
         cd_criacao = Trim("" & bd_aux("cd_criacao"))
         dt_criacao = pp.GrDtString(glb_tipo_banco, pp.MsDt("" & bd_aux("dt_criacao")))
         cd_alteracao = IIf(Trim("" & bd_aux("cd_alteracao")) = "", Trim("" & bd_aux("cd_criacao")), Trim("" & bd_aux("cd_alteracao")))
         cd_filial = CDbl(0 & bd_aux("cd_filial"))
         
         'Consulta o cobrador do contrato no dia de cadastro
         If CDbl(0 & bd_aux("cd_processo")) <> 0 Then
            mySQL = "SELECT " & cn.DefaultDatabase & ".dbo.MostraCobradorFicha(" & CDbl(0 & processo) & "," & dt_criacao & ") AS Cobrador"
            Set bd_cobrador = cn.Execute(mySQL, 3)
            If Not bd_cobrador.EOF Then
               If Trim("" & bd_cobrador("Cobrador")) <> "" Then
                  mySQL = "SELECT cd_usuario AS Cobrador FROM Cliente WHERE cd_cliente=" & CDbl(0 & bd_cobrador("Cobrador"))
                  Set bd_cobrador = cn.Execute(mySQL, 3)
                  If Not bd_cobrador.EOF Then
                     If Trim("" & bd_cobrador("Cobrador")) <> "" Then
                        cd_criacao = Trim("" & bd_cobrador("Cobrador"))
                     End If
                  End If
               End If
            End If
            bd_cobrador.Close
            
            'Pega a filial do Cobrador Foco
            fl_foco = 0
            mySQL = "SELECT cd_cliente FROM Cliente WHERE fl_foco=1 AND cd_usuario='" & cd_criacao & "'"
            Set bd_cobrador = cn.Execute(mySQL, 3)
            If Not bd_cobrador.EOF Then
               fl_foco = 1
            End If
            bd_cobrador.Close
            
            If CInt(0 & fl_foco) <> 0 And CDbl(0 & bd_aux("cd_devedor")) <> 0 Then
               mySQL = "SELECT cd_filialss AS Filial FROM Cliente WHERE cd_cliente=" & CDbl(0 & bd_aux("cd_devedor"))
               Set bd_cobrador = cn.Execute(mySQL, 3)
               If Not bd_cobrador.EOF Then
                  cd_filial = Trim("" & bd_cobrador("Filial"))
               End If
               bd_cobrador.Close
            Else
               mySQL = "SELECT " & cn.DefaultDatabase & ".dbo.MostraCodFilialUsuario('" & cd_criacao & "') AS Filial"
               Set bd_cobrador = cn.Execute(mySQL, 3)
               If Not bd_cobrador.EOF Then
                  If Trim("" & bd_cobrador("Filial")) <> "" Then
                     cd_filial = Trim("" & bd_cobrador("Filial"))
                  End If
               End If
               bd_cobrador.Close
            End If
         End If
      End If
      bd_aux.Close
      pp.PoeGrid gr_cons, 8, i, Trim(Matricula)
      
      pp.PoeGrid gr_cons, 9, i, Trim$(Mid$(registro, 393, 2))
      
      Select Case Trim$(Mid$(registro, 109, 2))
      Case "02", "06"
         pp.PoeGrid gr_cons, 6, i, "Título Pago"
         'Nelson - 09/06/2017 - Só grava se o campo estiver desmarcado
         If op_inclui_fase.Value = False Then
            mySQL = "UPDATE " & Tabela & " SET"
            mySQL = mySQL & " dt_pagamento='" & GrDtString(Format("" & Mid$(Mid$(registro, 111, 10), 1, 2) & "/" & Mid$(Mid$(registro, 111, 10), 3, 2) & "/" & Mid$(Mid$(registro, 111, 10), 5, 2), "dd/MM/yyyy")) & "',"
            mySQL = mySQL & " cd_filial=" & CDbl(0 & cd_filial) & ","
            mySQL = mySQL & " cd_criacao='" & Trim(cd_criacao) & "',"
            mySQL = mySQL & " cd_alteracao='" & Trim(cd_alteracao) & "',"
            mySQL = mySQL & " vl_pago=" & pp.GrVl(Format(CCur(Mid$(registro, 254, 11) & "," & Mid$(registro, 265, 2)), "#,##0.00"))
            mySQL = mySQL & " WHERE cd_boleto=" & CDbl(0 & nrBoleto)
            cn.Execute mySQL, , adExecuteNoRecords
            
            If processo <> 0 Then
               If ProAnt <> processo Then
                  ProAnt = processo
                  'If cheque Then
                  '   IncluiFase processo, 439, ""
                  'Else
                     IncluiFase processo, 69, ""
                  'End If
               End If
            End If
         End If
      End Select
   End If
fim_leitura:
Wend

i = i + 2
gr_cons.Rows = i + 1
pp.PoeGrid gr_cons, 1, i, "TOTAL"
pp.PoeGrid gr_cons, 2, i, Format(valor1, "#,##0.00")
pp.PoeGrid gr_cons, 3, i, Format(valor2, "#,##0.00")

i = i + 1
gr_cons.Rows = i + 1

Finalizacao:

If contador = 0 Then
   cn.RollbackTrans
   MsgAviso "Não foram encontrados títulos nesse arquivo!"
Else
   cn.CommitTrans
   If contador = 1 Then
      MsgAviso "Importação efetuada com sucesso! " & vbCrLf & "Foi lido " & contador & " do título do arquivo."
   Else
      MsgAviso "Importação efetuada com sucesso! " & vbCrLf & "Foram lidos o total de " & contador & " títulos do arquivo."
   End If
End If

If CCur(0 & valor1) <> CCur(0 & valor2) Then
   MsgAviso "Existem Títulos Pagos com valor menor ou maior que o impresso no boleto!" & vbCrLf & "Valor total do boletos: " & Format(valor1, "R$ #,##0.00") & "." & vbCrLf & "Valor total recebido: " & Format(valor2, "R$ #,##0.00") & "."
Else
   MsgAviso "Valor dos boletos igual ao valor pago." & vbCrLf & "Valor total do boletos: " & Format(valor1, "R$ #,##0.00") & "." & vbCrLf & "Valor total recebido: " & Format(valor2, "R$ #,##0.00") & "."
End If

Close numarq
pp.relogio 0
Exit Sub

Resume
Erro_importacao:
erro Err
cn.RollbackTrans
Close numarq
On Error GoTo 0
pp.relogio 0
End Sub

Private Sub Form_Load()
pp.relogio 1
pp.AlinhaTela Me, Me.Width, Me.Height
Formata_Grid
tpArquivo = "ESCRITORIOITAU"
If CInt(0 & emp_padrao) = 1 Then
   Tabela = "boleto_secojur"
   Tabela2 = "CalculoBoleto_secojur"
Else
   Tabela = "BoletoEscritorioItau"
   Tabela2 = "CalculoBoletoEscritorioItau"
End If
pp.relogio 0
End Sub

Private Sub gr_cons_DblClick()

If gr_cons.row = 0 Then
   gr_cons.row = 1
End If

If IsNumeric(PegaGrid(gr_cons, 0, gr_cons.row)) = True Then
   pp.relogio 1
   If tpArquivo = "FINASA" Then
      fo_cad_boleto_finasa.bo_codigo.text = CDbl(0 & PegaGrid(gr_cons, 1, gr_cons.row))
      fo_cad_boleto_finasa.Show
   ElseIf tpArquivo = "ESCRITORIOITAU" Then
      fo_cad_boleto_escritorio_itau.bo_codigo.text = CDbl(0 & PegaGrid(gr_cons, 1, gr_cons.row))
      fo_cad_boleto_escritorio_itau.Show
   ElseIf tpArquivo = "ESCRITORIOSANTANDER" Then
      fo_cad_boleto_escritorio_santander.bo_codigo.text = CDbl(0 & PegaGrid(gr_cons, 1, gr_cons.row))
      fo_cad_boleto_escritorio_santander.Show
   Else
      fo_cad_boleto.bo_codigo.text = CDbl(0 & PegaGrid(gr_cons, 1, gr_cons.row))
      fo_cad_boleto.Show
   End If
   pp.relogio 0
End If
End Sub

Private Sub op_itau_cnab400_Click(Value As Integer)
If op_itau_cnab400.Value = True Then
   bt_cnab400.Enabled = True
Else
   bt_cnab400.Enabled = False
End If
End Sub

Private Sub op_itau_webservice_Click(Value As Integer)
If op_itau_webservice.Value = True Then
   bt_json.Enabled = True
Else
   bt_json.Enabled = False
End If
End Sub

Private Sub op_tipo_Click()
bt_baixar.Enabled = False
If op_tipo.valor = "1" Then
   tpArquivo = "ESCRITORIOITAU"
   Tabela = "BoletoEscritorioItau"
   Tabela2 = "CalculoBoletoEscritorioItau"
ElseIf op_tipo.valor = "2" Then
   tpArquivo = "ESCRITORIO"
   Tabela = "Boleto"
   Tabela2 = "CalculoBoleto"
ElseIf op_tipo.valor = "3" Then
   tpArquivo = "ESCRITORIOSANTANDER"
   Tabela = "BoletoEscritorioSantander"
   Tabela2 = "CalculoBoletoEscritorioSantander"
ElseIf op_tipo.valor = "4" Then
   tpArquivo = "FINASA"
   Tabela = "BoletoFinasa"
   Tabela2 = "CalculoBoletoFinasa"
   bt_baixar.Enabled = True
End If
End Sub

Function Digito10(pNumero As String) As Integer
Dim pValTotal As Double
Dim dv As Integer
Dim X As Integer
Dim pFator As Integer
Dim pCalc As String

pValTotal = 0
pFator = 2
For X = Len(pNumero) To 1 Step -1
   pCalc = CInt(CInt(0 & Mid$(pNumero, X, 1)) * pFator)
   If CInt(pCalc) > 9 Then
      pValTotal = pValTotal + CInt(Left$(pCalc, 1)) + CInt(Right$(pCalc, 1))
   Else
      pValTotal = pValTotal + CInt(pCalc)
   End If
   pFator = pFator - 1
   If pFator = 0 Then
      pFator = 2
   End If
Next X

dv = 10 - (pValTotal Mod 10)

If dv = 10 Then
   dv = 0
End If
Digito10 = dv
End Function


