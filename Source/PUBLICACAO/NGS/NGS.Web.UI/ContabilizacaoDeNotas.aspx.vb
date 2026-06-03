Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ContabilizacaoDeNotas
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim Mensagem As String
    Dim Cliente As String
    Dim Campo() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ContabilizacaoDeNotas", "ACESSAR") Then
                VerificaUnidade()
                'ddlUnidade.Focus()
                'VerificaAcao()
                Limpar()
                If carregarContabilizar() Then
                    lnkContabilizar.Enabled = True
                    lnkContabilizar.Parent.Visible = True
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Contabil.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Sub VerificaUnidade()
        'Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
        '    "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
        '    "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
        '    " from Usuarios" & vbCrLf & _
        '    " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        'For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
        '    ddlUnidade.SelectedValue = Dr("AcessoUnidade")
        '    ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        'Next

        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa, True)
    End Sub

    Function carregarContabilizar()
        Dim ds As DataSet
        Dim sql As String = "SELECT Etapa_Id FROM ConfiguracaoXUsuario WHERE Usuario_Id = '" & Session("ssNomeUsuario") & "' AND Etapa_Id = 5"

        ds = Banco.ConsultaDataSet(sql, "ConfiguracaoXUsuario")

        If ds.Tables(0).Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue)
    End Sub

    Public Sub Contabilizar()

        'NÃO REMOVER, USANDO PARA VER QUAL É A NOTA QUANDO DÁ ALGUM ERRO - FURLAN - 27/04/2023
        Dim cCli As String = String.Empty
        Dim nNota As String = String.Empty

        Try
            Dim ds As DataSet

            Sql = "SELECT n.Empresa_Id," & vbCrLf & _
                  "       n.EndEmpresa_Id," & vbCrLf & _
                  "       n.Cliente_Id," & vbCrLf & _
                  "       n.EndCliente_Id, " & vbCrLf & _
                  "       n.EntradaSaida_Id," & vbCrLf & _
                  "       n.Serie_Id," & vbCrLf & _
                  "       n.Nota_Id" & vbCrLf & _
                  "  FROM NotasFiscais n" & vbCrLf & _
                  " INNER JOIN NotasFiscaisXItens ni " & vbCrLf & _
                  "    ON ni.Empresa_Id      = n.Empresa_Id " & vbCrLf & _
                  "   AND ni.EndEmpresa_Id   = n.EndEmpresa_Id " & vbCrLf & _
                  "   AND ni.Cliente_Id      = n.Cliente_Id " & vbCrLf & _
                  "   AND ni.EndCliente_Id   = n.EndCliente_Id " & vbCrLf & _
                  "   AND ni.EntradaSaida_Id = n.EntradaSaida_Id " & vbCrLf & _
                  "   AND ni.Serie_Id        = n.Serie_Id " & vbCrLf & _
                  "   AND ni.Nota_Id         = n.Nota_Id " & vbCrLf & _
                  " INNER JOIN SubOperacoes so " & vbCrLf & _
                  "    ON so.Operacao_Id     = n.Operacao " & vbCrLf & _
                  "   AND so.SubOperacoes_Id = n.SubOperacao " & vbCrLf & _
                  "   AND (so.Contabil = 'S') " & vbCrLf

            If Not Session("SqlContabilizar" & HID.Value) Is Nothing Then
                Sql &= Session("SqlContabilizar" & HID.Value)
            Else
                Sql &= " WHERE (n.Movimento BETWEEN '" & TxtdataInicial.Text.ToSqlDate() & "' AND '" & TxtdataFinal.Text.ToSqlDate() & "')" & vbCrLf & _
                       "   AND (n.Situacao =  1) " & vbCrLf

                If ddlEmpresa.Text <> "" Then
                    Cliente = ddlEmpresa.SelectedValue
                    Campo = Cliente.Split("-")
                    Sql &= " AND (n.Empresa_Id    =  '" & Campo(0) & "')" & vbCrLf & _
                           " AND (n.EndEmpresa_Id = " & Campo(1) & ")" & vbCrLf
                    If txtPedido.Text.Length > 0 Then
                        Sql &= " AND n.Pedido = " & txtPedido.Text & vbCrLf
                    End If
                End If
            End If

            'Sql &= " AND n.Cliente_Id in ('53267147000290','53140701000192','53267147000109','35828454000300','07395369000190','10654834000174','07770042000401')" & vbCrLf

            'Sql &= " AND ni.OperacaoXEstado in(3250) " & vbCrLf
            'Sql &= " AND n.Nota_Id in (10869,10870)" & vbCrLf

            'Sql &= " AND n.Operacao    = 69 " & vbCrLf
            'Sql &= " AND n.SubOperacao = 6 " & vbCrLf
            'Sql &= " AND n.Nota_Id in (153569,153593,153594,9730,9726,40403,40430,40398,40396,1228839,1228708,1228485,1228482,17980,17983,17985,17986,17987,17988,17989,17990,17991,17992,17993,17994,17995,17997,153627, " & vbCrLf
            'Sql &= " 153640,153949,154134,154142,154145,154163,154168,154180,154166,154210,154539,154557,154588,154494,154495,154615,9666,9645,9644,9639,9665,9657,9655,9652,9650,9668,9669,9671,9672,9681,9683,9684,9686, " & vbCrLf
            'Sql &= " 9688,9689,9732,9733,9734,9735,9737,9738,9739,9741,9742,9745,9746,9747,9757,9759,9763,9764,9767,9770,9772,9773,9775,9776,9777,9778,9905,9906,9909,9793,9794,9796,9792,9790,9808,1232670,1232780,1232654, " & vbCrLf
            'Sql &= " 1233071,1233106,1232338,1232457,1232579,1232642,1232643,1232650,154509) " & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Notas")

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Não existe(m) registro(s) para Contabilização no período informado.", eTitulo.Info)
                Exit Sub
            End If

            Dim Sqls As New ArrayList

            ChkResultado.Items.Clear()

            For Each row As DataRow In ds.Tables(0).Rows

                cCli = row("Cliente_Id")
                nNota = row("Nota_Id")

                'If cCli = "08592950000164" And nNota = "25107" Then
                '    Dim TESTE As String = ""
                'End If

                Dim nf As New [Lib].Negocio.NotaFiscal()
                nf.CodigoEmpresa = row("Empresa_Id")
                nf.EnderecoEmpresa = row("EndEmpresa_Id")
                nf.CodigoCliente = row("Cliente_Id")
                nf.EnderecoCliente = row("EndCliente_Id")
                nf.EntradaSaida = IIf(row("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                nf.Codigo = row("Nota_Id")
                nf.Serie = row("Serie_Id")
                nf = New [Lib].Negocio.NotaFiscal(nf)

                'If Funcoes.VerificaAcesso(nf.CodigoEmpresa, nf.EnderecoEmpresa, nf.Movimento.ToString("dd-MM-yyyy"), "CONTABIL") Then
                If nf.NFG Then
                    Dim cLote As Integer = 0
                    If nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) Or _
                        nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ReciboDeFrete) Or _
                        nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Or _
                        nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) Or _
                        nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_TOM) Or _
                        nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) Or _
                        nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_OUT) Or _
                        nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Anulacao) Then
                        cLote = 21
                    Else
                        cLote = 9
                    End If

                    nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
                    nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
                    nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
                    nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
                    nf.Razao.ContabilizarNotaSql(Sqls, cLote)
                Else
                    If nf.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.CTRC) _
                       AndAlso nf.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.ReciboDeFrete) _
                       AndAlso nf.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Estadia) _
                       AndAlso nf.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.CTRC_SEM_NF) _
                       AndAlso nf.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.CT_E_TOM) _
                       AndAlso nf.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.CT_E) _
                       AndAlso nf.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.CT_E_OUT) _
                       AndAlso nf.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Anulacao) Then

                        nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
                        nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
                        nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
                        nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
                        nf.Razao.ContabilizarNotaSql(Sqls, 10)
                    Else
                        nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
                        nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
                        nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
                        nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
                        nf.Razao.ContabilizarNotaSql(Sqls, 10)
                    End If
                End If
                'Else
                '    ChkResultado.Items.Add(New ListItem("Movimento Contábil já Fechado para esta data " & nf.Movimento.ToString("dd-MM-yyyy")))
                'End If
            Next

            If Banco.GravaBanco(Sqls) Then
                ChkResultado.Items.Add(New ListItem(" FIM - Processo de Contabilização realizado com Sucesso"))

                MsgBox(Me.Page, "Processo de Contabilização realizado com Sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Info, True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message & " - Cliente " & cCli & " - NOTA " & nNota, eTitulo.Info, True)
        End Try
    End Sub

    'Public Sub VerificaAcao()
    '    If RdNota.Checked Then
    '        Label1.Text = "Unidade:"
    '        Label2.Text = "Empresa:"
    '        divNota.Visible = False
    '        divNota2.Visible = False
    '    Else
    '        divNota.Visible = True
    '        divNota2.Visible = True
    '        Label1.Text = "Un.Transp.:"
    '        Label2.Text = "Transportadora:"
    '    End If
    'End Sub

    Private Sub Limpar()
        lnkContabilizar.Enabled = False
        lnkContabilizar.Parent.Visible = False
        TxtdataInicial.Text = Format(Today, "dd/MM/yyyy")
        TxtdataFinal.Text = Format(Today, "dd/MM/yyyy")

        ChkResultado.Items.Clear()

        HID.Value = Guid.NewGuid().ToString
        ucContabilizarNotas.SetarHID(HID.Value)

        LiberaEmpresa()
        'ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "")
        'ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "")
        'ddl.Carregar(ddlUnidadeContratante, CarregarDDL.Tabela.UnidadeDeNegocio, "")
        'ddl.Carregar(ddlEmpresaContratante, CarregarDDL.Tabela.Empresas, "")
        'If RdNota.Checked Then
        '    Label1.Text = "Unidade:"
        '    Label2.Text = "Empresa:"
        '    divNota.Visible = False
        '    divNota2.Visible = False
        'Else
        '    divNota.Visible = True
        '    divNota2.Visible = True
        '    Label1.Text = "Un.Transp.:"
        '    Label2.Text = "Transportadora:"
        'End If
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    'Protected Sub RdNota_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    VerificaAcao()
    'End Sub

    'Protected Sub rdFrete_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    VerificaAcao()
    'End Sub

    Protected Sub ddlUnidadeContratante_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(ddlEmpresaContratante, CarregarDDL.Tabela.Empresas, ddlUnidadeContratante.SelectedValue)
    End Sub

    Protected Sub lnkContabilizar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkContabilizar.Click
        ucContabilizarNotas.Limpar()
        Popup.ContabilizarNotas(Me.Page, "objContabilizar" & HID.Value)
    End Sub

    Protected Sub lnkProcessar_Click(sender As Object, e As EventArgs) Handles lnkProcessar.Click
        'If RdNota.Checked Then

        If Session("ssNomeUsuario") = "FURLAN" OrElse Session("ssNomeUsuario") = "KAYNÃ" Then
            Contabilizar()
        ElseIf ddlEmpresa.Text = "" Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
        ElseIf Not CDate(TxtdataInicial.Text).Year = CDate(TxtdataFinal.Text).Year Then
            MsgBox(Me.Page, "Ano da recontabilização não pode ser diferente.")
        ElseIf Not CDate(TxtdataInicial.Text).Month = CDate(TxtdataFinal.Text).Month Then
            MsgBox(Me.Page, "Mês da recontabilização não pode ser diferente.")
        ElseIf CDate(TxtdataFinal.Text).Subtract(CDate(TxtdataInicial.Text)).Days > 10 Then
            MsgBox(Me.Page, "`Pro ser um processo longo, favor recontabilizar no máximo de 10 em 10 dias.")
        Else
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            If Empresa(0) = "04854422000347" Then
                MsgBox(Me.Page, "Empresa não pode ser contabilizada.")
            Else
                Contabilizar()
            End If
        End If
        'Else
        '    Dim strTransportadora As String() = ddlEmpresa.SelectedValue.Split("-")
        '    Dim strContratante As String() = ddlEmpresaContratante.SelectedValue.Split("-")

        '    Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal()
        '    objNotaFiscal.CodigoEmpresa = strTransportadora(0)
        '    objNotaFiscal.EnderecoEmpresa = strTransportadora(1)
        '    If strContratante.Length > 1 Then
        '        objNotaFiscal.CodigoCliente = Trim(strContratante(0))
        '        objNotaFiscal.EnderecoCliente = Trim(strContratante(1))
        '    End If
        '    objNotaFiscal.Serie = "REC"
        '    objNotaFiscal.EntradaSaida = eEntradaSaida.Saida
        '    objNotaFiscal.Codigo = 1

        '    Dim razao As New [Lib].Negocio.Razao
        '    If razao.TransferirFretesNoRazao(strTransportadora(0), strTransportadora(1), TxtdataInicial.Text.ToSqlDate(), TxtdataFinal.Text.ToSqlDate(), objNotaFiscal.CodigoCliente, objNotaFiscal.EnderecoCliente, objNotaFiscal) Then
        '        MsgBox(Me.Page, "Contabilização Executada Com Sucesso.", eTitulo.Sucess)
        '    Else
        '        MsgBox(Me.Page, "Erro Durante o Processo de Contabilização")
        '    End If
        'End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ContabilizacaoDeNotas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class