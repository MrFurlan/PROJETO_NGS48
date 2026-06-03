Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaRelatorio
    Inherits BaseUserControl

#Region "Propriedades"

    Public Property MyParameters() As Dictionary(Of String, Object)
        Get
            Return CType(Session("MyParameters" & HID.Value), Dictionary(Of String, Object))
        End Get
        Set(ByVal value As Dictionary(Of String, Object))
            Session("MyParameters" & HID.Value) = value
        End Set
    End Property

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Limpar()
        End If
    End Sub

    Protected Sub rdoRodoviario_CheckedChanged(sender As Object, e As EventArgs) Handles rdoRodoviario.CheckedChanged
        habilitarCampos()
    End Sub

    Protected Sub rdoFerroviario_CheckedChanged(sender As Object, e As EventArgs) Handles rdoFerroviario.CheckedChanged
        habilitarCampos()
    End Sub

    Protected Sub rdoRodovXFerrov_CheckedChanged(sender As Object, e As EventArgs) Handles rdoRodovXFerrov.CheckedChanged
        habilitarCampos()
    End Sub

    Protected Sub btnGerar_Click(sender As Object, e As EventArgs) Handles btnGerar.Click
        Dim ds As New DataSet
        If rdoFerroviario.Checked AndAlso rdoPorNota.Checked Then
            ds = getDataSetByNota()
        ElseIf rdoFerroviario.Checked AndAlso rdoPorCtrc.Checked Then
            ds = getDataSetByCtrc()
        ElseIf rdoRodoviario.Checked Then
            ds = getDataSetRodoviario()
        ElseIf rdoRodovXFerrov.Checked Then
            ds = getDataSetRodoviarioXFerroviario()
        End If

        Dim rptName As String = ""
        If rdoFerroviario.Checked AndAlso rdoPorNota.Checked Then
            rptName = "Cr_CtrcFerroviarioPorNota"
        ElseIf rdoFerroviario.Checked AndAlso rdoPorCtrc.Checked Then
            rptName = "Cr_NotasFerroviarioPorCtrc"
        ElseIf rdoRodoviario.Checked Then
            rptName = "Cr_CtrcRodoviario"
        ElseIf rdoRodovXFerrov.Checked Then
            rptName = "Cr_CtrcRodoviarioXFerroviario"
        End If

        Dim parametrosConsulta As String = MyParameters("ParametrosConsulta")

        Dim param As New Dictionary(Of String, Object)()
        param.Add("parametros", parametrosConsulta)

        Funcoes.BindReport(Me.Page, ds, rptName, eExportType.PDF, param)

    End Sub

    Protected Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Popup.CloseDialog(Me.Page, "divConsultaRelatorio")
    End Sub

#End Region

#Region "Métodos"

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        rdoRodoviario.Checked = True
        rdoPorNota.Checked = True
        habilitarCampos()
    End Sub

    Public Sub pageLoad()
        rdoRodoviario.Checked = MyParameters("viaDeTransporte") = "R"
        rdoFerroviario.Checked = MyParameters("viaDeTransporte") = "F"
        habilitarCampos()
    End Sub

    Private Function getDataSetByCtrc() As DataSet
        Dim ds As New DataSet()
        Dim sql As String = "SELECT	                                                                                                         " & vbCrLf & _
        "		--CONHECIMENTO DE TRANSPORTE                                                                                                         " & vbCrLf & _
        "		cast(NotaFiscalReferencial.NotaReferencial_Id as varchar) + '-' + cast(NotaFiscalReferencial.SerieReferencial_Id as varchar) as Ctrc,                      " & vbCrLf & _
        "		NotaFiscalReferencial.EntradaSaidaReferencial_Id as Ctrc_EntradaSaida,                                                                          " & vbCrLf & _
        "		NotaFiscalReferencial.ClienteReferencial_Id + '-' + cast(NotaFiscalReferencial.EndClienteReferencial_Id as varchar) + '-' + cli2.Nome as Ctrc_Cliente,     " & vbCrLf & _
        "		cast(nf2.Operacao as varchar) + '-' + cast(nf2.SubOperacao as varchar) as Ctrc_Operacao,                                             " & vbCrLf & _
        "		nfer.Data as Ctrc_Data,                                                                                                              " & vbCrLf & _
        "		ISNULL(nxi2.Valor, 0) as Ctrc_ValorOficial,                                                                                          " & vbCrLf & _
        "		ISNULL(nxi2.QuantidadeFiscal,0) as Ctrc_PesoFiscal,                                                                                  " & vbCrLf & _
        "		0 as Ctrc_Peso,                                                                                                                      " & vbCrLf & _
        "		0 as Ctrc_Valor,                                                                                                                     " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "		--NOTA FISCAL                                                                                                                        " & vbCrLf & _
        "        cast(NotasFiscais.Nota_Id as varchar) + '-' + cast(NotasFiscais.Serie_Id as varchar) as Nota,                                       " & vbCrLf & _
        "		NotasFiscais.EntradaSaida_Id as EntradaSaida,                                                                                        " & vbCrLf & _
        "		NotasFiscais.Cliente_Id + '-' + cast(NotasFiscais.EndCliente_Id as varchar) + '-' + Clientes.Nome as Cliente,                        " & vbCrLf & _
        "		cast(NotasFiscais.Operacao as varchar) + '-' + cast(NotasFiscais.SubOperacao as varchar) as Operacao,                                " & vbCrLf & _
        "		NFERealizadas.Data,                                                                                                                  " & vbCrLf & _
        "		cast(((NotaFiscalReferencial.Quantidade * nxi2.Unitario) / 1000) as decimal(18,4)) as ValorOficial,                                  " & vbCrLf & _
        "		ISNULL(NotaFiscalReferencial.Quantidade,0) as PesoFiscal,                                                                                      " & vbCrLf & _
        "		0 as ConsumidoPeso,                                                                                                                  " & vbCrLf & _
        "		0 as PesoSaldo                                                                                                                       " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "	FROM NotaFiscalReferencial                                                                                                               " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "	INNER JOIN NotasFiscaisXItens nxi2                                                                                                       " & vbCrLf & _
        "		ON nxi2.Empresa_Id       = NotaFiscalReferencial.EmpresaReferencial_Id                                                               " & vbCrLf & _
        "		AND nxi2.EndEmpresa_Id   = NotaFiscalReferencial.EndEmpresaReferencial_Id                                                            " & vbCrLf & _
        "		AND nxi2.Cliente_Id      = NotaFiscalReferencial.ClienteReferencial_Id                                                               " & vbCrLf & _
        "		AND nxi2.EndCliente_Id   = NotaFiscalReferencial.EndClienteReferencial_Id                                                            " & vbCrLf & _
        "		AND nxi2.EntradaSaida_Id = NotaFiscalReferencial.EntradaSaidaReferencial_Id                                                          " & vbCrLf & _
        "		AND nxi2.Nota_Id         = NotaFiscalReferencial.NotaReferencial_Id                                                                  " & vbCrLf & _
        "		AND nxi2.Serie_Id        = NotaFiscalReferencial.SerieReferencial_Id                                                                 " & vbCrLf & _
        "		AND nxi2.Produto_Id      = NotaFiscalReferencial.ProdutoReferencial_Id                                                               " & vbCrLf & _
        "		AND nxi2.CFOP_Id         = NotaFiscalReferencial.CFOPReferencial_Id                                                                  " & vbCrLf & _
        "		AND nxi2.Sequencia_Id    = NotaFiscalReferencial.SequenciaReferencial_Id                                                             " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "	INNER JOIN NotasFiscais nf2                                                                                                              " & vbCrLf & _
        "	    ON nxi2.Empresa_Id       = nf2.Empresa_Id                                                                                            " & vbCrLf & _
        "		AND nxi2.EndEmpresa_Id   = nf2.EndEmpresa_Id                                                                                         " & vbCrLf & _
        "		AND nxi2.Cliente_Id      = nf2.Cliente_Id                                                                                            " & vbCrLf & _
        "		AND nxi2.EndCliente_Id   = nf2.EndCliente_Id                                                                                         " & vbCrLf & _
        "		AND nxi2.EntradaSaida_Id = nf2.EntradaSaida_Id                                                                                       " & vbCrLf & _
        "		AND nxi2.Serie_Id        = nf2.Serie_Id                                                                                              " & vbCrLf & _
        "		AND nxi2.Nota_Id         = nf2.Nota_Id                                                                                               " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "	INNER JOIN NotasFiscaisXItens                                                                                                            " & vbCrLf & _
        "		ON NotasFiscaisXItens.Empresa_Id       = NotaFiscalReferencial.Empresa_Id                                                            " & vbCrLf & _
        "		AND NotasFiscaisXItens.EndEmpresa_Id   = NotaFiscalReferencial.EndEmpresa_Id                                                         " & vbCrLf & _
        "		AND NotasFiscaisXItens.Cliente_Id      = NotaFiscalReferencial.Cliente_Id                                                            " & vbCrLf & _
        "		AND NotasFiscaisXItens.EndCliente_Id   = NotaFiscalReferencial.EndCliente_Id                                                         " & vbCrLf & _
        "		AND NotasFiscaisXItens.EntradaSaida_Id = NotaFiscalReferencial.EntradaSaida_Id                                                       " & vbCrLf & _
        "		AND NotasFiscaisXItens.Nota_Id         = NotaFiscalReferencial.Nota_Id                                                               " & vbCrLf & _
        "		AND NotasFiscaisXItens.Serie_Id        = NotaFiscalReferencial.Serie_Id                                                              " & vbCrLf & _
        "		AND NotasFiscaisXItens.Produto_Id      = NotaFiscalReferencial.Produto_Id                                                            " & vbCrLf & _
        "		AND NotasFiscaisXItens.CFOP_Id         = NotaFiscalReferencial.CFOP_Id                                                               " & vbCrLf & _
        "		AND NotasFiscaisXItens.Sequencia_Id    = NotaFiscalReferencial.Sequencia_Id                                                          " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "	INNER JOIN NotasFiscais                                                                                                                  " & vbCrLf & _
        "	    ON NotasFiscaisXItens.Empresa_Id       = NotasFiscais.Empresa_Id                                                                     " & vbCrLf & _
        "		AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id                                                                  " & vbCrLf & _
        "		AND NotasFiscaisXItens.Cliente_Id      = NotasFiscais.Cliente_Id                                                                     " & vbCrLf & _
        "		AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscais.EndCliente_Id                                                                  " & vbCrLf & _
        "		AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id                                                                " & vbCrLf & _
        "		AND NotasFiscaisXItens.Serie_Id        = NotasFiscais.Serie_Id                                                                       " & vbCrLf & _
        "		AND NotasFiscaisXItens.Nota_Id         = NotasFiscais.Nota_Id	                                                                     " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "     INNER JOIN (SELECT Empresa_Id,                                                                                                         " & vbCrLf & _
        "						EndEmpresa_Id,                                                                                                       " & vbCrLf & _
        "						Cliente_Id,                                                                                                          " & vbCrLf & _
        "						EndCliente_Id,                                                                                                       " & vbCrLf & _
        "						EntradaSaida_Id,                                                                                                     " & vbCrLf & _
        "						Serie_Id,                                                                                                            " & vbCrLf & _
        "						Nota_Id,                                                                                                             " & vbCrLf & _
        "						SUM(QuantidadeFiscal) AS QuantidadeFiscal,                                                                           " & vbCrLf & _
        "						SUM(Valor) AS Valor                                                                                                  " & vbCrLf & _
        "					FROM NotasFiscaisXItens                                                                                                  " & vbCrLf & _
        "					GROUP BY Empresa_id,                                                                                                     " & vbCrLf & _
        "					EndEmpresa_id,                                                                                                           " & vbCrLf & _
        "					Cliente_id,                                                                                                              " & vbCrLf & _
        "					EndCliente_id,                                                                                                           " & vbCrLf & _
        "					EntradaSaida_id,                                                                                                         " & vbCrLf & _
        "					Serie_id,                                                                                                                " & vbCrLf & _
        "					Nota_id                                                                                                                  " & vbCrLf & _
        "				) As NFI                                                                                                                     " & vbCrLf & _
        "		on NFI.Empresa_Id      = NotasFiscais.Empresa_Id                                                                                     " & vbCrLf & _
        "		AND NFI.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id                                                                                 " & vbCrLf & _
        "		AND NFI.Cliente_Id      = NotasFiscais.Cliente_Id                                                                                    " & vbCrLf & _
        "		AND NFI.EndCliente_Id   = NotasFiscais.EndCliente_Id                                                                                 " & vbCrLf & _
        "		AND NFI.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id                                                                               " & vbCrLf & _
        "		AND NFI.Serie_Id        = NotasFiscais.Serie_Id                                                                                      " & vbCrLf & _
        "		AND NFI.Nota_Id         = NotasFiscais.Nota_Id                                                                                       " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "    INNER JOIN Clientes                                                                                                                     " & vbCrLf & _
        "		ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id                                                                                  " & vbCrLf & _
        "		AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id                                                                                " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "	INNER JOIN Clientes cli2                                                                                                                 " & vbCrLf & _
        "		ON NotaFiscalReferencial.Cliente_Id    = cli2.Cliente_Id                                                                             " & vbCrLf & _
        "		AND NotaFiscalReferencial.EndCliente_Id = cli2.Endereco_Id                                                                           " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "	LEFT JOIN NFERealizadas                                                                                                                  " & vbCrLf & _
        "		ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id                                                                           " & vbCrLf & _
        "		AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id                                                                       " & vbCrLf & _
        "		AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id                                                                          " & vbCrLf & _
        "		AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id                                                                       " & vbCrLf & _
        "		AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id                                                                     " & vbCrLf & _
        "		AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id                                                                            " & vbCrLf & _
        "		AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id                                                                             " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "	LEFT JOIN NFERealizadas nfer                                                                                                             " & vbCrLf & _
        "		ON nf2.Empresa_Id      = nfer.Empresa_Id                                                                                             " & vbCrLf & _
        "		AND nf2.EndEmpresa_Id   = nfer.EndEmpresa_Id                                                                                         " & vbCrLf & _
        "		AND nf2.Cliente_Id      = nfer.Cliente_Id                                                                                            " & vbCrLf & _
        "		AND nf2.EndCliente_Id   = nfer.EndCliente_Id                                                                                         " & vbCrLf & _
        "		AND nf2.EntradaSaida_Id = nfer.EntradaSaida_Id                                                                                       " & vbCrLf & _
        "		AND nf2.Serie_Id        = nfer.Serie_Id                                                                                              " & vbCrLf & _
        "		AND nf2.Nota_Id         = nfer.Nota_Id		                                                                                         " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "	LEFT JOIN (SELECT	nxi.Empresa_Id,                                                                                                      " & vbCrLf & _
        "						nxi.EndEmpresa_Id,                                                                                                   " & vbCrLf & _
        "						nxi.Cliente_Id,                                                                                                      " & vbCrLf & _
        "						nxi.EndCliente_Id,                                                                                                   " & vbCrLf & _
        "						nxi.EntradaSaida_Id,                                                                                                 " & vbCrLf & _
        "						nxi.Serie_Id,                                                                                                        " & vbCrLf & _
        "						nxi.Nota_Id,                                                                                                         " & vbCrLf & _
        "						SUM(nfr.Quantidade) AS QuantidadeFiscalConsumido,                                                                    " & vbCrLf & _
        "						SUM(nfr.Valor) AS ValorConsumido                                                                                     " & vbCrLf & _
        "				FROM NotasFiscaisXItens nxi                                                                                                  " & vbCrLf & _
        "					INNER JOIN NotaFiscalReferencial nfr                                                                                     " & vbCrLf & _
        "						ON nxi.Empresa_Id = nfr.Empresa_Id                                                                                   " & vbCrLf & _
        "						AND nxi.EndEmpresa_Id = nfr.EndEmpresa_Id                                                                            " & vbCrLf & _
        "						AND nxi.Cliente_Id = nfr.Cliente_Id                                                                                  " & vbCrLf & _
        "						AND nxi.EndCliente_Id = nfr.EndCliente_Id                                                                            " & vbCrLf & _
        "						AND nxi.EntradaSaida_Id = nfr.EntradaSaida_Id                                                                        " & vbCrLf & _
        "						AND nxi.Nota_Id = nfr.Nota_Id                                                                                        " & vbCrLf & _
        "						AND nxi.Serie_Id = nfr.Serie_Id                                                                                      " & vbCrLf & _
        "						AND nxi.Produto_Id = nfr.Produto_Id                                                                                  " & vbCrLf & _
        "						AND nxi.CFOP_Id = nfr.CFOP_Id                                                                                        " & vbCrLf & _
        "						AND nxi.Sequencia_Id = nfr.Sequencia_Id                                                                              " & vbCrLf & _
        "				GROUP BY nxi.Empresa_id,                                                                                                     " & vbCrLf & _
        "				nxi.EndEmpresa_id,                                                                                                           " & vbCrLf & _
        "				nxi.Cliente_id,                                                                                                              " & vbCrLf & _
        "				nxi.EndCliente_id,                                                                                                           " & vbCrLf & _
        "				nxi.EntradaSaida_id,                                                                                                         " & vbCrLf & _
        "				nxi.Serie_id,                                                                                                                " & vbCrLf & _
        "				nxi.Nota_id                                                                                                                  " & vbCrLf & _
        "				) As Consumido                                                                                                               " & vbCrLf & _
        "		on Consumido.Empresa_Id      = NotasFiscais.Empresa_Id                                                                               " & vbCrLf & _
        "		AND Consumido.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id                                                                           " & vbCrLf & _
        "		AND Consumido.Cliente_Id      = NotasFiscais.Cliente_Id                                                                              " & vbCrLf & _
        "		AND Consumido.EndCliente_Id   = NotasFiscais.EndCliente_Id                                                                           " & vbCrLf & _
        "		AND Consumido.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id                                                                         " & vbCrLf & _
        "		AND Consumido.Serie_Id        = NotasFiscais.Serie_Id                                                                                " & vbCrLf & _
        "		AND Consumido.Nota_Id         = NotasFiscais.Nota_Id                                                                                 " & vbCrLf & _
        "                                                                                                                                            " & vbCrLf & _
        "    LEFT JOIN NotaFiscalDevolucaoXNotaFiscal Dev                                                                                            " & vbCrLf & _
        "		ON Dev.EmpresaDevolucao_Id	 = NotasFiscais.Empresa_Id                                                                               " & vbCrLf & _
        "		AND Dev.EndEmpresaDevolucao_Id = NotasFiscais.EndEmpresa_Id                                                                          " & vbCrLf & _
        "		AND Dev.ClienteDevolucao_Id	 = NotasFiscais.Cliente_Id                                                                               " & vbCrLf & _
        "		AND Dev.EndClienteDevolucao_Id = NotasFiscais.EndCliente_Id                                                                          " & vbCrLf & _
        "		AND Dev.Nota_Id			     = NotasFiscais.Nota_Id                                                                                  " & vbCrLf & _
        "		AND Dev.Serie_Id		         = NotasFiscais.Serie_Id                                                                             " & vbCrLf & _
        "		AND Dev.EntradaSaida_Id        = NotasFiscais.EntradaSaida_Id                                                                        " & vbCrLf & _
        "                                                                                                                                                                " & vbCrLf & _
        "  WHERE 1=1                                                                                                              " & vbCrLf & _
        "       AND NotasFiscais.Empresa_Id = '" & MyParameters("empresa") & "'" & vbCrLf & _
        "       AND NotasFiscais.EndEmpresa_Id = " & MyParameters("enderecoEmpresa") & vbCrLf & _
        "       AND NotasFiscais.EntradaSaida_Id = '" & MyParameters("es") & "'" & vbCrLf & _
        "       AND NotasFiscais.Situacao = 1 " & vbCrLf & _
        "       AND NotasFiscais.CIFFOB = '" & MyParameters("ciffob") & "'" & vbCrLf

        If Not String.IsNullOrWhiteSpace(MyParameters("codcliente")) Then
            sql &= "  AND NotasFiscais.Cliente_Id = '" & MyParameters("codcliente").Split("-")(0) & "'" & vbCrLf
            sql &= "  AND NotasFiscais.EndCliente_Id = " & MyParameters("codcliente").Split("-")(1) & "" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(MyParameters("numconhecimento")) Then
            Dim strNota() As String = MyParameters("numconhecimento").Trim().Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
            If strNota IsNot Nothing AndAlso strNota.Length > 1 Then
                sql &= "  AND NotasFiscais.Nota_Id IN (" & String.Join(",", strNota) & ") " & vbCrLf
            Else
                sql &= "  AND NotasFiscais.Nota_Id IN (" & strNota(0) & ") " & vbCrLf
            End If
        End If

        sql &= "  AND nf2.Movimento between '" & MyParameters("data1").tosqldate() & "' and '" & MyParameters("data2").tosqldate() & "'" & vbCrLf & _
            "   ORDER BY NotaFiscalReferencial.Nota_Id, NotaFiscalReferencial.Serie_Id, nfer.Data, " & vbCrLf & _
            "            NotasFiscais.Nota_Id, NotasFiscais.Serie_Id, NFERealizadas.Data DESC " & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "CtrcFerroviario")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim cnpj As String = Funcoes.FormatarCpfCnpj(row("Ctrc_Cliente").ToString().Split("-")(0))
            Dim nome As String = row("Ctrc_Cliente").ToString().Substring(row("Ctrc_Cliente").ToString().IndexOf("-"), row("Ctrc_Cliente").ToString().Length - row("Ctrc_Cliente").ToString().IndexOf("-"))
            row("Ctrc_Cliente") = String.Format("{0} {1}", cnpj, nome)

            cnpj = Funcoes.FormatarCpfCnpj(row("Cliente").ToString().Split("-")(0))
            nome = row("Cliente").ToString().Substring(row("Cliente").ToString().IndexOf("-"), row("Cliente").ToString().Length - row("Cliente").ToString().IndexOf("-"))
            row("Cliente") = String.Format("{0} {1}", cnpj, nome)
        Next

        Return ds
    End Function

    Private Function getDataSetByNota() As DataSet
        Dim ds As New DataSet()

        Dim sql As String = "SELECT	                                                                                                                                                         " & vbCrLf & _
            "		--CONHECIMENTO DE TRANSPORTE                                                                                                                             " & vbCrLf & _
            "		cast(NotaFiscalReferencial.NotaReferencial_Id as varchar) + '-' + cast(NotaFiscalReferencial.SerieReferencial_Id as varchar) as Ctrc,                    " & vbCrLf & _
            "		NotaFiscalReferencial.EntradaSaidaReferencial_Id as Ctrc_EntradaSaida,                                                                                   " & vbCrLf & _
            "		NotaFiscalReferencial.ClienteReferencial_Id + '-' + cast(NotaFiscalReferencial.EndClienteReferencial_Id as varchar) + '-' + cli2.Nome as Ctrc_Cliente,   " & vbCrLf & _
            "		cast(nf2.Operacao as varchar) + '-' + cast(nf2.SubOperacao as varchar) as Ctrc_Operacao,                                                                 " & vbCrLf & _
            "		nfer.Data as Ctrc_Data,                                                                                                                                  " & vbCrLf & _
            "		ISNULL(nxi2.Valor, 0) as Ctrc_ValorOficial,                                                                                                              " & vbCrLf & _
            "		ISNULL(nxi2.QuantidadeFiscal,0) as Ctrc_PesoFiscal,                                                                                                      " & vbCrLf & _
            "		ISNULL(NotaFiscalReferencial.Quantidade,0) as Ctrc_Peso,                                                                                                           " & vbCrLf & _
            "		ISNULL(NotaFiscalReferencial.Valor,0) as Ctrc_Valor,                                                                                                               " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "		--NOTA FISCAL                                                                                                                                            " & vbCrLf & _
            "        cast(NotasFiscais.Nota_Id as varchar) + '-' + cast(NotasFiscais.Serie_Id as varchar) as Nota,                                                           " & vbCrLf & _
            "		NotasFiscais.EntradaSaida_Id as EntradaSaida,                                                                                                            " & vbCrLf & _
            "		NotasFiscais.Cliente_Id + '-' + cast(NotasFiscais.EndCliente_Id as varchar) + '-' + Clientes.Nome as Cliente,                                            " & vbCrLf & _
            "		cast(NotasFiscais.Operacao as varchar) + '-' + cast(NotasFiscais.SubOperacao as varchar) as Operacao,                                                    " & vbCrLf & _
            "		NFERealizadas.Data,                                                                                                                                      " & vbCrLf & _
            "		ISNULL(NFI.Valor, 0) as ValorOficial,                                                                                                                    " & vbCrLf & _
            "		ISNULL(NFI.QuantidadeFiscal,0) as PesoFiscal,                                                                                                            " & vbCrLf & _
            "		ISNULL(consumido.QuantidadeFiscalConsumido,0) as ConsumidoPeso,                                                                                         " & vbCrLf & _
            "		ISNULL(NFI.QuantidadeFiscal,0) - ISNULL(consumido.QuantidadeFiscalConsumido ,0) as PesoSaldo                                                             " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "	FROM NotaFiscalReferencial                                                                                                                                   " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "	INNER JOIN NotasFiscaisXItens nxi2                                                                                                                           " & vbCrLf & _
            "		ON nxi2.Empresa_Id       = NotaFiscalReferencial.EmpresaReferencial_Id                                                                                   " & vbCrLf & _
            "		AND nxi2.EndEmpresa_Id   = NotaFiscalReferencial.EndEmpresaReferencial_Id                                                                                " & vbCrLf & _
            "		AND nxi2.Cliente_Id      = NotaFiscalReferencial.ClienteReferencial_Id                                                                                   " & vbCrLf & _
            "		AND nxi2.EndCliente_Id   = NotaFiscalReferencial.EndClienteReferencial_Id                                                                                " & vbCrLf & _
            "		AND nxi2.EntradaSaida_Id = NotaFiscalReferencial.EntradaSaidaReferencial_Id                                                                              " & vbCrLf & _
            "		AND nxi2.Nota_Id         = NotaFiscalReferencial.NotaReferencial_Id                                                                                      " & vbCrLf & _
            "		AND nxi2.Serie_Id        = NotaFiscalReferencial.SerieReferencial_Id                                                                                     " & vbCrLf & _
            "		AND nxi2.Produto_Id      = NotaFiscalReferencial.ProdutoReferencial_Id                                                                                   " & vbCrLf & _
            "		AND nxi2.CFOP_Id         = NotaFiscalReferencial.CFOPReferencial_Id                                                                                      " & vbCrLf & _
            "		AND nxi2.Sequencia_Id    = NotaFiscalReferencial.SequenciaReferencial_Id                                                                                 " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "	INNER JOIN NotasFiscais nf2                                                                                                                                  " & vbCrLf & _
            "	    ON nxi2.Empresa_Id       = nf2.Empresa_Id                                                                                                                " & vbCrLf & _
            "		AND nxi2.EndEmpresa_Id   = nf2.EndEmpresa_Id                                                                                                             " & vbCrLf & _
            "		AND nxi2.Cliente_Id      = nf2.Cliente_Id                                                                                                                " & vbCrLf & _
            "		AND nxi2.EndCliente_Id   = nf2.EndCliente_Id                                                                                                             " & vbCrLf & _
            "		AND nxi2.EntradaSaida_Id = nf2.EntradaSaida_Id                                                                                                           " & vbCrLf & _
            "		AND nxi2.Serie_Id        = nf2.Serie_Id                                                                                                                  " & vbCrLf & _
            "		AND nxi2.Nota_Id         = nf2.Nota_Id                                                                                                                   " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "	INNER JOIN NotasFiscaisXItens                                                                                                                                " & vbCrLf & _
            "		ON NotasFiscaisXItens.Empresa_Id       = NotaFiscalReferencial.Empresa_Id                                                                                " & vbCrLf & _
            "		AND NotasFiscaisXItens.EndEmpresa_Id   = NotaFiscalReferencial.EndEmpresa_Id                                                                             " & vbCrLf & _
            "		AND NotasFiscaisXItens.Cliente_Id      = NotaFiscalReferencial.Cliente_Id                                                                                " & vbCrLf & _
            "		AND NotasFiscaisXItens.EndCliente_Id   = NotaFiscalReferencial.EndCliente_Id                                                                             " & vbCrLf & _
            "		AND NotasFiscaisXItens.EntradaSaida_Id = NotaFiscalReferencial.EntradaSaida_Id                                                                           " & vbCrLf & _
            "		AND NotasFiscaisXItens.Nota_Id         = NotaFiscalReferencial.Nota_Id                                                                                   " & vbCrLf & _
            "		AND NotasFiscaisXItens.Serie_Id        = NotaFiscalReferencial.Serie_Id                                                                                  " & vbCrLf & _
            "		AND NotasFiscaisXItens.Produto_Id      = NotaFiscalReferencial.Produto_Id                                                                                " & vbCrLf & _
            "		AND NotasFiscaisXItens.CFOP_Id         = NotaFiscalReferencial.CFOP_Id                                                                                   " & vbCrLf & _
            "		AND NotasFiscaisXItens.Sequencia_Id    = NotaFiscalReferencial.Sequencia_Id                                                                              " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "	INNER JOIN NotasFiscais                                                                                                                                      " & vbCrLf & _
            "	    ON NotasFiscaisXItens.Empresa_Id       = NotasFiscais.Empresa_Id                                                                                         " & vbCrLf & _
            "		AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id                                                                                      " & vbCrLf & _
            "		AND NotasFiscaisXItens.Cliente_Id      = NotasFiscais.Cliente_Id                                                                                         " & vbCrLf & _
            "		AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscais.EndCliente_Id                                                                                      " & vbCrLf & _
            "		AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id                                                                                    " & vbCrLf & _
            "		AND NotasFiscaisXItens.Serie_Id        = NotasFiscais.Serie_Id                                                                                           " & vbCrLf & _
            "		AND NotasFiscaisXItens.Nota_Id         = NotasFiscais.Nota_Id	                                                                                         " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "     INNER JOIN (SELECT Empresa_Id,                                                                                                                             " & vbCrLf & _
            "						EndEmpresa_Id,                                                                                                                           " & vbCrLf & _
            "						Cliente_Id,                                                                                                                              " & vbCrLf & _
            "						EndCliente_Id,                                                                                                                           " & vbCrLf & _
            "						EntradaSaida_Id,                                                                                                                         " & vbCrLf & _
            "						Serie_Id,                                                                                                                                " & vbCrLf & _
            "						Nota_Id,                                                                                                                                 " & vbCrLf & _
            "						SUM(QuantidadeFiscal) AS QuantidadeFiscal,                                                                                               " & vbCrLf & _
            "						SUM(Valor) AS Valor                                                                                                                      " & vbCrLf & _
            "					FROM NotasFiscaisXItens                                                                                                                      " & vbCrLf & _
            "					GROUP BY Empresa_id,                                                                                                                         " & vbCrLf & _
            "					EndEmpresa_id,                                                                                                                               " & vbCrLf & _
            "					Cliente_id,                                                                                                                                  " & vbCrLf & _
            "					EndCliente_id,                                                                                                                               " & vbCrLf & _
            "					EntradaSaida_id,                                                                                                                             " & vbCrLf & _
            "					Serie_id,                                                                                                                                    " & vbCrLf & _
            "					Nota_id                                                                                                                                      " & vbCrLf & _
            "				) As NFI                                                                                                                                         " & vbCrLf & _
            "		on NFI.Empresa_Id      = NotasFiscais.Empresa_Id                                                                                                         " & vbCrLf & _
            "		AND NFI.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id                                                                                                     " & vbCrLf & _
            "		AND NFI.Cliente_Id      = NotasFiscais.Cliente_Id                                                                                                        " & vbCrLf & _
            "		AND NFI.EndCliente_Id   = NotasFiscais.EndCliente_Id                                                                                                     " & vbCrLf & _
            "		AND NFI.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id                                                                                                   " & vbCrLf & _
            "		AND NFI.Serie_Id        = NotasFiscais.Serie_Id                                                                                                          " & vbCrLf & _
            "		AND NFI.Nota_Id         = NotasFiscais.Nota_Id                                                                                                           " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "    INNER JOIN Clientes                                                                                                                                         " & vbCrLf & _
            "		ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id                                                                                                      " & vbCrLf & _
            "		AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id                                                                                                    " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "	INNER JOIN Clientes cli2                                                                                                                                     " & vbCrLf & _
            "		ON NotaFiscalReferencial.ClienteReferencial_Id    = cli2.Cliente_Id                                                                                      " & vbCrLf & _
            "		AND NotaFiscalReferencial.EndClienteReferencial_Id = cli2.Endereco_Id                                                                                    " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "	LEFT JOIN NFERealizadas                                                                                                                                      " & vbCrLf & _
            "		ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id                                                                                               " & vbCrLf & _
            "		AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id                                                                                           " & vbCrLf & _
            "		AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id                                                                                              " & vbCrLf & _
            "		AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id                                                                                           " & vbCrLf & _
            "		AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id                                                                                         " & vbCrLf & _
            "		AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id                                                                                                " & vbCrLf & _
            "		AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id                                                                                                 " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "	LEFT JOIN NFERealizadas nfer                                                                                                                                 " & vbCrLf & _
            "		ON nf2.Empresa_Id      = nfer.Empresa_Id                                                                                                                 " & vbCrLf & _
            "		AND nf2.EndEmpresa_Id   = nfer.EndEmpresa_Id                                                                                                             " & vbCrLf & _
            "		AND nf2.Cliente_Id      = nfer.Cliente_Id                                                                                                                " & vbCrLf & _
            "		AND nf2.EndCliente_Id   = nfer.EndCliente_Id                                                                                                             " & vbCrLf & _
            "		AND nf2.EntradaSaida_Id = nfer.EntradaSaida_Id                                                                                                           " & vbCrLf & _
            "		AND nf2.Serie_Id        = nfer.Serie_Id                                                                                                                  " & vbCrLf & _
            "		AND nf2.Nota_Id         = nfer.Nota_Id		                                                                                                             " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "	LEFT JOIN (SELECT	nxi.Empresa_Id,                                                                                                                          " & vbCrLf & _
            "						nxi.EndEmpresa_Id,                                                                                                                       " & vbCrLf & _
            "						nxi.Cliente_Id,                                                                                                                          " & vbCrLf & _
            "						nxi.EndCliente_Id,                                                                                                                       " & vbCrLf & _
            "						nxi.EntradaSaida_Id,                                                                                                                     " & vbCrLf & _
            "						nxi.Serie_Id,                                                                                                                            " & vbCrLf & _
            "						nxi.Nota_Id,                                                                                                                             " & vbCrLf & _
            "						SUM(nfr.Quantidade) AS QuantidadeFiscalConsumido,                                                                                        " & vbCrLf & _
            "						SUM(nfr.Valor) AS ValorConsumido                                                                                                         " & vbCrLf & _
            "				FROM NotasFiscaisXItens nxi                                                                                                                      " & vbCrLf & _
            "					INNER JOIN NotaFiscalReferencial nfr                                                                                                         " & vbCrLf & _
            "						ON nxi.Empresa_Id = nfr.Empresa_Id                                                                                                       " & vbCrLf & _
            "						AND nxi.EndEmpresa_Id = nfr.EndEmpresa_Id                                                                                                " & vbCrLf & _
            "						AND nxi.Cliente_Id = nfr.Cliente_Id                                                                                                      " & vbCrLf & _
            "						AND nxi.EndCliente_Id = nfr.EndCliente_Id                                                                                                " & vbCrLf & _
            "						AND nxi.EntradaSaida_Id = nfr.EntradaSaida_Id                                                                                            " & vbCrLf & _
            "						AND nxi.Nota_Id = nfr.Nota_Id                                                                                                            " & vbCrLf & _
            "						AND nxi.Serie_Id = nfr.Serie_Id                                                                                                          " & vbCrLf & _
            "						AND nxi.Produto_Id = nfr.Produto_Id                                                                                                      " & vbCrLf & _
            "						AND nxi.CFOP_Id = nfr.CFOP_Id                                                                                                            " & vbCrLf & _
            "						AND nxi.Sequencia_Id = nfr.Sequencia_Id                                                                                                  " & vbCrLf & _
            "				GROUP BY nxi.Empresa_id,                                                                                                                         " & vbCrLf & _
            "				nxi.EndEmpresa_id,                                                                                                                               " & vbCrLf & _
            "				nxi.Cliente_id,                                                                                                                                  " & vbCrLf & _
            "				nxi.EndCliente_id,                                                                                                                               " & vbCrLf & _
            "				nxi.EntradaSaida_id,                                                                                                                             " & vbCrLf & _
            "				nxi.Serie_id,                                                                                                                                    " & vbCrLf & _
            "				nxi.Nota_id                                                                                                                                      " & vbCrLf & _
            "				) As Consumido                                                                                                                                   " & vbCrLf & _
            "		on Consumido.Empresa_Id      = NotasFiscais.Empresa_Id                                                                                                   " & vbCrLf & _
            "		AND Consumido.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id                                                                                               " & vbCrLf & _
            "		AND Consumido.Cliente_Id      = NotasFiscais.Cliente_Id                                                                                                  " & vbCrLf & _
            "		AND Consumido.EndCliente_Id   = NotasFiscais.EndCliente_Id                                                                                               " & vbCrLf & _
            "		AND Consumido.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id                                                                                             " & vbCrLf & _
            "		AND Consumido.Serie_Id        = NotasFiscais.Serie_Id                                                                                                    " & vbCrLf & _
            "		AND Consumido.Nota_Id         = NotasFiscais.Nota_Id                                                                                                     " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
            "    LEFT JOIN NotaFiscalDevolucaoXNotaFiscal Dev                                                                                                                " & vbCrLf & _
            "		ON Dev.EmpresaDevolucao_Id	 = NotasFiscais.Empresa_Id                                                                                                   " & vbCrLf & _
            "		AND Dev.EndEmpresaDevolucao_Id = NotasFiscais.EndEmpresa_Id                                                                                              " & vbCrLf & _
            "		AND Dev.ClienteDevolucao_Id	 = NotasFiscais.Cliente_Id                                                                                                   " & vbCrLf & _
            "		AND Dev.EndClienteDevolucao_Id = NotasFiscais.EndCliente_Id                                                                                              " & vbCrLf & _
            "		AND Dev.Nota_Id			     = NotasFiscais.Nota_Id                                                                                                      " & vbCrLf & _
            "		AND Dev.Serie_Id		         = NotasFiscais.Serie_Id                                                                                                 " & vbCrLf & _
            "		AND Dev.EntradaSaida_Id        = NotasFiscais.EntradaSaida_Id                                                                                            " & vbCrLf & _
            "                                                                                                                                                                " & vbCrLf & _
                  "                                                                                                                                                                " & vbCrLf & _
        "  WHERE 1=1                                                                                                              " & vbCrLf & _
        "       AND NotasFiscais.Empresa_Id = '" & MyParameters("empresa") & "'" & vbCrLf & _
        "       AND NotasFiscais.EndEmpresa_Id = " & MyParameters("enderecoEmpresa") & vbCrLf & _
        "       AND NotasFiscais.EntradaSaida_Id = '" & MyParameters("es") & "'" & vbCrLf & _
        "       AND NotasFiscais.Situacao = 1 " & vbCrLf & _
        "       AND NotasFiscais.CIFFOB = '" & MyParameters("ciffob") & "'" & vbCrLf

        If Not String.IsNullOrWhiteSpace(MyParameters("codcliente")) Then
            sql &= "  AND NotasFiscais.Cliente_Id = '" & MyParameters("codcliente").Split("-")(0) & "'" & vbCrLf
            sql &= "  AND NotasFiscais.EndCliente_Id = " & MyParameters("codcliente").Split("-")(1) & "" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(MyParameters("numconhecimento")) Then
            Dim strNota() As String = MyParameters("numconhecimento").Trim().Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
            If strNota IsNot Nothing AndAlso strNota.Length > 1 Then
                sql &= "  AND NotasFiscais.Nota_Id IN (" & String.Join(",", strNota) & ") " & vbCrLf
            Else
                sql &= "  AND NotasFiscais.Nota_Id IN (" & strNota(0) & ") " & vbCrLf
            End If
        End If

        sql &= "  AND NotasFiscais.Movimento between '" & MyParameters("data1").tosqldate() & "' and '" & MyParameters("data2").tosqldate() & "'" & vbCrLf & _
            "   ORDER BY NotasFiscais.Nota_Id, NotasFiscais.Serie_Id, NFERealizadas.Data DESC " & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "CtrcFerroviario")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim cnpj As String = Funcoes.FormatarCpfCnpj(row("Ctrc_Cliente").ToString().Split("-")(0))
            Dim nome As String = row("Ctrc_Cliente").ToString().Substring(row("Ctrc_Cliente").ToString().IndexOf("-"), row("Ctrc_Cliente").ToString().Length - row("Ctrc_Cliente").ToString().IndexOf("-"))
            row("Ctrc_Cliente") = String.Format("{0} {1}", cnpj, nome)

            cnpj = Funcoes.FormatarCpfCnpj(row("Cliente").ToString().Split("-")(0))
            nome = row("Cliente").ToString().Substring(row("Cliente").ToString().IndexOf("-"), row("Cliente").ToString().Length - row("Cliente").ToString().IndexOf("-"))
            row("Cliente") = String.Format("{0} {1}", cnpj, nome)
        Next

        Return ds
    End Function

    Private Function getDataSetRodoviario() As DataSet
        Dim ds As New DataSet
        Dim sql As String = ""

        sql &= "SELECT t.* FROM (                                                                                           " & vbCrLf & _
        "       SELECT DISTINCT                                                                                             " & vbCrLf & _
        "           nf.Pedido,                                                                                              " & vbCrLf & _
        "           nf.Cliente_Id + '-' + cast(nf.EndCliente_Id as varchar) + '-' + c.Nome as Cliente_Id,                   " & vbCrLf & _
        "           nf.Movimento,                                                                                           " & vbCrLf & _
        "           nxi.Produto_Id,                                                                                         " & vbCrLf & _
        "           p.Nome as DescProduto,                                                                                  " & vbCrLf & _
        "           cast(nf.Nota_Id as varchar) as Nota_Id,                                                                 " & vbCrLf & _
        "           nf.Serie_Id,                                                                                            " & vbCrLf & _
        "           nxt.Placa,                                                                                              " & vbCrLf & _
        "           isnull(rxp.Pesagem_Id,0) as NumeroTicket,                                                               " & vbCrLf & _
        "           isnull((CASE WHEN nxi.QuantidadeFisica <= 0 THEN nxi.QuantidadeFiscal ELSE nxi.QuantidadeFisica END),0) as PesoBalanca, " & vbCrLf & _
        "           isnull(nxd.PesoLiquido,0) as PesoChegada,                                                               " & vbCrLf & _
        "           isnull(CASE WHEN isnull(nxd.PesoLiquido,0) > 0                                                          " & vbCrLf & _
        "              THEN (isnull(nxd.PesoLiquido,0) - isnull(nxi.QuantidadeFisica,0))                                    " & vbCrLf & _
        "              ELSE 0                                                                                               " & vbCrLf & _
        "           END, 0) as Quebra,                                                                                      " & vbCrLf & _
        "           isnull(nxi.QuantidadeFiscal,0) as PesoNotaFiscal,                                                       " & vbCrLf & _
        "           isnull(cast((nxi.Valor / nxi.QuantidadeFiscal) as decimal(18,2)),0) as Unitario,                        " & vbCrLf & _
        "           isnull(nxi.Valor,0) as ValorNotaFiscal,                                                                 " & vbCrLf & _
        "           cast(cte.Nota_Id as varchar) + '-' + cast(cte.Serie_Id as varchar) as Ctrc_Id,                          " & vbCrLf & _
        "           isnull(ctexi.Unitario,0) as UnitarioFrete,                                                              " & vbCrLf & _
        "           (((isnull(ctexi.Valor,0) + isnull(cteXenc.Valor,0)) / (CASE WHEN nxi.QuantidadeFisica <= 0 THEN nxi.QuantidadeFiscal ELSE nxi.QuantidadeFisica END)) * 1000) as TarifaFrete,  " & vbCrLf & _
        "           isnull(cteXenc.Valor,0) as TarifaPedagio,                                                               " & vbCrLf & _
        "           isnull(ctexi.Valor,0) as ValorFrete,                                                                    " & vbCrLf & _
        "           emb.Nome as LocalEmbarque                                                                               " & vbCrLf & _
        "       FROM NotasFiscais nf                                            " & vbCrLf & _
        "                                                                       " & vbCrLf & _
        "INNER JOIN NotasFiscaisXItens nxi                                      " & vbCrLf & _
        "   ON nf.Empresa_Id = nxi.Empresa_Id                                   " & vbCrLf & _
        "   and nf.EndEmpresa_Id = nxi.EndEmpresa_Id                            " & vbCrLf & _
        "   and nf.Cliente_Id = nxi.Cliente_Id                                  " & vbCrLf & _
        "   and nf.EndCliente_Id = nxi.EndCliente_Id                            " & vbCrLf & _
        "   and nf.EntradaSaida_Id = nxi.EntradaSaida_Id                        " & vbCrLf & _
        "   and nf.Serie_Id = nxi.Serie_Id                                      " & vbCrLf & _
        "   and nf.Nota_Id = nxi.Nota_Id                                        " & vbCrLf & _
        "                                                                       " & vbCrLf & _
        "INNER JOIN Produtos p	                                                " & vbCrLf & _
        "   ON p.Produto_Id = nxi.Produto_Id                                    " & vbCrLf & _
        "                                                                       " & vbCrLf & _
        "INNER JOIN Clientes c" & vbCrLf & _
        "   ON c.Cliente_Id = NF.Cliente_Id    " & vbCrLf & _
        "   AND C.Endereco_Id = NF.EndCliente_Id" & vbCrLf & _
        "                                                                       " & vbCrLf & _
        "LEFT JOIN NotasFiscaisXTransportadores nxt                             " & vbCrLf & _
        "   ON nf.Empresa_Id = nxt.Empresa_Id                                   " & vbCrLf & _
        "   and nf.EndEmpresa_Id = nxt.EndEmpresa_Id                            " & vbCrLf & _
        "   and nf.Cliente_Id = nxt.Cliente_Id                                  " & vbCrLf & _
        "   and nf.EndCliente_Id = nxt.EndCliente_Id                            " & vbCrLf & _
        "   and nf.EntradaSaida_Id = nxt.EntradaSaida_Id                        " & vbCrLf & _
        "   and nf.Serie_Id = nxi.Serie_Id                                      " & vbCrLf & _
        "   and nf.Nota_Id = nxt.Nota_Id                                        " & vbCrLf & _
        "                                                                       " & vbCrLf & _
        "LEFT JOIN NotasFiscaisXRomaneios nfr                                   " & vbCrLf & _
        "   ON nxi.Empresa_Id = nfr.Empresa_Id                                  " & vbCrLf & _
        "   and nxi.EndEmpresa_Id = nfr.EndEmpresa_Id                           " & vbCrLf & _
        "   and nxi.Cliente_Id = nfr.Cliente_Id                                 " & vbCrLf & _
        "   and nxi.EndCliente_Id = nfr.EndCliente_Id                           " & vbCrLf & _
        "   and nxi.EntradaSaida_Id = nfr.EntradaSaida_Id                       " & vbCrLf & _
        "   and nxi.Serie_Id = nfr.Serie_Id                                     " & vbCrLf & _
        "   and nxi.Nota_Id = nfr.Nota_Id                                       " & vbCrLf & _
        "                                                                       " & vbCrLf & _
        "LEFT JOIN RomaneiosXPesagens rxp                                       " & vbCrLf & _
        "   ON nfr.Empresa_Id = rxp.Empresa_Id                                  " & vbCrLf & _
        "   and nfr.EndEmpresa_Id = rxp.EndEmpresa_Id                           " & vbCrLf & _
        "   and nfr.Romaneio_Id = rxp.Romaneio_Id                               " & vbCrLf & _
        "                                                                        " & vbCrLf & _
        "LEFT JOIN NotasXDestinos nxd                                            " & vbCrLf & _
        "   ON nxd.Empresa_Id = nxi.Empresa_Id                                   " & vbCrLf & _
        "   and nxd.EndEmpresa_Id = nxi.EndEmpresa_Id                            " & vbCrLf & _
        "   and nxd.Cliente_Id = nxi.Cliente_Id                                  " & vbCrLf & _
        "   and nxd.EndCliente_Id = nxi.EndCliente_Id                            " & vbCrLf & _
        "   and nxd.EntradaSaida_Id = nxi.EntradaSaida_Id                        " & vbCrLf & _
        "   and nxd.Serie_Id = nxi.Serie_Id                                      " & vbCrLf & _
        "   and nxd.Nota_Id = nxi.Nota_Id                                        " & vbCrLf & _
        "                                                                        " & vbCrLf & _
        "LEFT JOIN Clientes emb                                                  " & vbCrLf & _
        "   ON emb.Cliente_Id = isnull(nf.LocalEmbarque, nf.Deposito)            " & vbCrLf & _
        "   and emb.Endereco_Id = isnull(nf.EndLocalEmbarque, nf.EndDeposito)    " & vbCrLf & _
        "                                                                        " & vbCrLf & _
        "LEFT JOIN (                                                             " & vbCrLf & _
        "   select nxn.* from NotasXNotas nxn                                    " & vbCrLf & _
        "   LEFT JOIN NotaFiscalReferencial ref                                  " & vbCrLf & _
        "   ON ref.EmpresaReferencial_Id = nxn.Empresa_Id                        " & vbCrLf & _
        "   and ref.EndEmpresaReferencial_Id = nxn.EndEmpresa_Id                 " & vbCrLf & _
        "   and ref.ClienteReferencial_Id = nxn.Cliente_Id                       " & vbCrLf & _
        "   and ref.EndClienteReferencial_Id = nxn.EndCliente_Id                 " & vbCrLf & _
        "   and ref.EntradaSaidaReferencial_Id = nxn.EntradaSaida_Id             " & vbCrLf & _
        "   and ref.SerieReferencial_Id = nxn.Serie_Id                           " & vbCrLf & _
        "   and ref.NotaReferencial_Id = nxn.Nota_Id                             " & vbCrLf & _
        "   WHERE 1=1                                                            " & vbCrLf & _
        "   and ref.NotaReferencial_Id is null                                   " & vbCrLf & _
        ") as nxn                                                                " & vbCrLf & _
        "ON nxn.OrigemEmpresa_Id = nf.Empresa_Id                                 " & vbCrLf & _
        "   and nxn.OrigemEndEmpresa_Id = nf.EndEmpresa_Id                       " & vbCrLf & _
        "   and nxn.OrigemCliente_Id = nf.Cliente_Id                             " & vbCrLf & _
        "   and nxn.OrigemEndCliente_Id = nf.EndCliente_Id                       " & vbCrLf & _
        "   and nxn.OrigemEntradaSaida_Id = nf.EntradaSaida_Id                   " & vbCrLf & _
        "   and nxn.OrigemSerie_Id = nf.Serie_Id                                 " & vbCrLf & _
        "   and nxn.OrigemNota_Id = nf.Nota_Id                                   " & vbCrLf & _
        "                                                                       " & vbCrLf & _
        "LEFT JOIN NotasFiscais cte                                             " & vbCrLf & _
        "   ON nxn.Empresa_Id = cte.Empresa_Id                                  " & vbCrLf & _
        "   and nxn.EndEmpresa_Id = cte.EndEmpresa_Id                           " & vbCrLf & _
        "   and nxn.Cliente_Id = cte.Cliente_Id                                 " & vbCrLf & _
        "   and nxn.EndCliente_Id = cte.EndCliente_Id                           " & vbCrLf & _
        "   and nxn.EntradaSaida_Id = cte.EntradaSaida_Id                       " & vbCrLf & _
        "   and nxn.Serie_Id = cte.Serie_Id                                     " & vbCrLf & _
        "   and nxn.Nota_Id = cte.Nota_Id                                       " & vbCrLf & _
        "                                                                       " & vbCrLf & _
        "LEFT JOIN NotasFiscaisXItens ctexi                                     " & vbCrLf & _
        "   ON cte.Empresa_Id = ctexi.Empresa_Id                                " & vbCrLf & _
        "   and cte.EndEmpresa_Id = ctexi.EndEmpresa_Id                         " & vbCrLf & _
        "   and cte.Cliente_Id = ctexi.Cliente_Id                               " & vbCrLf & _
        "   and cte.EndCliente_Id = ctexi.EndCliente_Id                         " & vbCrLf & _
        "   and cte.EntradaSaida_Id = ctexi.EntradaSaida_Id                     " & vbCrLf & _
        "   and cte.Serie_Id = ctexi.Serie_Id                                   " & vbCrLf & _
        "   and cte.Nota_Id = ctexi.Nota_Id                                     " & vbCrLf & _
        "                                                                       " & vbCrLf & _
        "LEFT JOIN NotasFiscaisXEncargos cteXenc                                " & vbCrLf & _
        "   ON cteXenc.Empresa_Id = ctexi.Empresa_Id                            " & vbCrLf & _
        "   and cteXenc.EndEmpresa_Id = ctexi.EndEmpresa_Id                     " & vbCrLf & _
        "   and cteXenc.Cliente_Id = ctexi.Cliente_Id                           " & vbCrLf & _
        "   and cteXenc.EndCliente_Id = ctexi.EndCliente_Id                     " & vbCrLf & _
        "   and cteXenc.EntradaSaida_Id = ctexi.EntradaSaida_Id                 " & vbCrLf & _
        "   and cteXenc.Serie_Id = ctexi.Serie_Id                               " & vbCrLf & _
        "   and cteXenc.Nota_Id = ctexi.Nota_Id                                 " & vbCrLf & _
        "   and cteXenc.Produto_Id = ctexi.Produto_Id                           " & vbCrLf & _
        "   and cteXenc.CFOP_Id = ctexi.CFOP_Id                                 " & vbCrLf & _
        "   and cteXenc.Encargo_Id = 'TARIFA PEDAGIO'                           " & vbCrLf & _
        "                                                                    " & vbCrLf & _
        "WHERE      nf.Empresa_Id = '" & MyParameters("empresa") & "'" & vbCrLf & _
        "       AND nf.EndEmpresa_Id = " & MyParameters("enderecoEmpresa") & vbCrLf & _
        "       AND nf.EntradaSaida_Id = '" & MyParameters("es") & "'" & vbCrLf & _
        "       AND nxi.QuantidadeFiscal > 0 " & vbCrLf

        If Not String.IsNullOrWhiteSpace(MyParameters("numconhecimento")) Then
            sql &= "   AND nf.Nota_Id = '" & MyParameters("numconhecimento") & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(MyParameters("codcliente")) Then
            sql &= "   AND nf.Cliente_Id = '" & MyParameters("codcliente").Split("-")(0) & "'" & vbCrLf & _
            "   AND nf.EndCliente_Id = " & MyParameters("codcliente").Split("-")(1) & "" & vbCrLf
        End If

        sql &= "   AND nf.Situacao in (1,4,7)                                                  " & vbCrLf & _
        "   AND nf.TipoDeDocumento = 1                                                  " & vbCrLf & _
        "   AND nf.Movimento between '" & MyParameters("data1").tosqldate() & "' and '" & MyParameters("data2").tosqldate() & "'" & vbCrLf & _
        " ) as t                                                                                " & vbCrLf & _
        " ORDER BY t.Nota_Id, t.Movimento Desc                                        " & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "NotasDoPedido")
        '#NFD 01

        For Each row As DataRow In ds.Tables(0).Rows
            Dim cnpj As String = Funcoes.FormatarCpfCnpj(row("Cliente_Id").ToString().Split("-")(0))
            Dim nome As String = row("Cliente_Id").ToString().Substring(row("Cliente_Id").ToString().IndexOf("-"), row("Cliente_Id").ToString().Length - row("Cliente_Id").ToString().IndexOf("-"))
            row("Cliente_Id") = String.Format("{0} {1}", cnpj, nome)
        Next

        Return ds
    End Function

    Private Function getDataSetRodoviarioXFerroviario() As DataSet
        Dim ds As New DataSet
        Dim sql As String = "   SELECT t.* FROM (                                 " & vbCrLf & _
            "   	select f.* from vw_ferroviario f              " & vbCrLf & _
            "   	UNION ALL                                     " & vbCrLf & _
            "   	select r.* from vw_rodoviario r               " & vbCrLf & _
            "   ) as t                                            " & vbCrLf & _
            "   WHERE 1=1                                         " & vbCrLf & _
            "   AND t.Empresa_Id = '" & MyParameters("empresa") & "'" & vbCrLf & _
            "   AND t.EndEmpresa_Id = " & MyParameters("enderecoEmpresa") & vbCrLf & _
            "   AND t.EntradaSaida = '" & MyParameters("es") & "'" & vbCrLf

        If Not String.IsNullOrWhiteSpace(MyParameters("codcliente")) Then
            sql &= "  AND t.Cliente_Id = '" & MyParameters("codcliente").Split("-")(0) & "'" & vbCrLf
            sql &= "  AND t.EndCliente_Id = " & MyParameters("codcliente").Split("-")(1) & "" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(MyParameters("numconhecimento")) Then
            Dim strNota() As String = MyParameters("numconhecimento").Trim().Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
            If strNota IsNot Nothing AndAlso strNota.Length > 1 Then
                sql &= "  AND t.Nota_Id IN (" & String.Join(",", strNota) & ") " & vbCrLf
            Else
                sql &= "  AND t.Nota_Id IN (" & strNota(0) & ") " & vbCrLf
            End If
        End If

        If Not String.IsNullOrWhiteSpace(MyParameters("numconhecimento")) Then
            sql &= "   AND t.Nota_Id in (" & Replace(MyParameters("numconhecimento"), ";", ",") & ")                           " & vbCrLf
        End If

        sql &= "   AND t.Data between '" & MyParameters("data1").tosqldate() & "' and '" & MyParameters("data2").tosqldate() & "'  " & vbCrLf & _
            "   ORDER BY t.Nota_Id, t.Serie_Id, t.Data DESC       " & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "RodoviarioXFerroviario")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim cnpj As String = ""
            Dim nome As String = ""
            If Not IsDBNull(row("Ctrc_Cliente")) AndAlso Not String.IsNullOrWhiteSpace(row("Ctrc_Cliente")) Then
                cnpj = Funcoes.FormatarCpfCnpj(row("Ctrc_Cliente").ToString().Split("-")(0))
                nome = row("Ctrc_Cliente").ToString().Substring(row("Ctrc_Cliente").ToString().IndexOf("-"), row("Ctrc_Cliente").ToString().Length - row("Ctrc_Cliente").ToString().IndexOf("-"))
                row("Ctrc_Cliente") = String.Format("{0} {1}", cnpj, nome)
            End If
            If Not IsDBNull(row("Cliente")) AndAlso Not String.IsNullOrWhiteSpace(row("Cliente")) Then
                cnpj = Funcoes.FormatarCpfCnpj(row("Cliente").ToString().Split("-")(0))
                nome = row("Cliente").ToString().Substring(row("Cliente").ToString().IndexOf("-"), row("Cliente").ToString().Length - row("Cliente").ToString().IndexOf("-"))
                row("Ctrc_Cliente") = String.Format("{0} {1}", cnpj, nome)
            End If
        Next

        Return ds
    End Function

    Private Sub habilitarCampos()
        rdoPorCtrc.Enabled = rdoFerroviario.Checked
        rdoPorNota.Enabled = rdoFerroviario.Checked
    End Sub

#End Region

End Class