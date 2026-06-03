Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Reclassificacao
    Inherits BasePage

    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "", False)
                If Funcoes.VerificaPermissao("Reclassificacao", "ACESSAR") Then
                    Limpar()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página. As permissões são: ACESSAR, LEITURA e ALTERAR.", "Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        txtLaudo.Enabled = True
        lnkConsultar.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False

        cmdAjustar.Visible = False

        HID.Value = Guid.NewGuid().ToString

        txtLaudo.Text = String.Empty
        txtRomaneio.Text = String.Empty
        txtPrimeiraPesagem.Text = String.Empty
        txtSegundaPesagem.Text = String.Empty
        txtPesoBruto.Text = String.Empty
        txtDesconto.Text = String.Empty
        txtLiquido.Text = String.Empty

        gridDescontos.DataSource = Nothing
        gridDescontos.DataBind()
        gridDescontos.Parent.Visible = False
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub ConsultarLaudo()
        Dim Emp As String() = ddlEmpresa.SelectedValue.Split("-")
        objLaudo = New Pesagem(Emp(0), Emp(1), txtLaudo.Text)

        If objLaudo.Codigo = 0 Then
            MsgBox(Me.Page, "Laudo de Pesagem não encontrado")
            txtLaudo.Focus()
            Exit Sub
        ElseIf objLaudo.CodigoSituacao = 2 Then
            MsgBox(Me.Page, "Laudo de Pesagem Cancelado não pode ser alterado")
            txtLaudo.Focus()
            Exit Sub
        ElseIf objLaudo.Analises Is Nothing OrElse objLaudo.Analises.Count = 0 Then
            MsgBox(Me.Page, "Laudo de Pesagem não possui análises")
            txtLaudo.Focus()
            Exit Sub
        End If

        SessaoSalvaLaudo()

        txtRomaneio.Text = objLaudo.CodigoRomaneio

        txtPrimeiraPesagem.Text = objLaudo.PrimeiraPesagem
        txtSegundaPesagem.Text = objLaudo.SegundaPesagem
        txtPesoBruto.Text = objLaudo.BrutoBalanca
        txtDesconto.Text = objLaudo.Desconto
        txtLiquido.Text = objLaudo.Liquido

        '------Analises------------------------------------
        gridDescontos.DataSource = objLaudo.Analises
        gridDescontos.DataBind()
        gridDescontos.Parent.Visible = True

        For Each row As GridViewRow In gridDescontos.Rows
            CType(row.FindControl("txtPercentual"), TextBox).Enabled = Not (objLaudo.TemNota OrElse objLaudo.Romaneios.Count > 1)
            CType(row.FindControl("txtIndice"), TextBox).Enabled = False
            CType(row.FindControl("txtDesconto"), TextBox).Enabled = False

            If row.Cells(0).Text = 12 AndAlso objLaudo.TemNota Then
                If Funcoes.VerificaPermissao("AjustarRomaneio", "ACESSAR") Then
                    CType(row.FindControl("ddlOpcao"), DropDownList).Enabled = True
                    cmdAjustar.Visible = True
                End If
            End If
        Next

        txtLaudo.Enabled = False
        cmdCalcular.Enabled = True
        lnkConsultar.Parent.Visible = False

        Dim temMosanto As Boolean = False

        If objLaudo.TemNota Or objLaudo.TemRateio Then
            cmdCalcular.Enabled = False

            If objLaudo.TemRateio Then
                MsgBox(Me.Page, "Laudo de pesagem com rateio não pode ser alterado.")
            Else
                MsgBox(Me.Page, "Laudo de pesagem vinculado com nota fiscal não pode ser alterado. Apenas ajustar Intacta/Monsanto.")
            End If
        End If
    End Sub


    Private Sub AlterarLaudo()
        Dim SqlArray As New ArrayList
        Dim i As Integer
        Dim Empresa() As String = ddlEmpresa.SelectedValue.Split("-")

        Sql = "UPDATE Pesagem SET " & vbCrLf & _
              "   PrimeiraPesagem      = " & CInt(txtPrimeiraPesagem.Text) & " " & vbCrLf & _
              "  ,SegundaPesagem       = " & CInt(txtSegundaPesagem.Text) & " " & vbCrLf & _
              "  ,BrutoBalanca         = " & CInt(txtPesoBruto.Text) & " " & vbCrLf & _
              "  ,Liquido              = " & CInt(txtLiquido.Text) & " " & vbCrLf & _
              "  ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "' " & vbCrLf & _
              "  ,UsuarioAlteracaoData ='" & Now().ToString("yyyy-MM-dd hh:mm:ss") & "' " & vbCrLf & _
              " WHERE Empresa_id    ='" & Empresa(0) & "'" & vbCrLf & _
              "   AND EndEmpresa_id = " & Empresa(1) & vbCrLf & _
              "   AND Pesagem_Id    = " & txtLaudo.Text & vbCrLf & _
              "   AND Sequencia_Id  = 0;" & vbCrLf
        SqlArray.Add(Sql)

        Sql = "Delete from PesagemXAnalises " & vbCrLf & _
              "WHERE Empresa_id    ='" & Empresa(0) & "'" & vbCrLf & _
              "  AND EndEmpresa_id = " & Empresa(1) & vbCrLf & _
              "  AND Pesagem_Id    = " & txtLaudo.Text & vbCrLf & _
              "  AND Sequencia_Id  = 0;" & vbCrLf
        SqlArray.Add(Sql)

        For i = 0 To gridDescontos.Rows.Count - 1
            Sql = "Insert Into PesagemXAnalises (Empresa_Id, EndEmpresa_Id, Pesagem_Id, Sequencia_Id, Analise_Id, Percentual, Indice, Desconto)  " & vbCrLf & _
                  "Values('" & Empresa(0) & "', " & Empresa(1) & ", " & txtLaudo.Text & ", 0, " & vbCrLf & _
                  "" & gridDescontos.Rows(i).Cells(0).Text() & ", " & vbCrLf & _
                  "" & CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text.Replace(",", ".") & ", " & vbCrLf & _
                  "" & CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text.Replace(",", ".") & ", " & vbCrLf & _
                  "" & CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text.Replace(",", ".") & ");" & vbCrLf
            SqlArray.Add(Sql)
        Next i

        Sql = "UPDATE Romaneios SET " & vbCrLf & _
              "   PesoBruto   = " & CInt(txtPesoBruto.Text) & " " & vbCrLf & _
              "  ,Desconto    = " & (CInt(txtPesoBruto.Text) - CInt(txtLiquido.Text)) & " " & vbCrLf & _
              "  ,PesoLiquido = " & CInt(txtLiquido.Text) & " " & vbCrLf & _
              " WHERE Empresa_id    ='" & Empresa(0) & "'" & vbCrLf & _
              "   AND EndEmpresa_id = " & Empresa(1) & vbCrLf & _
              "   AND Romaneio_Id   = " & txtRomaneio.Text & ";" & vbCrLf
        SqlArray.Add(Sql)

        Sql = "Delete RomaneiosXDescontos " & vbCrLf & _
              " WHERE Empresa_id    ='" & Empresa(0) & "'" & vbCrLf & _
              "   AND EndEmpresa_id = " & Empresa(1) & vbCrLf & _
              "   AND Romaneio_Id   = " & txtRomaneio.Text & ";" & vbCrLf
        SqlArray.Add(Sql)

        For i = 0 To gridDescontos.Rows.Count - 1
            Sql = "Insert Into RomaneiosXDescontos(Empresa_Id, EndEmpresa_Id,Romaneio_ID, Analise_Id,Desconto, Percentual, Indice)" & vbCrLf & _
                  "Values('" & Empresa(0) & "'," & Empresa(1) & ", " & txtRomaneio.Text & ", " & gridDescontos.Rows(i).Cells(0).Text() & ", " & vbCrLf & _
                  "" & CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text.Replace(",", ".") & ", " & vbCrLf & _
                  "" & CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text.Replace(",", ".") & ", " & vbCrLf & _
                  "" & CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text.Replace(",", ".") & ");" & vbCrLf
            SqlArray.Add(Sql)
        Next i

        If Banco.GravaBanco(SqlArray) = False Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        Else
            MsgBox(Me.Page, "Laudo ajustado com sucesso.")

            If Imprimir(txtLaudo.Text) Then
                Limpar()
                txtLaudo.Focus()
            End If
        End If
    End Sub

    Public Function Imprimir(ByVal Laudo As String) As Boolean
        Dim ds As New DataSet
        Dim retorno As Boolean = True

        SessaoRecuperaLaudo()

        Sql = "SELECT Pesagem.Pesagem_Id AS Laudo, Pesagem.Produto, Pesagem.Placa, Pesagem.EntradaSaida, Pesagem.PrimeiraPesagem, " & vbCrLf &
              "		  Pesagem.SegundaPesagem, Pesagem.BrutoBalanca, Pesagem.BrutoBalanca - Pesagem.Liquido AS Descontos, Pesagem.Liquido, Pesagem.EntradaPatio, " & vbCrLf &
              "		  Pesagem.EntradaBalanca, Pesagem.SaidaBalanca, Pesagem.Movimento, Pesagem.NumeroDaNota AS NotaFiscal, Pesagem.SerieDaNota AS SerieNota, " & vbCrLf &
              "		  Pesagem.PesoFiscal, Pesagem.Observacoes, Produtos.Nome AS NomeProduto, Clientes.Cliente_Id AS CodigoCliente, Clientes.Endereco_Id AS EndCliente, " & vbCrLf &
              "		  Clientes.Nome AS NomeCliente, Clientes.Reduzido AS ReduzidoCliente, Clientes.Endereco AS EnderecoCliente, Clientes.Cidade AS CidadeCliente, " & vbCrLf &
              "		  Clientes.Estado AS EstadoCliente, isnull(Transportes.Cliente_Id,'') AS CodigoTransportador, isnull(Transportes.Endereco_Id,0) AS EndTransportador, " & vbCrLf &
              "       isnull(Transportes.Nome,'') AS NomeTransportador, isnull(Transportes.Reduzido,'') AS ReduzidoTransportador, isnull(Transportes.Endereco,'') AS EnderecoTransportador, " & vbCrLf &
              "       isnull(Transportes.Cidade,'') AS CidadeTransportador, isnull(Transportes.Estado,'') AS EstadoTransportador, " & vbCrLf &
              "		  Depositos.Cliente_Id AS CodigoDeposito, Depositos.Endereco_Id AS EndDeposito, Depositos.Nome AS NomeDeposito, Depositos.Reduzido AS ReduzidoDeposito, " & vbCrLf &
              "       Depositos.Endereco AS EnderecoDeposito, Depositos.Cidade AS CidadeDeposito, Depositos.Estado AS EstadoDeposito, Depositos.Inscricao AS InscricaoDeposito, " & vbCrLf &
              "		  isnull(Placas.Placa01,'') AS Placa01, isnull(Placas.Placa02,'') AS Placa02, isnull(Placas.Placa03,'') AS Placa03, isnull(Placas.CidadePlaca,'') AS CidadePlaca, isnull(Placas.EstadoPlaca,'') AS EstadoPlaca, " & vbCrLf &
              "		  ISNULL(Motorista.Nome,'') AS NomeMotorista, ISNULL(Motorista.Cidade,'') AS CidadeMotorista, ISNULL(Motorista.Estado,'') AS EstadoMotorista, " & vbCrLf &
              "		  ISNULL(Motorista.Habilitacao,'') AS Habilitacao, " & vbCrLf &
              "		  CASE " & vbCrLf &
              "		  	WHEN LEN(ISNULL(Motorista.Cliente_Id,'')) = 0 " & vbCrLf &
              "		  	  THEN '' " & vbCrLf &
              "		  	  ELSE CASE " & vbCrLf &
              "		  			 WHEN LEN(Motorista.Cliente_Id) < 11 " & vbCrLf &
              "		  			   THEN '' " & vbCrLf &
              "		  			   ELSE SUBSTRING(Motorista.Cliente_Id, 1, 3) + '.' + SUBSTRING(Motorista.Cliente_Id, 4, 3) + '.' + SUBSTRING(Motorista.Cliente_Id, 7, 3) + '-' + SUBSTRING(Motorista.Cliente_Id, 10, 2) " & vbCrLf &
              "		  			END " & vbCrLf &
              "		  END AS CpfMotorista, " & vbCrLf &
              "		  isnull(EstadoPlacaMotorista.Descricao,'') AS NomeEstadoMotorista, " & vbCrLf &
              "		  isnull(EstadoPlaca.Descricao,'') AS NomeEstadoPlaca, " & vbCrLf &
              "		  Pedido, Depositos.Numero AS NumeroDeposito, " & vbCrLf &
              "		  Depositos.Complemento AS ComplementoDeposito, Depositos.Bairro AS BairroDeposito, Clientes.Numero AS NumeroCliente, Clientes.Complemento AS ComplementoCliente, " & vbCrLf &
              "		  Clientes.Bairro AS BairroCliente, Clientes.Inscricao AS InscricaoCliente, Empresa.Cliente_Id  AS CodigoEmpresa, Empresa.Endereco_Id AS EndEmpresa, " & vbCrLf &
              "		  Empresa.Nome AS NomeEmpresa, Empresa.Reduzido AS ReduzidoEmpresa, Empresa.Endereco AS EnderecoEmpresa, Empresa.Cidade  AS CidadeEmpresa, " & vbCrLf &
              "		  Empresa.Estado AS EstadoEmpresa, Empresa.Inscricao AS InscricaoEmpresa, Empresa.Numero AS NumeroEmpresa, Empresa.Complemento AS ComplementoEmpresa, " & vbCrLf &
              "		  Empresa.Bairro AS BairroEmpresa, isnull(rXp.Romaneio_id,0) AS Romaneio " & vbCrLf &
              "  FROM Pesagem " & vbCrLf &
              "	INNER JOIN Produtos " & vbCrLf &
              "	   ON Pesagem.Produto = Produtos.Produto_Id " & vbCrLf &
              "	INNER JOIN Clientes " & vbCrLf &
              "	   ON Pesagem.Cliente = Clientes.Cliente_Id " & vbCrLf &
              "	  AND Pesagem.EndCliente = Clientes.Endereco_Id " & vbCrLf &
              "	 LEFT JOIN Clientes AS Transportes " & vbCrLf &
              "	   ON Pesagem.Transportador = Transportes.Cliente_Id " & vbCrLf &
              "	  AND Pesagem.EndTransportador = Transportes.Endereco_Id " & vbCrLf &
              "  LEFT JOIN Clientes AS Motorista " & vbCrLf &
              "	   ON Pesagem.Motorista = Motorista.Cliente_Id " & vbCrLf &
              "	  AND Pesagem.EndMotorista = Motorista.Endereco_Id " & vbCrLf &
              "  LEFT OUTER JOIN Estados AS EstadoPlacaMotorista " & vbCrLf &
              "    ON Motorista.Estado = EstadoPlacaMotorista.Estado_Id " & vbCrLf &
              "	INNER JOIN Clientes AS Depositos " & vbCrLf &
              "	   ON Pesagem.Deposito = Depositos.Cliente_Id " & vbCrLf &
              "	  AND Pesagem.EndDeposito = Depositos.Endereco_Id " & vbCrLf &
              "	 LEFT JOIN Placas " & vbCrLf &
              "	   ON Pesagem.Placa = Placas.Placa_Id " & vbCrLf &
              "	 LEFT JOIN Estados AS EstadoPlaca " & vbCrLf &
              "	   ON Placas.EstadoPlaca = EstadoPlaca.Estado_Id " & vbCrLf &
              "	Inner Join Clientes as Empresa " & vbCrLf &
              "	   on Empresa.cliente_id = Pesagem.Empresa_id " & vbCrLf &
              "	  and Empresa.endereco_id = Pesagem.EndEmpresa_id " & vbCrLf &
              "	 LEFT JOIN RomaneiosXPesagens rXp " & vbCrLf &
              "	   on rXp.Empresa_id     = Pesagem.Empresa_id " & vbCrLf &
              "	  and rXp.EndEmpresa_Id = Pesagem.EndEmpresa_id " & vbCrLf &
              "	  and rXp.Pesagem_id    = Pesagem.Pesagem_id " & vbCrLf &
              "	  and rXp.Sequencia_Id  = Pesagem.Sequencia_Id " & vbCrLf &
              " WHERE (Pesagem.Empresa_Id = '" & objLaudo.CodigoEmpresa & "') " & vbCrLf &
              "   AND (Pesagem.EndEmpresa_Id = " & objLaudo.EnderecoEmpresa & ") " & vbCrLf &
              "   AND (Pesagem.Pesagem_Id = " & objLaudo.Codigo & ") " & vbCrLf &
              "   AND (Pesagem.Sequencia_Id = 0)"

        ds = Banco.ConsultaDataSet(Sql, "Laudo")

        Sql = "Select CLA.Analise_Id AS Analise, " & vbCrLf &
              "       case " & vbCrLf &
              "         when Len(isnull(Analises.Opcao,'')) > 0 " & vbCrLf &
              "           then Analises.Descricao + ' - ' + (select valor from fnStringToArray(Analises.Opcao,';') where left(valor,len(convert(int, pxa.percentual))+1) = convert(varchar,convert(int,pxa.percentual)) + '-')" & vbCrLf &
              "           else Analises.Descricao " & vbCrLf &
              "       end AS Descricao, " & vbCrLf &
              "       isnull(pxa.percentual,0) as Percentual," & vbCrLf &
              "       isnull(pxa.Indice,CLA.Indice) as indice," & vbCrLf &
              "       isnull(pxa.Desconto,0) as Desconto " & vbCrLf &
              "  from classificacoes CLA " & vbCrLf &
              "  left join Analises " & vbCrLf &
              "    on CLA.Analise_Id = Analises.Analise_Id " & vbCrLf &
              "  left join Pesagem p" & vbCrLf &
              "	   on p.Empresa_Id      ='" & objLaudo.CodigoEmpresa & "'" & vbCrLf &
              "	  and p.EndEmpresa_Id   = " & objLaudo.EnderecoEmpresa & vbCrLf &
              "	  and p.Pesagem_Id      = " & objLaudo.Codigo & vbCrLf &
              "   and p.sequencia_id    = " & objLaudo.Sequencia & vbCrLf &
              "  Left Join PesagemXAnalises pxa" & vbCrLf &
              "    ON pxa.Empresa_Id    = p.Empresa_Id " & vbCrLf &
              "   AND pxa.EndEmpresa_Id = p.EndEmpresa_Id " & vbCrLf &
              "   AND pxa.Pesagem_Id    = p.Pesagem_Id " & vbCrLf &
              "   AND pxa.Sequencia_Id  = p.Sequencia_Id " & vbCrLf &
              "   And pxa.Analise_id    = cla.Analise_Id" & vbCrLf &
              "  Where CLA.Tabela_id    = " & objLaudo.CodigoTabelaDeClassificacao & vbCrLf &
              "    and CLA.Produto_id   ='" & objLaudo.CodigoProduto & "'" & vbCrLf &
              "    and CLA.Sequencia_Id = 1 " & vbCrLf &
              "  ORDER BY CLA.Analise_Id "

        ds.Merge(Banco.ConsultaDataSet(Sql, "Analises"))

        Dim crpt As New ReportDocument()
        crpt.FileName = Server.MapPath("~/Reports/Cr_LaudoDePesagem.rpt")
        crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

        Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
        Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
        Dim arquivo As String = NomeArquivo

        Try
            crpt.SetDataSource(ds)

            Dim crparametervalues As ParameterValues
            Dim crparameterdiscretevalue As ParameterDiscreteValue
            Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

            crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

            crparameterfielddefinition = crparameterfielddefinitions.Item("Reemissao")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = "REEMISSÃO"
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
            Funcoes.AbrirArquivo(Me.Page, NomeArquivo2)
        Catch ex As Exception
            retorno = False
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try

        Return retorno
    End Function

    Protected Sub CalcularDesconto()
        SessaoRecuperaLaudo()

        For i = 0 To gridDescontos.Rows.Count - 1
            Dim codanalise As Integer = gridDescontos.Rows(i).Cells(0).Text
            Dim Analis As PesagemXAnalises = objLaudo.Analises.Where(Function(s) s.CodigoAnalise = codanalise).First

            If Analis.Analise.Opcao.Length = 0 Then
                Dim Percentual As String = CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text
                Analis.Percentual = IIf(IsNumeric(Percentual), Percentual, 0)
            Else
                Dim ddlop As DropDownList = CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList)
                Analis.Percentual = ddlop.SelectedValue
            End If
        Next

        Dim Erro As String = objLaudo.Analises.CalcularDescontos()
        txtDesconto.Text = "0"

        If (String.IsNullOrEmpty(Erro)) Then
            gridDescontos.DataSource = objLaudo.Analises
            gridDescontos.DataBind()

            Dim pDesconto As Decimal = 0
            For Each aN In objLaudo.Analises
                If Not aN.CodigoAnalise = 6 AndAlso Not aN.CodigoAnalise = 12 Then
                    pDesconto += aN.Desconto
                End If
            Next

            'Teste da soma dos descontos 27/02/2016
            Dim desc As Decimal = objLaudo.Analises.Sum(Function(s) s.Desconto)
            If desc <> pDesconto Then
                MsgBox(Me.Page, "Erro #12 avise a TI")
            End If

            txtLiquido.Text = objLaudo.Liquido

            For i = 0 To gridDescontos.Rows.Count - 1
                CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Enabled = False
                CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Enabled = False
                CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Enabled = False
            Next

            lnkAtualizar.Parent.Visible = True
        Else
            MsgBox(Me.Page, Erro)
        End If
    End Sub

#Region "Session"
    Private objLaudo As [Lib].Negocio.Pesagem

    Private Sub SessaoSalvaLaudo()
        Session("ssLaudo" & HID.Value) = objLaudo
    End Sub

    Private Sub SessaoRecuperaLaudo()
        objLaudo = CType(Session("ssLaudo" & HID.Value), [Lib].Negocio.Pesagem)
    End Sub
#End Region

#Region "Botoes"
    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("Reclassificacao", "LEITURA") Then
                If txtLaudo.Text.Length = 0 Then
                    MsgBox(Me.Page, "Número da pesagem não foi informado")
                Else
                    ConsultarLaudo()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consulta registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Reclassificacao", "ALTERAR") Then
                AlterarLaudo()
            Else
                MsgBox(Me.Page, "Usuario sem permissão para alterar registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Reclassificacao")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdCalcular_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) AndAlso Not txtSegundaPesagem.Text.Equals("0") Then
            CalcularDesconto()
        Else
            For i = 0 To gridDescontos.Rows.Count - 1
                CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text = 0
                CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text = 0
                CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text = 0
            Next
            txtPesoBruto.Text = txtPrimeiraPesagem.Text
            txtLiquido.Text = txtPrimeiraPesagem.Text
            MsgBox(Me.Page, "Segunda pesagem deve ser maior que zero!")
        End If
    End Sub

    Protected Sub cmdAjustar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim SqlArray As New ArrayList
            Dim i As Integer
            Dim Empresa() As String = ddlEmpresa.SelectedValue.Split("-")

            Sql = "UPDATE Pesagem SET " & vbCrLf &
              "   UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "' " & vbCrLf &
              "  ,UsuarioAlteracaoData ='" & Now().ToString("yyyy-MM-dd hh:mm:ss") & "' " & vbCrLf &
              " WHERE Empresa_id    ='" & Empresa(0) & "'" & vbCrLf &
              "   AND EndEmpresa_id = " & Empresa(1) & vbCrLf &
              "   AND Pesagem_Id    = " & txtLaudo.Text & vbCrLf &
              "   AND Sequencia_Id  = 0;" & vbCrLf
            SqlArray.Add(Sql)

            Sql = "Delete from PesagemXAnalises " & vbCrLf &
              "WHERE Empresa_id    ='" & Empresa(0) & "'" & vbCrLf &
              "  AND EndEmpresa_id = " & Empresa(1) & vbCrLf &
              "  AND Pesagem_Id    = " & txtLaudo.Text & vbCrLf &
              "  AND Sequencia_Id  = 0;" & vbCrLf
            SqlArray.Add(Sql)

            For i = 0 To gridDescontos.Rows.Count - 1

                If gridDescontos.Rows(i).Cells(0).Text = 12 Then
                    Sql = "Insert Into PesagemXAnalises (Empresa_Id, EndEmpresa_Id, Pesagem_Id, Sequencia_Id, Analise_Id, Percentual, Indice, Desconto)  " & vbCrLf &
                          "Values('" & Empresa(0) & "', " & Empresa(1) & ", " & txtLaudo.Text & ", 0, " & vbCrLf &
                          "" & gridDescontos.Rows(i).Cells(0).Text() & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList).SelectedValue & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text.Replace(",", ".") & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text.Replace(",", ".") & ");" & vbCrLf
                Else
                    Sql = "Insert Into PesagemXAnalises (Empresa_Id, EndEmpresa_Id, Pesagem_Id, Sequencia_Id, Analise_Id, Percentual, Indice, Desconto)  " & vbCrLf &
                          "Values('" & Empresa(0) & "', " & Empresa(1) & ", " & txtLaudo.Text & ", 0, " & vbCrLf &
                          "" & gridDescontos.Rows(i).Cells(0).Text() & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text.Replace(",", ".") & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text.Replace(",", ".") & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text.Replace(",", ".") & ");" & vbCrLf
                End If

                SqlArray.Add(Sql)
            Next i

            Sql = "Delete RomaneiosXDescontos " & vbCrLf &
              " WHERE Empresa_id    ='" & Empresa(0) & "'" & vbCrLf &
              "   AND EndEmpresa_id = " & Empresa(1) & vbCrLf &
              "   AND Romaneio_Id   = " & txtRomaneio.Text & ";" & vbCrLf
            SqlArray.Add(Sql)

            For i = 0 To gridDescontos.Rows.Count - 1

                If gridDescontos.Rows(i).Cells(0).Text = 12 Then
                    Sql = "Insert Into RomaneiosXDescontos(Empresa_Id, EndEmpresa_Id,Romaneio_ID, Analise_Id,Desconto, Percentual, Indice)" & vbCrLf &
                          "Values('" & Empresa(0) & "'," & Empresa(1) & ", " & txtRomaneio.Text & ", " & gridDescontos.Rows(i).Cells(0).Text() & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text.Replace(",", ".") & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList).SelectedValue & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text.Replace(",", ".") & ");" & vbCrLf
                Else
                    Sql = "Insert Into RomaneiosXDescontos(Empresa_Id, EndEmpresa_Id,Romaneio_ID, Analise_Id,Desconto, Percentual, Indice)" & vbCrLf &
                          "Values('" & Empresa(0) & "'," & Empresa(1) & ", " & txtRomaneio.Text & ", " & gridDescontos.Rows(i).Cells(0).Text() & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text.Replace(",", ".") & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text.Replace(",", ".") & ", " & vbCrLf &
                          "" & CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text.Replace(",", ".") & ");" & vbCrLf
                End If

                SqlArray.Add(Sql)
            Next i

            If Banco.GravaBanco(SqlArray) Then

                MsgBox(Me.Page, "Laudo ajustado com sucesso.")

                If Imprimir(txtLaudo.Text) Then
                    Limpar()
                    txtLaudo.Focus()
                End If
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

#End Region

    Protected Sub gridDescontos_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridDescontos.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim PxA As PesagemXAnalises = CType(e.Row.DataItem, PesagemXAnalises)
                If PxA.Analise.Opcao.Length > 0 Then
                    CType(e.Row.FindControl("txtPercentual"), TextBox).Visible = False
                    CType(e.Row.FindControl("txtIndice"), TextBox).Visible = False
                    CType(e.Row.FindControl("txtDesconto"), TextBox).Visible = False

                    RemovedControl(CType(e.Row.FindControl("txtPercentual"), TextBox))
                    RemovedControl(CType(e.Row.FindControl("txtIndice"), TextBox))
                    RemovedControl(CType(e.Row.FindControl("txtDesconto"), TextBox))

                    Dim ddlA As DropDownList = CType(e.Row.FindControl("ddlOpcao"), DropDownList)
                    ddlA.Visible = True
                    ddl.Carregar(ddlA, PxA.Analise.Opcao, ";", "-")

                    ddlA.Enabled = False
                    ddlA.SelectedValue = CInt(PxA.Percentual)
                    e.Row.Cells(2).ColumnSpan = 3
                End If
        End Select
    End Sub
End Class