Imports Microsoft.VisualBasic
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports System.Collections.Generic
Imports System.Web
Imports System.Web.UI
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class MemorandoDeExportacaoEspelho
    Implements IBaseEntity

    Public Sub ExibirEspelho(ByVal page As System.Web.UI.Page, ByVal EmpresaMemorando As String, ByVal EndEmpresaMemorando As Integer, ByVal Memorando_Id As Integer)
        Dim MemEx As New MemorandoDeExportacao

        MemEx.CodigoEmpresaMemorando = EmpresaMemorando
        MemEx.EnderecoEmpresaMemorando = EndEmpresaMemorando
        MemEx.CodigoMemorando = Memorando_Id

        ExibirEspelho(page, MemEx)
    End Sub

    Public Sub ExibirEspelho(ByVal page As System.Web.UI.Page, ByVal MemExp As MemorandoDeExportacao)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String

        sql = "Select MemorandoDeExportacao.EmpresaMemorando_Id, MemorandoDeExportacao.EndEmpresaMemorando_Id, EmpresaMemorando.Nome As NomeEmpresaMemorando," & vbCrLf & _
              "       MemorandoDeExportacao.Memorando_Id, " & vbCrLf & _
              "       ClienteMemorando, EndClienteMemorando, ClienteMemorando.Nome AS NomeClienteMemorando," & vbCrLf & _
              "       NossaEmissao, DataMemorando, NrDespachoExp, DataDespachoExp, DataAverbacao, TipoDocumento, " & vbCrLf & _
              "       PaisDestino, Pais.Descricao AS NomePaisDestino, Navio, Representante.Nome AS Representante, Representante.Cliente_Id AS Representante_Id," & vbCrLf & _
              "       Empresa, EndEmpresa,ISNULL(Empresa.Nome,'') AS NomeEmpresa,  " & vbCrLf & _
              "       Cliente, EndCliente, ISNULL(Cliente.Nome,'') AS NomeCliente, " & vbCrLf & _
              "       Nota, Serie, EntradaSaida," & vbCrLf & _
              "       Produto, P.Nome As NomeProduto," & vbCrLf & _
              "       ExportadorEquiparado, EndExportadorEquiparado, MemorandoEquiparado, Isnull(ExportadorEquiparado.Nome,'') As NomeExportadorEquiparado, " & vbCrLf & _
              "       ExportadorEquiparado.Cidade AS CidadeExportadorEquiparado," & vbCrLf & _
              "       ExportadorEquiparado.Estado    AS EstadoExportadorEquiparado," & vbCrLf & _
              "       ExportadorEquiparado.Complemento AS ComplementoExportadorEquiparado," & vbCrLf & _
              "       ExportadorEquiparado.Inscricao   AS InscricaoExportadorEquiparado," & vbCrLf & _
              "       ExportadorEquiparado.Endereco + ' N. ' + convert(nvarchar,ExportadorEquiparado.Numero) + ' CEP. ' + ExportadorEquiparado.Cep AS EnderecoExportadorEquiparado " & vbCrLf & _
              "  From MemorandoDeExportacao " & vbCrLf & _
              " INNER JOIN Clientes ClienteMemorando" & vbCrLf & _
              "    on ClienteMemorando.cliente_id  = MemorandoDeExportacao.ClienteMemorando " & vbCrLf & _
              "   and ClienteMemorando.Endereco_Id = MemorandoDeExportacao.EndClienteMemorando " & vbCrLf & _
              " INNER JOIN Clientes EmpresaMemorando" & vbCrLf & _
              "    on EmpresaMemorando.cliente_id  = MemorandoDeExportacao.EmpresaMemorando_Id " & vbCrLf & _
              "   and EmpresaMemorando.Endereco_Id = MemorandoDeExportacao.EndEmpresaMemorando_Id " & vbCrLf & _
              "  LEFT JOIN Clientes Empresa" & vbCrLf & _
              "    on Empresa.cliente_id  = MemorandoDeExportacao.Empresa " & vbCrLf & _
              "   and Empresa.Endereco_Id = MemorandoDeExportacao.EndEmpresa " & vbCrLf & _
              "  LEFT JOIN Clientes Cliente" & vbCrLf & _
              "    ON Cliente.cliente_id  = MemorandoDeExportacao.Cliente " & vbCrLf & _
              "   and Cliente.Endereco_Id = MemorandoDeExportacao.EndCliente " & vbCrLf & _
              "  LEFT JOIN Clientes  ExportadorEquiparado" & vbCrLf & _
              "    ON ExportadorEquiparado.cliente_id  = MemorandoDeExportacao.ExportadorEquiparado " & vbCrLf & _
              "   and ExportadorEquiparado.Endereco_Id = MemorandoDeExportacao.EndExportadorEquiparado " & vbCrLf & _
              "  LEFT JOIN Clientes  Representante" & vbCrLf & _
              "    ON Representante.cliente_id  = MemorandoDeExportacao.Representante " & vbCrLf & _
              "   and Representante.Endereco_Id = MemorandoDeExportacao.endRepresentante " & vbCrLf & _
              " INNER JOIN Produtos P" & vbCrLf & _
              "    ON P.Produto_id = MemorandoDeExportacao.Produto " & vbCrLf & _
              "  LEFT JOIN Pais " & vbCrLf & _
              "    ON Pais.Pais_id = MemorandoDeExportacao.PaisDestino " & vbCrLf & _
              " where MemorandoDeExportacao.EmpresaMemorando_Id    ='" & MemExp.CodigoEmpresaMemorando & "'" & vbCrLf & _
              "   and MemorandoDeExportacao.EndEmpresaMemorando_Id = " & MemExp.EnderecoEmpresaMemorando & vbCrLf & _
              "   and MemorandoDeExportacao.Memorando_Id           ='" & MemExp.CodigoMemorando & "'" & vbCrLf
        ds = Banco.ConsultaDataSet(sql, "MemorandoDeExportacao")

        sql = "SELECT NF.UsuarioInclusao," & vbCrLf & _
              "       NF.UsuarioInclusaoData," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then NF.Empresa_Id       else NF.Cliente_Id       end Emitente, " & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then NF.EndEmpresa_Id    else NF.EndCliente_Id    end EndEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Empresa.Nome        else Cliente.Nome        end NomeEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Empresa.Cidade      else Cliente.Cidade      end CidadeEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Empresa.Estado      else Cliente.Estado      end EstadoEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Empresa.Complemento else Cliente.Complemento end ComplementoEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Empresa.Inscricao   else Cliente.Inscricao   end InscricaoEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Empresa.Endereco + ' N. ' + convert(nvarchar,Empresa.Numero) + ' CEP. ' + Empresa.Cep   else Cliente.Endereco + ' N. ' + convert(nvarchar,Cliente.Numero) + ' CEP. ' + Cliente.Cep end EnderecoEmitente, " & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then NF.Cliente_Id       else NF.Empresa_Id       end Destinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then NF.EndCliente_Id    else NF.EndEmpresa_Id    end EndDestinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Cliente.Nome        else Empresa.Nome        end NomeDestinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Cliente.Cidade      else Empresa.Cidade      end CidadeDestinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Cliente.Estado      else Empresa.Estado      end EstadoDestinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Cliente.Complemento else Empresa.Complemento end ComplementoDestinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Cliente.Inscricao   else Empresa.Inscricao   end InscricaoDestinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'E' then Cliente.Endereco + ' N. ' + convert(nvarchar,Cliente.Numero) + ' CEP. ' + Cliente.Cep else Empresa.Endereco + ' N. ' + convert(nvarchar,Empresa.Numero) + ' CEP. ' + Empresa.Cep    end EnderecoDestinatario, " & vbCrLf & _
              "       NF.EntradaSaida_Id," & vbCrLf & _
              "       NF.Serie_Id," & vbCrLf & _
              "       NF.Nota_Id," & vbCrLf & _
              "       NF.Pedido," & vbCrLf & _
              "       NF.Operacao," & vbCrLf & _
              "       NF.SubOperacao," & vbCrLf & _
              "       SO.Descricao as DescSubOperacao," & vbCrLf & _
              "       NF.Movimento," & vbCrLf & _
              "       NF.DataDaNota," & vbCrLf & _
              "       NFxI.Produto_Id," & vbCrLf & _
              "       P.Unidade," & vbCrLf & _
              "       P.Nome as NomeProduto," & vbCrLf & _
              "       P.NCM as NCMProduto," & vbCrLf & _
              "       NFxI.CFOP_Id," & vbCrLf & _
              "       NFxI.Sequencia_Id," & vbCrLf & _
              "       MeXNF.Quantidade AS QuantidadeFiscal," & vbCrLf & _
              "       NFxI.Unitario," & vbCrLf & _
              "       NFxI.Valor as ValorItem " & vbCrLf & _
              "  FROM NotasFiscais NF" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
              "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
              "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
              " INNER JOIN Clientes Empresa" & vbCrLf & _
              "    on Empresa.cliente_id  = NF.Empresa_Id" & vbCrLf & _
              "   and Empresa.Endereco_Id = NF.EndEmpresa_Id" & vbCrLf & _
              " INNER JOIN Clientes Cliente" & vbCrLf & _
              "    ON Cliente.cliente_id  = NF.Cliente_id" & vbCrLf & _
              "   and Cliente.Endereco_Id = NF.EndCliente_Id" & vbCrLf & _
              " INNER JOIN SubOperacoes SO" & vbCrLf & _
              "    on SO.Operacao_id     = NF.Operacao" & vbCrLf & _
              "   and SO.SubOperacoes_id = NF.SubOperacao" & vbCrLf & _
              " INNER JOIN Produtos P" & vbCrLf & _
              "    ON P.Produto_id = NFxI.Produto_Id" & vbCrLf & _
              " INNER JOIN  MemorandoDeExportacaoXNotaFiscal AS MeXNF " & vbCrLf & _
              "    ON NF.Empresa_Id      = MeXNF.Empresa_Id  " & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = MeXNF.EndEmpresa_Id  " & vbCrLf & _
              "   AND NF.Cliente_Id      = MeXNF.Cliente_Id  " & vbCrLf & _
              "   AND NF.EndCliente_Id   = MeXNF.EndCliente_Id  " & vbCrLf & _
              "   AND NF.EntradaSaida_Id = MeXNF.EntradaSaida_Id  " & vbCrLf & _
              "   AND NF.Serie_Id        = MeXNF.Serie_Id  " & vbCrLf & _
              "   AND NF.Nota_Id         = MeXNF.Nota_Id " & vbCrLf & _
              " where MeXNF.EmpresaMemorando_Id    ='" & MemExp.CodigoEmpresaMemorando & "'" & vbCrLf & _
              "   and MeXNF.EndEmpresaMemorando_Id = " & MemExp.EnderecoEmpresaMemorando & vbCrLf & _
              "   and MeXNF.Memorando_Id           ='" & MemExp.CodigoMemorando & "'" & vbCrLf & _
              " Order by NF.Nota_id " & vbCrLf
        ds.Merge(Banco.ConsultaDataSet(sql, "MemorandoDeExportacaoXNotaFiscal"))

        sql = "SELECT MemxRegExp.EmpresaMemorando_Id, MemxRegExp.EndEmpresaMemorando_Id,  " & vbCrLf & _
              "       MemxRegExp.Memorando_Id, MemxRegExp.RegistroDeExportacao_Id,  " & vbCrLf & _
              "       MemxRegExp.DataExportacao " & vbCrLf & _
              "  FROM MemorandoDeExportacaoXRegistroDeExportacao AS MemxRegExp  " & vbCrLf & _
              " INNER JOIN MemorandoDeExportacao  " & vbCrLf & _
              "    ON MemxRegExp.EmpresaMemorando_Id    = MemorandoDeExportacao.EmpresaMemorando_Id  " & vbCrLf & _
              "   AND MemxRegExp.EndEmpresaMemorando_Id = MemorandoDeExportacao.EndEmpresaMemorando_Id  " & vbCrLf & _
              "   AND MemxRegExp.Memorando_Id           = MemorandoDeExportacao.Memorando_Id " & vbCrLf & _
              " INNER JOIN Clientes AS Empresa  " & vbCrLf & _
              "    ON Empresa.Cliente_Id  = MemxRegExp.EmpresaMemorando_Id  " & vbCrLf & _
              "   AND Empresa.Endereco_Id = MemxRegExp.EndEmpresaMemorando_Id  " & vbCrLf & _
              " where MemorandoDeExportacao.EmpresaMemorando_Id    ='" & MemExp.CodigoEmpresaMemorando & "'" & vbCrLf & _
              "   and MemorandoDeExportacao.EndEmpresaMemorando_Id = " & MemExp.EnderecoEmpresaMemorando & vbCrLf & _
              "   and MemorandoDeExportacao.Memorando_Id           ='" & MemExp.CodigoMemorando & "'" & vbCrLf & _
              " Order by MemxRegExp.RegistroDeExportacao_Id " & vbCrLf
        ds.Merge(Banco.ConsultaDataSet(sql, "MemorandoDeExportacaoXRegistroDeExportacao"))

        sql = "SELECT MemxConhec.EmpresaMemorando_Id, MemxConhec.EndEmpresaMemorando_Id,   " & vbCrLf & _
              "       MemxConhec.Memorando_Id, " & vbCrLf & _
              "       MemxConhec.ConhecimentoDeEmbarque_Id, MemxConhec.DataConhecimento, TipoConhecimento,  " & vbCrLf & _
              "       Case   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '01' then 'AWB'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '02' then 'MAWB'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '03' then 'HAWB'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '04' then 'COMAT'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '06' then 'R. EXPRESSAS'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '07' then 'ETIQ. REXPRESSAS'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '08' then 'HR. EXPRESSAS'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '09' then 'AV7'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '10' then 'BL'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '11' then 'MBL'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '12' then 'HBL'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '13' then 'CRT'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '14' then 'DSIC'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '16' then 'COMAT BL'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '17' then 'RWB'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '18' then 'HRWB'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '19' then 'TIF/DTA'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '20' then 'CP2'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '91' then 'NÂO IATA'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '92' then 'MNAO IATA'   " & vbCrLf & _
              "         when MemxConhec.TipoConhecimento = '93' then 'HNAO IATA'   " & vbCrLf & _
              "         When MemxConhec.TipoConhecimento = '99' then 'OUTROS'   " & vbCrLf & _
              "       end DescTipoConhecimento  " & vbCrLf & _
              "  FROM MemorandoDeExportacaoXConhecimentoDeEmbarque AS MemxConhec   " & vbCrLf & _
              " INNER JOIN MemorandoDeExportacao   " & vbCrLf & _
              "    ON MemxConhec.EmpresaMemorando_Id    = MemorandoDeExportacao.EmpresaMemorando_Id   " & vbCrLf & _
              "   AND MemxConhec.EndEmpresaMemorando_Id = MemorandoDeExportacao.EndEmpresaMemorando_Id   " & vbCrLf & _
              "   AND MemxConhec.Memorando_Id           = MemorandoDeExportacao.Memorando_Id  " & vbCrLf & _
              " where MemorandoDeExportacao.EmpresaMemorando_Id    ='" & MemExp.CodigoEmpresaMemorando & "'" & vbCrLf & _
              "   and MemorandoDeExportacao.EndEmpresaMemorando_Id = " & MemExp.EnderecoEmpresaMemorando & vbCrLf & _
              "   and MemorandoDeExportacao.Memorando_Id           ='" & MemExp.CodigoMemorando & "'" & vbCrLf & _
              " Order by MemxConhec.ConhecimentoDeEmbarque_Id " & vbCrLf
        ds.Merge(Banco.ConsultaDataSet(sql, "MemorandoDeExportacaoXConhecimentoDeEmbarque"))

        'Dados da NF de Exportação
        sql = " SELECT " & vbCrLf & _
                " NF.Nota_Id, " & vbCrLf & _
                " NF.Serie_Id," & vbCrLf & _
                " SUM(NFxI.Valor) AS Valor," & vbCrLf & _
                " SUM(NFxI.QuantidadeFisica) AS Quantidade " & vbCrLf & _
                " FROM NotasFiscais NF " & vbCrLf & _
                " INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf & _
                "	ON NF.Empresa_Id		= NFxI.Empresa_Id " & vbCrLf & _
                "	AND NF.EndEmpresa_Id	= NFxI.EndEmpresa_Id " & vbCrLf & _
                "	AND NF.Cliente_Id		= NFxI.Cliente_Id  " & vbCrLf & _
                "	AND NF.EndCliente_Id	= NFxI.EndCliente_Id " & vbCrLf & _
                "	AND NF.EntradaSaida_Id	= NFxI.EntradaSaida_Id " & vbCrLf & _
                "	AND NF.Serie_Id			= NFxI.Serie_Id " & vbCrLf & _
                "	AND NF.Nota_Id			= NFxI.Nota_Id  " & vbCrLf & _
                "where NF.Empresa_Id		= '" & MemExp.CodigoEmpresaMemorando & "'" & vbCrLf & _
                "	and NF.EndEmpresa_Id	= " & MemExp.EnderecoEmpresaMemorando & vbCrLf & _
                "	and NF.Nota_Id			= " & MemExp.NumeroNota & vbCrLf & _
                "	and NF.Serie_Id			= " & MemExp.Serie & vbCrLf & _
                "	and NF.Cliente_Id		= '" & MemExp.CodigoCliente & "'" & vbCrLf & _
                "	and NF.EndCliente_Id	= " & MemExp.EnderecoCliente & vbCrLf & _
                "	AND NF.EntradaSaida_Id	= '" & MemExp.EntradaSaida & "'" & vbCrLf & _
                "        Group BY " & vbCrLf & _
                "        NF.Nota_Id, " & vbCrLf & _
                "        NF.Serie_Id " & vbCrLf

        ds.Merge(Banco.ConsultaDataSet(sql, "MemorandoNFExportacao"))

        'Imagem
        Dim dtImagem As DataTable = ds.Tables.Add("Images")
        dtImagem.Columns.Add("path", GetType(String))
        dtImagem.Columns.Add("image", GetType(System.Byte()))

        Dim drImagem As DataRow = dtImagem.NewRow()
        Dim strCaminhoImagem As String = HttpContext.Current.Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))

        drImagem("path") = strCaminhoImagem
        drImagem("image") = Conversoes.ConverterImagemEmByte(strCaminhoImagem)
        dtImagem.Rows.Add(drImagem)

        Dim crpt As New ReportDocument()

        crpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/Cr_MemorandoDeExportacao1.rpt")
        crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

        Dim NomeArquivo2 As String = "Files/" & MemExp.CodigoEmpresaMemorando & "-" & MemExp.CodigoClienteMemorando & "-" & Funcoes.EliminarCaracteresEspeciais(MemExp.CodigoMemorando) & "-Ctrl" & (New Random).Next & ".PDF"
        Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo2)
        Dim arquivo As String = NomeArquivo

        Try
            crpt.SetDataSource(ds)
            Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()
            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)
            crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
            Funcoes.AbrirArquivo(page, NomeArquivo2)
        Catch ex As Exception
            MsgBox(page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try
    End Sub

End Class
