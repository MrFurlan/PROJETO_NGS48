Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListAutorizacaoDeRetirada
    Inherits List(Of AutorizacaoDeRetirada)

#Region "Variaveis Auxiliares"
    Private Sql As String
    Private Banco As New AcessaBanco
    Private ds As DataSet
#End Region

#Region "Fields"
    Private _QtdeAutorizadaFisica As Double = 0
    Private _QtdeEntregueFisica As Double = 0
    Private _QtdeAutorizadaFiscal As Double = 0
    Private _QtdeEntregueFiscal As Double = 0
    Private _QtdeSaldoFisico As Double = 0
    Private _QtdeSaldoFiscal As Double = 0
#End Region

#Region "Property"
    Public Property QtdeAutorizadaFisica() As Double
        Get
            Return _QtdeAutorizadaFisica
        End Get
        Set(ByVal value As Double)
            _QtdeAutorizadaFisica = value
        End Set
    End Property

    Public Property QtdeEntregueFisica() As Double
        Get
            Return _QtdeEntregueFisica
        End Get
        Set(ByVal value As Double)
            _QtdeEntregueFisica = value
        End Set
    End Property

    Public Property QtdeAutorizadaFiscal() As Double
        Get
            Return _QtdeAutorizadaFiscal
        End Get
        Set(ByVal value As Double)
            _QtdeAutorizadaFiscal = value
        End Set
    End Property

    Public Property QtdeEntregueFiscal() As Double
        Get
            Return _QtdeEntregueFiscal
        End Get
        Set(ByVal value As Double)
            _QtdeEntregueFiscal = value
        End Set
    End Property

    Public Property QtdeSaldoFisico() As Double
        Get
            Return _QtdeSaldoFisico
        End Get
        Set(ByVal value As Double)
            _QtdeSaldoFisico = value
        End Set
    End Property

    Public Property QtdeSaldoFiscal() As Double
        Get
            Return _QtdeSaldoFiscal
        End Get
        Set(ByVal value As Double)
            _QtdeSaldoFiscal = value
        End Set
    End Property
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pcliente As String, ByVal pEndCliente As Integer, ByVal pPedido As Integer, ByVal Classe As [Lib].Negocio.eClassesOperacoes, Optional ByVal pRomaneio As Boolean = False, Optional ByVal pAno As String = "")
        If pRomaneio = True Then
            Sql = "	SELECT Au.Empresa_Id, Au.EndEmpresa_id, Au.Pedido_id, Au.Autorizacao_Id, Au.Movimento, Au.DataBaseCalculo, Au.DataVencimento, Au.ClienteRetirada, Au.EndClienteRetirada, " & vbCrLf & _
                  "		   Au.Taxa, Au.Autorizante, Au.EndAutorizante, ISNULL(Au.PedidoServico,0) PedidoServico, Au.SequenciaPedidoServico, CONVERT(varchar,convert(decimal(18,0),isnull(Ent.QuantidadeFisica,0))) as EntregueFisica, CONVERT(varchar,convert(decimal(18,0),isnull(Ent.QuantidadeFisica,0))) as EntregueFiscal, Au.QuantidadeFisica, Au.QuantidadeFiscal " & vbCrLf & _
                  "	  FROM AutorizacaoDeRetirada AS Au" & vbCrLf & _
                  "	  left Join (" & vbCrLf & _
                  "				SELECT RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao, " & vbCrLf & _
                  "					sum(Ro.PesoLiquido) AS QuantidadeFisica " & vbCrLf & _
                  "				  FROM Romaneios RO" & vbCrLf & _
                  "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
                  "					ON sb.Operacao_Id     = RO.Operacao " & vbCrLf & _
                  "				   AND sb.SubOperacoes_Id = RO.SubOperacao " & vbCrLf & _
                  "				 Group by RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao " & vbCrLf & _
                  "				 ) Ent" & vbCrLf & _
                  "		on Ent.Empresa_Id    = Au.Empresa_Id" & vbCrLf & _
                  "	   and Ent.EndEmpresa_Id = Au.EndEmpresa_Id" & vbCrLf & _
                  "	   and Ent.Pedido        = Au.Pedido_id" & vbCrLf & _
                  "	   and Ent.Autorizacao   = Au.Autorizacao_Id" & vbCrLf & _
                  "   inner join Pedidos P " & vbCrLf & _
                  "      on P.Empresa_id    = Au.Empresa_Id " & vbCrLf & _
                  "     and P.EndEmpresa_id = Au.EndEmpresa_id " & vbCrLf & _
                  "     and P.Pedido_id     = AU.Pedido_Id " & vbCrLf & _
                  "   inner join Clientes " & vbCrLf & _
                  "      on Au.ClienteRetirada    = Clientes.Cliente_id " & vbCrLf & _
                  "     and Au.EndClienteRetirada = Clientes.Endereco_Id " & vbCrLf & _
                  "  Where Au.Empresa_Id            ='" & pEmpresa & "'" & vbCrLf & _
                  "    and Au.EndEmpresa_Id         = " & pEndEmpresa & vbCrLf & _
                  "    and Au.Pedido_id             ='" & pPedido & "'" & vbCrLf & _
                  "    and isnull(P.FiscalAberto,1) = 1" & vbCrLf & _
                  "    and year(Au.Movimento) = '" & IIf(String.IsNullOrWhiteSpace(pAno), DateTime.Now.Year.ToString, pAno) & "'"
        Else
            Sql = "	SELECT Au.Empresa_Id, Au.EndEmpresa_id, Au.Pedido_id, Au.Autorizacao_Id, Au.Movimento, Au.DataBaseCalculo, Au.DataVencimento, Au.ClienteRetirada, Au.EndClienteRetirada, " & vbCrLf & _
                    "		   Au.Taxa, Au.Autorizante, Au.EndAutorizante, ISNULL(Au.PedidoServico,0) PedidoServico, Au.SequenciaPedidoServico, " & vbCrLf & _
                    "          case " & vbCrLf & _
                    "             when Ent.ControlarRomaneio = 1 " & vbCrLf & _
                    "                then " & vbCrLf & _
                    "					case " & vbCrLf & _
                    "						when isnull(REnt.QuantidadeFisica,0) = 0 " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS OrElse Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "							then isnull(EntDev.QuantidadeFisica,0) " & vbCrLf & _
                        "							else isnull(REntDev.QuantidadeFisica,0) " & vbCrLf
            Else
                Sql &= "							then isnull(Ent.QuantidadeFisica,0) - isnull(EntDEV.QuantidadeFisica,0) " & vbCrLf & _
                        "							else isnull(REnt.QuantidadeFisica,0) - isnull(REntDEV.QuantidadeFisica,0) " & vbCrLf
            End If

            Sql &= "					end " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS OrElse Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "                else isnull(EntDev.QuantidadeFisica,0) " & vbCrLf
            Else
                Sql &= "                else isnull(Ent.QuantidadeFisica,0) - isnull(EntDEV.QuantidadeFisica,0) " & vbCrLf
            End If

            Sql &= "          end AS EntregueFisica, " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS OrElse Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "          isnull(EntDev.QuantidadeFiscal,0) as EntregueFiscal, Au.QuantidadeFisica, Au.QuantidadeFiscal " & vbCrLf
            Else
                Sql &= "          isnull(Ent.QuantidadeFiscal,0) - isnull(EntDEV.QuantidadeFiscal,0) as EntregueFiscal, Au.QuantidadeFisica, Au.QuantidadeFiscal " & vbCrLf
            End If

            Sql &= "	  FROM AutorizacaoDeRetirada AS Au " & vbCrLf & _
                        "	  left Join ( " & vbCrLf & _
                        "				SELECT NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio, " & vbCrLf & _
                        "					sum(NFxI.QuantidadeFisica) AS QuantidadeFisica, " & vbCrLf & _
                        "					sum(NFxI.QuantidadeFiscal) AS QuantidadeFiscal " & vbCrLf & _
                        "				  FROM NotasFiscais NF " & vbCrLf & _
                        "				 INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf & _
                        "					ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
                        "				   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
                        "				   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
                        "				   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
                        "				   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
                        "				   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
                        "				   AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
                        "				 INNER JOIN Produtos P " & vbCrLf & _
                        "					ON NFxI.Produto_Id    = P.Produto_Id " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS OrElse Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
                       "					ON sb.Operacao_Id      = NF.Operacao " & vbCrLf & _
                       "					AND sb.SubOperacoes_Id = NF.SubOperacao " & vbCrLf & _
                       "					AND sb.Devolucao       = 'N' " & vbCrLf
            End If

            Sql &= "               WHERE NF.Situacao in (1,4,7) " & vbCrLf & _
                        "                 AND NF.TipoDeDocumento = 1 " & vbCrLf & _
                        "				Group by NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio " & vbCrLf & _
                        "				) Ent " & vbCrLf & _
                        "		on Au.Empresa_Id     = Ent.Empresa_Id " & vbCrLf & _
                        "	   and Au.EndEmpresa_Id  = Ent.EndEmpresa_Id " & vbCrLf & _
                        "	   and Au.Pedido_Id      = Ent.Pedido " & vbCrLf

            If Not (Classe = eClassesOperacoes.DEPOSITOS OrElse Classe = eClassesOperacoes.AFIXAR) Then
                Sql &= "	   and Au.Autorizacao_Id = Ent.Autorizacao " & vbCrLf
            End If

            Sql &= "	  left Join ( " & vbCrLf & _
                        "				SELECT NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio, " & vbCrLf & _
                        "					sum(NFxI.QuantidadeFisica) AS QuantidadeFisica, " & vbCrLf & _
                        "					sum(NFxI.QuantidadeFiscal) AS QuantidadeFiscal " & vbCrLf & _
                        "				  FROM NotasFiscais NF " & vbCrLf & _
                        "				 INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf & _
                        "					ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
                        "				   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
                        "				   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
                        "				   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
                        "				   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
                        "				   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
                        "				   AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
                        "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
                        "					ON sb.Operacao_Id     = NF.Operacao " & vbCrLf & _
                        "				   AND sb.SubOperacoes_Id = NF.SubOperacao " & vbCrLf

            If Not (Classe = eClassesOperacoes.DEPOSITOS OrElse Classe = eClassesOperacoes.AFIXAR) Then
                Sql &= "				   AND sb.PrecoFixo       = 'S' " & vbCrLf
            End If

            Sql &= "				   AND sb.Devolucao       = 'S' " & vbCrLf & _
                            "				 INNER JOIN Produtos P " & vbCrLf & _
                            "					ON NFxI.Produto_Id    = P.Produto_Id " & vbCrLf & _
                            "               WHERE NF.Situacao in (1,4,7) " & vbCrLf & _
                            "                 AND NF.TipoDeDocumento = 1 " & vbCrLf & _
                            "				Group by NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio " & vbCrLf & _
                            "				) EntDEV " & vbCrLf & _
                            "		on Au.Empresa_Id     = EntDEV.Empresa_Id " & vbCrLf & _
                            "	   and Au.EndEmpresa_Id  = EntDEV.EndEmpresa_Id " & vbCrLf & _
                            "	   and Au.Pedido_Id      = EntDEV.Pedido " & vbCrLf
            If Classe = eClassesOperacoes.DEPOSITOS OrElse Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "	   and Au.Autorizacao_Id = EntDEV.Autorizacao " & vbCrLf
            End If

            Sql &= "	  left Join ( " & vbCrLf & _
                            "				SELECT RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao, " & vbCrLf & _
                            "					sum(Ro.PesoLiquido) AS QuantidadeFisica " & vbCrLf & _
                            "				  FROM Romaneios RO " & vbCrLf & _
                            "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
                            "					ON sb.Operacao_Id     = RO.Operacao " & vbCrLf & _
                            "				   AND sb.SubOperacoes_Id = RO.SubOperacao " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS OrElse Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "				   AND sb.Devolucao       = 'N' " & vbCrLf
            End If

            Sql &= "                INNER JOIN NotasFiscaisXRomaneios NxR " & vbCrLf & _
                            "	                ON NxR.Empresa_id    = RO.Empresa_Id " & vbCrLf & _
                            "                  AND NxR.EndEmpresa_Id = Ro.EndEmpresa_Id " & vbCrLf & _
                            "                  AND NxR.Romaneio_id   = Ro.Romaneio_Id " & vbCrLf & _
                            "                 JOIN NotasFiscais Nf " & vbCrLf & _
                            "                   ON NxR.Empresa_id    = nf.Empresa_Id " & vbCrLf & _
                            "                  AND NxR.EndEmpresa_Id = nf.EndEmpresa_Id " & vbCrLf & _
                            "                  AND NXR.Cliente_Id = NF.Cliente_Id" & vbCrLf & _
                            "                  AND NXR.EndCliente_Id = NF.EndCliente_Id" & vbCrLf & _
                            "                  AND NXR.EntradaSaida_Id = NF.Entradasaida_Id" & vbCrLf & _
                            "                  AND NXR.Serie_Id        = NF.Serie_Id" & vbCrLf & _
                            "                  AND NXR.Nota_Id         = NF.Nota_Id" & vbCrLf & _
                            "                WHERE NF.Situacao in (1,4,7)" & vbCrLf & _
                            "                  AND NF.TipoDeDocumento = 1 " & vbCrLf & _
                            "			     Group by RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao " & vbCrLf & _
                            "				 ) REnt " & vbCrLf & _
                            "		on REnt.Empresa_Id    = Au.Empresa_Id " & vbCrLf & _
                            "	   and REnt.EndEmpresa_Id = Au.EndEmpresa_Id " & vbCrLf & _
                            "	   and REnt.Pedido        = Au.Pedido_id " & vbCrLf

            If Not (Classe = eClassesOperacoes.DEPOSITOS OrElse Classe = eClassesOperacoes.AFIXAR) Then
                Sql &= "	   and REnt.Autorizacao   = Au.Autorizacao_Id " & vbCrLf
            End If

            Sql &= "	  left Join ( " & vbCrLf & _
                            "				SELECT RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao, " & vbCrLf & _
                            "					sum(Ro.PesoLiquido) AS QuantidadeFisica " & vbCrLf & _
                            "				  FROM Romaneios RO " & vbCrLf & _
                            "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
                            "					ON sb.Operacao_Id     = RO.Operacao " & vbCrLf & _
                            "				   AND sb.SubOperacoes_Id = RO.SubOperacao " & vbCrLf
            If Not (Classe = eClassesOperacoes.DEPOSITOS OrElse Classe = eClassesOperacoes.AFIXAR) Then
                Sql &= "				   AND sb.PrecoFixo       = 'S' " & vbCrLf
            End If

            Sql &= "				   AND sb.Devolucao       = 'S' " & vbCrLf & _
                            "                INNER JOIN NotasFiscaisXRomaneios NxR " & vbCrLf & _
                            "	                ON NxR.Empresa_id    = RO.Empresa_Id " & vbCrLf & _
                            "                  AND NxR.EndEmpresa_Id = Ro.EndEmpresa_Id " & vbCrLf & _
                            "                  AND NxR.Romaneio_id   = Ro.Romaneio_Id " & vbCrLf & _
                            "                 JOIN NotasFiscais Nf " & vbCrLf & _
                            "                   ON NxR.Empresa_id    = nf.Empresa_Id " & vbCrLf & _
                            "                  AND NxR.EndEmpresa_Id = nf.EndEmpresa_Id " & vbCrLf & _
                            "                  AND NxR.Nota_Id = nf.Nota_Id   " & vbCrLf & _
                            "                  AND NXR.Cliente_Id = NF.Cliente_Id" & vbCrLf & _
                            "                  AND NXR.EndCliente_Id = NF.EndCliente_Id" & vbCrLf & _
                             "                  AND NXR.EntradaSaida_Id = NF.Entradasaida_Id" & vbCrLf & _
                            "                  AND NXR.Serie_Id        = NF.Serie_Id" & vbCrLf & _
                            "                WHERE NF.Situacao in (1,4,7)" & vbCrLf & _
                            "                  AND NF.TipoDeDocumento = 1 " & vbCrLf & _
                            "				 Group by RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao " & vbCrLf & _
                            "				 ) REntDEV " & vbCrLf & _
                            "		on REntDEV.Empresa_Id    = Au.Empresa_Id " & vbCrLf & _
                            "	   and REntDEV.EndEmpresa_Id = Au.EndEmpresa_Id " & vbCrLf & _
                            "	   and REntDEV.Pedido        = Au.Pedido_id " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS OrElse Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "	   and REntDEV.Autorizacao   = Au.Autorizacao_Id " & vbCrLf
            End If

            Sql &= "   inner join Pedidos P " & vbCrLf & _
                        "      on P.Empresa_id    = Au.Empresa_Id " & vbCrLf & _
                        "     and P.EndEmpresa_id = Au.EndEmpresa_id " & vbCrLf & _
                        "     and P.Pedido_id     = AU.Pedido_Id " & vbCrLf & _
                        "  Where Au.Empresa_Id            ='" & pEmpresa & "'" & vbCrLf & _
                        "    and Au.EndEmpresa_Id         = " & pEndEmpresa & vbCrLf & _
                        "    and isnull(P.FiscalAberto,1) = 1" & vbCrLf
            '"    and year(Au.Movimento)       = '" & IIf(String.IsNullOrWhiteSpace(pAno), DateTime.Now.Year.ToString, pAno) & "'"

            If pPedido > 0 Then
                Sql &= "    and Au.Pedido_id      ='" & pPedido & "'" & vbCrLf
            End If

            If pcliente.Length > 0 Then
                Sql &= "    and P.Cliente    ='" & pcliente & "'" & vbCrLf & _
                       "    and P.EndCliente = " & pEndCliente & vbCrLf
            End If
        End If

        Sql &= " order by Au.Autorizacao_id "

        ds = Banco.ConsultaDataSet(Sql, "Autorizacao")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Au As New AutorizacaoDeRetirada
            Au.IUD = "U"
            Au.CodigoEmpresa = row("Empresa_Id")
            Au.EnderecoEmpresa = row("EndEmpresa_id")
            Au.CodigoPedido = row("Pedido_id")
            'Au.Sequencia = row("Sequencia_id")
            Au.Autorizacao = row("Autorizacao_Id")

            Au.Movimento = row("Movimento")
            Au.DataBaseCalculo = row("DataBaseCalculo")
            Au.DataVencimento = row("DataVencimento")

            Au.CodigoClienteRetirada = row("ClienteRetirada")
            Au.EnderecoClienteRetirada = row("EndClienteRetirada")
            Au.Taxa = row("Taxa")
            Au.CodigoAutorizante = row("Autorizante")
            Au.EnderecoAutorizante = row("EndAutorizante")
            Au.CodigoPedidoServico = row("PedidoServico")
            Au.SequenciaPedidoServico = row("SequenciaPedidoServico")
            Au.QuantidadeAutorizadaFiscal = row("QuantidadeFiscal")
            Au.QuantidadeAutorizadaFisica = row("QuantidadeFisica")
            Au.QuantidadeEntregueFisica = row("EntregueFisica")
            Au.QuantidadeEntregueFiscal = row("EntregueFiscal")

            Me.Add(Au)
            _QtdeAutorizadaFisica += Au.QuantidadeAutorizadaFisica
            _QtdeAutorizadaFiscal += Au.QuantidadeAutorizadaFiscal
            _QtdeEntregueFisica += Au.QuantidadeEntregueFisica
            _QtdeEntregueFiscal += Au.QuantidadeEntregueFiscal
            _QtdeSaldoFisico += Au.SaldoFisico
            _QtdeSaldoFiscal += Au.SaldoFiscal
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class AutorizacaoDeRetirada

#Region "Variaveis Auxiliares"
    Private Sql As String
    Private Banco As New AcessaBanco
    Private ds As DataSet
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer)
        Dim n As New Numerador(pEmpresa, pEndEmpresa, 50)
        _Autorizacao = n.Sequencia
    End Sub

    Public Sub New(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pPedido As Integer, ByVal pAutorizacao As Integer, ByVal Classe As [Lib].Negocio.eClassesOperacoes, Optional ByVal pRomaneio As Boolean = False)
        If pRomaneio = True Then
            Sql = "	SELECT Au.Empresa_Id, Au.EndEmpresa_id, Au.Pedido_id, Au.Autorizacao_Id, Au.Movimento, Au.DataBaseCalculo, Au.DataVencimento, Au.ClienteRetirada, Au.EndClienteRetirada, " & vbCrLf & _
                  " isnull(Au.Observacao, '') AS Observacao, " & vbCrLf & _
                  "		   Au.Taxa, Au.Autorizante, Au.EndAutorizante, ISNULL(Au.PedidoServico,0) PedidoServico, Au.SequenciaPedidoServico, CONVERT(varchar,convert(decimal(18,0),isnull(Ent.QuantidadeFisica,0))) as EntregueFisica, CONVERT(varchar,convert(decimal(18,0),isnull(Ent.QuantidadeFisica,0))) as EntregueFiscal, Au.QuantidadeFisica, Au.QuantidadeFiscal, Au.UsuarioDeInclusao, Au.UsuarioInclusaoData, isnull(Au.UsuarioDeAlteracao,'') as UsuarioDeAlteracao, isnull(Au.UsuarioAlteracaoData,'') as UsuarioAlteracaoData  " & vbCrLf & _
                  "	  FROM AutorizacaoDeRetirada AS Au" & vbCrLf & _
                  "	  left Join (" & vbCrLf & _
                  "				SELECT RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao, " & vbCrLf & _
                  "					sum(Ro.PesoLiquido) AS QuantidadeFisica " & vbCrLf & _
                  "				  FROM Romaneios RO" & vbCrLf & _
                  "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
                  "					ON sb.Operacao_Id     = RO.Operacao " & vbCrLf & _
                  "				   AND sb.SubOperacoes_Id = RO.SubOperacao " & vbCrLf & _
                  "				 Group by RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao " & vbCrLf & _
                  "				 ) Ent" & vbCrLf & _
                  "		on Ent.Empresa_Id    = Au.Empresa_Id" & vbCrLf & _
                  "	   and Ent.EndEmpresa_Id = Au.EndEmpresa_Id" & vbCrLf & _
                  "	   and Ent.Pedido        = Au.Pedido_id" & vbCrLf & _
                  "	   and Ent.Autorizacao   = Au.Autorizacao_Id" & vbCrLf & _
                  "   inner join Pedidos P " & vbCrLf & _
                  "      on P.Empresa_id    = Au.Empresa_Id " & vbCrLf & _
                  "     and P.EndEmpresa_id = Au.EndEmpresa_id " & vbCrLf & _
                  "     and P.Pedido_id     = AU.Pedido_Id " & vbCrLf & _
                  "   inner join Clientes " & vbCrLf & _
                  "      on Au.ClienteRetirada    = Clientes.Cliente_id " & vbCrLf & _
                  "     and Au.EndClienteRetirada = Clientes.Endereco_Id " & vbCrLf & _
                  "  Where Au.Empresa_Id     ='" & pEmpresa & "'" & vbCrLf & _
                  "    and Au.EndEmpresa_Id  = " & pEndEmpresa & vbCrLf & _
                  "    and AU.Pedido_id      = " & pPedido & vbCrLf & _
                  "    and Au.Autorizacao_id = " & pAutorizacao
        Else
            'Sql = "	SELECT Au.Empresa_Id, Au.EndEmpresa_id, Au.Pedido_id, Au.Sequencia_id, Au.Autorizacao_Id, Au.Movimento, Au.DataBaseCalculo, Au.DataVencimento, Au.ClienteRetirada, Au.EndClienteRetirada, " & vbCrLf & _
            '        "		   isnull(Au.Observacao, '') AS Observacao, Au.Taxa, Au.Autorizante, Au.EndAutorizante, Au.PedidoServico, Au.SequenciaPedidoServico, " & vbCrLf & _
            '        "          case " & vbCrLf & _
            '        "             when Ent.ControlarRomaneio = 1 " & vbCrLf & _
            '        "                then " & vbCrLf & _
            '        "					case " & vbCrLf & _
            '        "						when isnull(REnt.QuantidadeFisica,0) = 0 " & vbCrLf & _
            '        "							then isnull(Ent.QuantidadeFisica,0) - isnull(EntDEV.QuantidadeFisica,0) " & vbCrLf & _
            '        "							else isnull(REnt.QuantidadeFisica,0) - isnull(REntDEV.QuantidadeFisica,0) " & vbCrLf & _
            '        "					end " & vbCrLf & _
            '        "                else isnull(Ent.QuantidadeFisica,0) - isnull(EntDEV.QuantidadeFisica,0) " & vbCrLf & _
            '        "          end AS EntregueFisica, " & vbCrLf & _
            '        "          isnull(Ent.QuantidadeFiscal,0) - isnull(EntDEV.QuantidadeFiscal,0) as EntregueFiscal, Au.QuantidadeFisica, Au.QuantidadeFiscal, " & vbCrLf & _
            '        "          Au.UsuarioDeInclusao, Au.UsuarioInclusaoData, isnull(Au.UsuarioDeAlteracao,'') as UsuarioDeAlteracao, isnull(Au.UsuarioAlteracaoData,'') as UsuarioAlteracaoData " & vbCrLf & _
            '        "	  FROM AutorizacaoDeRetirada AS Au " & vbCrLf & _
            '        "	  left Join ( " & vbCrLf & _
            '        "				SELECT NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio, " & vbCrLf & _
            '        "					sum(NFxI.QuantidadeFisica) AS QuantidadeFisica, " & vbCrLf & _
            '        "					sum(NFxI.QuantidadeFiscal) AS QuantidadeFiscal " & vbCrLf & _
            '        "				  FROM NotasFiscais NF " & vbCrLf & _
            '        "				 INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf & _
            '        "					ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
            '        "				   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
            '        "				   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
            '        "				   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
            '        "				   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
            '        "				   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
            '        "				   AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
            '        "				 INNER JOIN Produtos P " & vbCrLf & _
            '        "					ON NFxI.Produto_Id    = P.Produto_Id " & vbCrLf & _
            '        "               WHERE NF.Situacao in (1,4,7) " & vbCrLf & _
            '        "                 AND NF.TipoDeDocumento = 1 " & vbCrLf & _
            '        "				Group by NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio " & vbCrLf & _
            '        "				) Ent " & vbCrLf & _
            '        "		on Au.Empresa_Id     = Ent.Empresa_Id " & vbCrLf & _
            '        "	   and Au.EndEmpresa_Id  = Ent.EndEmpresa_Id " & vbCrLf & _
            '        "	   and Au.Pedido_Id      = Ent.Pedido " & vbCrLf & _
            '        "	   and Au.Autorizacao_Id = Ent.Autorizacao " & vbCrLf & _
            '        "	  left Join ( " & vbCrLf & _
            '        "				SELECT NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio, " & vbCrLf & _
            '        "					sum(NFxI.QuantidadeFisica) AS QuantidadeFisica, " & vbCrLf & _
            '        "					sum(NFxI.QuantidadeFiscal) AS QuantidadeFiscal " & vbCrLf & _
            '        "				  FROM NotasFiscais NF " & vbCrLf & _
            '        "				 INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf & _
            '        "					ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
            '        "				   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
            '        "				   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
            '        "				   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
            '        "				   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
            '        "				   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
            '        "				   AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
            '        "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
            '        "					ON sb.Operacao_Id     = NF.Operacao " & vbCrLf & _
            '        "				   AND sb.SubOperacoes_Id = NF.SubOperacao " & vbCrLf & _
            '        "				   AND sb.PrecoFixo       = 'S' " & vbCrLf & _
            '        "				   AND sb.Devolucao       = 'S' " & vbCrLf & _
            '        "				 INNER JOIN Produtos P " & vbCrLf & _
            '        "					ON NFxI.Produto_Id    = P.Produto_Id " & vbCrLf & _
            '        "               WHERE NF.Situacao in (1,4,7) " & vbCrLf & _
            '        "                 AND NF.TipoDeDocumento = 1 " & vbCrLf & _
            '        "				Group by NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio " & vbCrLf & _
            '        "				) EntDEV " & vbCrLf & _
            '        "		on Au.Empresa_Id     = EntDEV.Empresa_Id " & vbCrLf & _
            '        "	   and Au.EndEmpresa_Id  = EntDEV.EndEmpresa_Id " & vbCrLf & _
            '        "	   and Au.Pedido_Id      = EntDEV.Pedido " & vbCrLf & _
            '        "	  left Join ( " & vbCrLf & _
            '        "				SELECT RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao, " & vbCrLf & _
            '        "					sum(Ro.PesoLiquido) AS QuantidadeFisica " & vbCrLf & _
            '        "				  FROM Romaneios RO " & vbCrLf & _
            '        "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
            '        "					ON sb.Operacao_Id     = RO.Operacao " & vbCrLf & _
            '        "				   AND sb.SubOperacoes_Id = RO.SubOperacao " & vbCrLf & _
            '        "                INNER JOIN NotasFiscaisXRomaneios NxR " & vbCrLf & _
            '        "	                ON NxR.Empresa_id    = RO.Empresa_Id " & vbCrLf & _
            '        "                  AND NxR.EndEmpresa_Id = Ro.EndEmpresa_Id " & vbCrLf & _
            '        "                  AND NxR.Romaneio_id   = Ro.Romaneio_Id " & vbCrLf & _
            '        "                 JOIN NotasFiscais NF " & vbCrLf & _
            '        "                   ON NF.Empresa_Id = NxR.Empresa_Id " & vbCrLf & _
            '        "                  AND NF.EndEmpresa_Id = NxR.EndEmpresa_Id " & vbCrLf & _
            '        "                  AND NF.Cliente_Id = NxR.Cliente_Id " & vbCrLf & _
            '        "                  AND NF.EndCliente_Id = NxR.EndCliente_Id " & vbCrLf & _
            '        "                  AND NF.EntradaSaida_Id = NxR.EntradaSaida_Id " & vbCrLf & _
            '        "                  AND NF.Serie_Id = NxR.Serie_Id " & vbCrLf & _
            '        "                  AND NF.Nota_Id = NxR.Nota_Id" & vbCrLf & _
            '        "                WHERE NF.Situacao in (1,4,7) " & vbCrLf & _
            '        "                  AND NF.TipoDeDocumento = 1 " & vbCrLf & _
            '        "				 Group by RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao " & vbCrLf & _
            '        "				 ) REnt " & vbCrLf & _
            '        "		on REnt.Empresa_Id    = Au.Empresa_Id " & vbCrLf & _
            '        "	   and REnt.EndEmpresa_Id = Au.EndEmpresa_Id " & vbCrLf & _
            '        "	   and REnt.Pedido        = Au.Pedido_id " & vbCrLf & _
            '        "	   and REnt.Autorizacao   = Au.Autorizacao_Id " & vbCrLf & _
            '        "	  left Join ( " & vbCrLf & _
            '        "				SELECT RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao, " & vbCrLf & _
            '        "					sum(Ro.PesoLiquido) AS QuantidadeFisica " & vbCrLf & _
            '        "				  FROM Romaneios RO " & vbCrLf & _
            '        "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
            '        "					ON sb.Operacao_Id     = RO.Operacao " & vbCrLf & _
            '        "				   AND sb.SubOperacoes_Id = RO.SubOperacao " & vbCrLf & _
            '        "				   AND sb.PrecoFixo       = 'S' " & vbCrLf & _
            '        "				   AND sb.Devolucao       = 'S' " & vbCrLf & _
            '        "                INNER JOIN NotasFiscaisXRomaneios NxR " & vbCrLf & _
            '        "	                ON NxR.Empresa_id    = RO.Empresa_Id " & vbCrLf & _
            '        "                  AND NxR.EndEmpresa_Id = Ro.EndEmpresa_Id " & vbCrLf & _
            '        "                  AND NxR.Romaneio_id   = Ro.Romaneio_Id " & vbCrLf & _
            '        "                 JOIN NotasFiscais NF " & vbCrLf & _
            '        "                   ON NF.Empresa_Id = NxR.Empresa_Id " & vbCrLf & _
            '        "                  AND NF.EndEmpresa_Id = NxR.EndEmpresa_Id " & vbCrLf & _
            '        "                  AND NF.Cliente_Id = NxR.Cliente_Id " & vbCrLf & _
            '        "                  AND NF.EndCliente_Id = NxR.EndCliente_Id " & vbCrLf & _
            '        "                  AND NF.EntradaSaida_Id = NxR.EntradaSaida_Id " & vbCrLf & _
            '        "                  AND NF.Serie_Id = NxR.Serie_Id " & vbCrLf & _
            '        "                  AND NF.Nota_Id = NxR.Nota_Id" & vbCrLf & _
            '        "                WHERE NF.Situacao in (1,4,7) " & vbCrLf & _
            '        "                  AND NF.TipoDeDocumento = 1 " & vbCrLf & _
            '        "				 Group by RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao " & vbCrLf & _
            '        "				 ) REntDEV " & vbCrLf & _
            '        "		on REntDEV.Empresa_Id    = Au.Empresa_Id " & vbCrLf & _
            '        "	   and REntDEV.EndEmpresa_Id = Au.EndEmpresa_Id " & vbCrLf & _
            '        "	   and REntDEV.Pedido        = Au.Pedido_id " & vbCrLf & _
            '        "   inner join Pedidos P " & vbCrLf & _
            '        "      on P.Empresa_id    = Au.Empresa_Id " & vbCrLf & _
            '        "     and P.EndEmpresa_id = Au.EndEmpresa_id " & vbCrLf & _
            '        "     and P.Pedido_id     = AU.Pedido_Id " & vbCrLf & _
            '        "  Where Au.Empresa_Id     ='" & pEmpresa & "'" & vbCrLf & _
            '        "    and Au.EndEmpresa_Id  = " & pEndEmpresa & vbCrLf & _
            '        "    and Au.Autorizacao_id = " & pAutorizacao

            Sql = "	SELECT Au.Empresa_Id, Au.EndEmpresa_id, Au.Pedido_id, Au.Autorizacao_Id, Au.Movimento, Au.DataBaseCalculo, Au.DataVencimento, Au.ClienteRetirada, Au.EndClienteRetirada, " & vbCrLf & _
                    "	isnull(Au.Observacao, '') AS Observacao, Au.Taxa, Au.Autorizante, Au.EndAutorizante, ISNULL(Au.PedidoServico,0) PedidoServico, Au.SequenciaPedidoServico, " & vbCrLf & _
                    "          case " & vbCrLf & _
                    "             when Ent.ControlarRomaneio = 1 " & vbCrLf & _
                    "                then " & vbCrLf & _
                    "					case " & vbCrLf & _
                    "						when isnull(REnt.QuantidadeFisica,0) = 0 " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS Or Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "							then isnull(EntDev.QuantidadeFisica,0) " & vbCrLf & _
                        "							else isnull(REntDev.QuantidadeFisica,0) " & vbCrLf
            Else
                Sql &= "							then isnull(Ent.QuantidadeFisica,0) - isnull(EntDEV.QuantidadeFisica,0) " & vbCrLf & _
                        "							else isnull(REnt.QuantidadeFisica,0) - isnull(REntDEV.QuantidadeFisica,0) " & vbCrLf
            End If

            Sql &= "					end " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS Or Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "                else isnull(EntDev.QuantidadeFisica,0) " & vbCrLf
            Else
                Sql &= "                else isnull(Ent.QuantidadeFisica,0) - isnull(EntDEV.QuantidadeFisica,0) " & vbCrLf
            End If

            Sql &= "          end AS EntregueFisica, " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS Or Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "          isnull(EntDev.QuantidadeFiscal,0) as EntregueFiscal, Au.QuantidadeFisica, Au.QuantidadeFiscal, " & vbCrLf
            Else
                Sql &= "          isnull(Ent.QuantidadeFiscal,0) - isnull(EntDEV.QuantidadeFiscal,0) as EntregueFiscal, Au.QuantidadeFisica, Au.QuantidadeFiscal," & vbCrLf
            End If

            Sql &= "  Au.UsuarioDeInclusao, Au.UsuarioInclusaoData, isnull(Au.UsuarioDeAlteracao,'') as UsuarioDeAlteracao, isnull(Au.UsuarioAlteracaoData,'') as UsuarioAlteracaoData " & vbCrLf & _
                    "	  FROM AutorizacaoDeRetirada AS Au " & vbCrLf & _
                    "	  left Join ( " & vbCrLf & _
                    "				SELECT NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio, " & vbCrLf & _
                    "					sum(NFxI.QuantidadeFisica) AS QuantidadeFisica, " & vbCrLf & _
                    "					sum(NFxI.QuantidadeFiscal) AS QuantidadeFiscal " & vbCrLf & _
                    "				  FROM NotasFiscais NF " & vbCrLf & _
                    "				 INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf & _
                    "					ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
                    "				   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
                    "				   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
                    "				   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
                    "				   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
                    "				   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
                    "				   AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
                    "				 INNER JOIN Produtos P " & vbCrLf & _
                    "					ON NFxI.Produto_Id    = P.Produto_Id " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS Or Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
                       "					ON sb.Operacao_Id      = NF.Operacao " & vbCrLf & _
                       "					AND sb.SubOperacoes_Id = NF.SubOperacao " & vbCrLf & _
                       "					AND sb.Devolucao       = 'N' " & vbCrLf
            End If

            Sql &= "               WHERE NF.Situacao in (1,4,7) " & vbCrLf & _
                    "                 AND NF.TipoDeDocumento = 1 " & vbCrLf & _
                    "				Group by NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio " & vbCrLf & _
                    "				) Ent " & vbCrLf & _
                    "		on Au.Empresa_Id     = Ent.Empresa_Id " & vbCrLf & _
                    "	   and Au.EndEmpresa_Id  = Ent.EndEmpresa_Id " & vbCrLf & _
                    "	   and Au.Pedido_Id      = Ent.Pedido " & vbCrLf

            If Not Classe = eClassesOperacoes.DEPOSITOS Or Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "	   and Au.Autorizacao_Id = Ent.Autorizacao " & vbCrLf
            End If

            Sql &= "	  left Join ( " & vbCrLf & _
                    "				SELECT NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio, " & vbCrLf & _
                    "					sum(NFxI.QuantidadeFisica) AS QuantidadeFisica, " & vbCrLf & _
                    "					sum(NFxI.QuantidadeFiscal) AS QuantidadeFiscal " & vbCrLf & _
                    "				  FROM NotasFiscais NF " & vbCrLf & _
                    "				 INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf & _
                    "					ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
                    "				   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
                    "				   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
                    "				   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
                    "				   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
                    "				   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
                    "				   AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
                    "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
                    "					ON sb.Operacao_Id     = NF.Operacao " & vbCrLf & _
                    "				   AND sb.SubOperacoes_Id = NF.SubOperacao " & vbCrLf

            If Not Classe = eClassesOperacoes.DEPOSITOS Or Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "				   AND sb.PrecoFixo       = 'S' " & vbCrLf
            End If

            Sql &= "				   AND sb.Devolucao       = 'S' " & vbCrLf & _
                    "				 INNER JOIN Produtos P " & vbCrLf & _
                    "					ON NFxI.Produto_Id    = P.Produto_Id " & vbCrLf & _
                    "               WHERE NF.Situacao in (1,4,7) " & vbCrLf & _
                    "                 AND NF.TipoDeDocumento = 1 " & vbCrLf & _
                    "				Group by NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Autorizacao, P.ControlarRomaneio " & vbCrLf & _
                    "				) EntDEV " & vbCrLf & _
                    "		on Au.Empresa_Id     = EntDEV.Empresa_Id " & vbCrLf & _
                    "	   and Au.EndEmpresa_Id  = EntDEV.EndEmpresa_Id " & vbCrLf & _
                    "	   and Au.Pedido_Id      = EntDEV.Pedido " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS Or Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "	   and Au.Autorizacao_Id = EntDEV.Autorizacao " & vbCrLf
            End If

            Sql &= "	  left Join ( " & vbCrLf & _
                    "				SELECT RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao, " & vbCrLf & _
                    "					sum(Ro.PesoLiquido) AS QuantidadeFisica " & vbCrLf & _
                    "				  FROM Romaneios RO " & vbCrLf & _
                    "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
                    "					ON sb.Operacao_Id     = RO.Operacao " & vbCrLf & _
                    "				   AND sb.SubOperacoes_Id = RO.SubOperacao " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS Or Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "				   AND sb.Devolucao       = 'N' " & vbCrLf
            End If

            Sql &= "                INNER JOIN NotasFiscaisXRomaneios NxR " & vbCrLf & _
                    "	                ON NxR.Empresa_id    = RO.Empresa_Id " & vbCrLf & _
                    "                  AND NxR.EndEmpresa_Id = Ro.EndEmpresa_Id " & vbCrLf & _
                    "                  AND NxR.Romaneio_id   = Ro.Romaneio_Id " & vbCrLf & _
                    "                 JOIN NotasFiscais Nf " & vbCrLf & _
                    "                   ON NxR.Empresa_id    = nf.Empresa_Id " & vbCrLf & _
                    "                  AND NxR.EndEmpresa_Id = nf.EndEmpresa_Id " & vbCrLf & _
                    "                  AND NxR.Nota_Id = nf.Nota_Id   " & vbCrLf & _
                    "                  AND NXR.Cliente_Id = NF.Cliente_Id" & vbCrLf & _
                    "                  AND NXR.EndCliente_Id = NF.EndCliente_Id" & vbCrLf & _
                    "                  AND NXR.Serie_Id        = NF.Serie_Id" & vbCrLf & _
                    "                  AND nxr.EntradaSaida_Id = nf.EntradaSaida_Id" & vbCrLf & _
                    "                WHERE NF.Situacao in (1,4,7)" & vbCrLf & _
                    "                  AND NF.TipoDeDocumento = 1 " & vbCrLf & _
                    "			     Group by RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao " & vbCrLf & _
                    "				 ) REnt " & vbCrLf & _
                    "		on REnt.Empresa_Id    = Au.Empresa_Id " & vbCrLf & _
                    "	   and REnt.EndEmpresa_Id = Au.EndEmpresa_Id " & vbCrLf & _
                    "	   and REnt.Pedido        = Au.Pedido_id " & vbCrLf

            If Not Classe = eClassesOperacoes.DEPOSITOS Or Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "	   and REnt.Autorizacao   = Au.Autorizacao_Id " & vbCrLf
            End If

            Sql &= "	  left Join ( " & vbCrLf & _
                    "				SELECT RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao, " & vbCrLf & _
                    "					sum(Ro.PesoLiquido) AS QuantidadeFisica " & vbCrLf & _
                    "				  FROM Romaneios RO " & vbCrLf & _
                    "				 INNER JOIN SubOperacoes sb " & vbCrLf & _
                    "					ON sb.Operacao_Id     = RO.Operacao " & vbCrLf & _
                    "				   AND sb.SubOperacoes_Id = RO.SubOperacao " & vbCrLf

            If Not Classe = eClassesOperacoes.DEPOSITOS Or Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "				   AND sb.PrecoFixo       = 'S' " & vbCrLf
            End If

            Sql &= "				   AND sb.Devolucao       = 'S' " & vbCrLf & _
                    "                INNER JOIN NotasFiscaisXRomaneios NxR " & vbCrLf & _
                    "	                ON NxR.Empresa_id    = RO.Empresa_Id " & vbCrLf & _
                    "                  AND NxR.EndEmpresa_Id = Ro.EndEmpresa_Id " & vbCrLf & _
                    "                  AND NxR.Romaneio_id   = Ro.Romaneio_Id " & vbCrLf & _
                    "                 JOIN NotasFiscais Nf " & vbCrLf & _
                    "                   ON NxR.Empresa_id    = nf.Empresa_Id " & vbCrLf & _
                    "                  AND NxR.EndEmpresa_Id = nf.EndEmpresa_Id " & vbCrLf & _
                    "                  AND NxR.Nota_Id = nf.Nota_Id   " & vbCrLf & _
                    "                  AND NXR.Cliente_Id = NF.Cliente_Id" & vbCrLf & _
                    "                  AND NXR.EndCliente_Id = NF.EndCliente_Id" & vbCrLf & _
                    "                  AND NXR.Serie_Id        = NF.Serie_Id" & vbCrLf & _
                    "                  AND nxr.EntradaSaida_Id = nf.EntradaSaida_Id" & vbCrLf & _
                    "                WHERE NF.Situacao in (1,4,7)" & vbCrLf & _
                    "                  AND NF.TipoDeDocumento = 1 " & vbCrLf & _
                    "				 Group by RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, RO.Autorizacao " & vbCrLf & _
                    "				 ) REntDEV " & vbCrLf & _
                    "		on REntDEV.Empresa_Id    = Au.Empresa_Id " & vbCrLf & _
                    "	   and REntDEV.EndEmpresa_Id = Au.EndEmpresa_Id " & vbCrLf & _
                    "	   and REntDEV.Pedido        = Au.Pedido_id " & vbCrLf

            If Classe = eClassesOperacoes.DEPOSITOS Or Classe = eClassesOperacoes.AFIXAR Then
                Sql &= "	   and REntDEV.Autorizacao   = Au.Autorizacao_Id " & vbCrLf
            End If

            Sql &= "   inner join Pedidos P " & vbCrLf & _
                    "      on P.Empresa_id    = Au.Empresa_Id " & vbCrLf & _
                    "     and P.EndEmpresa_id = Au.EndEmpresa_id " & vbCrLf & _
                    "     and P.Pedido_id     = AU.Pedido_Id " & vbCrLf & _
                    "  Where Au.Empresa_Id     ='" & pEmpresa & "'" & vbCrLf & _
                    "    and Au.EndEmpresa_Id  = " & pEndEmpresa & vbCrLf & _
                    "    and Au.Pedido_id      = " & pPedido & vbCrLf & _
                    "    and Au.Autorizacao_id = " & pAutorizacao
        End If

        ds = Banco.ConsultaDataSet(Sql, "Autorizacao")

        For Each row As DataRow In ds.Tables(0).Rows
            _IUD = "U"
            _CodigoEmpresa = row("Empresa_Id")
            _EnderecoEmpresa = row("EndEmpresa_id")
            _CodigoPedido = row("Pedido_id")
            '_Sequencia = row("Sequencia_id")
            _Autorizacao = row("Autorizacao_Id")

            _Movimento = row("Movimento")
            _DataBaseCalculo = row("DataBaseCalculo")
            _DataVencimento = row("DataVencimento")

            _CodigoClienteRetirada = row("ClienteRetirada")
            _EnderecoClienteRetirada = row("EndClienteRetirada")
            _Taxa = row("Taxa")
            _CodigoAutorizante = row("Autorizante")
            _EnderecoAutorizante = row("EndAutorizante")
            _CodigoPedidoServico = row("PedidoServico")
            _SequenciaPedidoServico = row("SequenciaPedidoServico")
            _QuantidadeAutorizadaFiscal = row("QuantidadeFiscal")
            _QuantidadeAutorizadaFisica = row("QuantidadeFisica")
            _QuantidadeEntregueFisica = row("EntregueFisica")
            _QuantidadeEntregueFiscal = row("EntregueFiscal")
            _UsuarioDeInclusao = row("UsuarioDeInclusao")
            _UsuarioInclusaoData = row("UsuarioInclusaoData")
            _UsuarioDeAlteracao = row("UsuarioDeAlteracao")
            _UsuarioAlteracaoData = row("UsuarioAlteracaoData")
            _Observacao = row("Observacao")
        Next
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""

    Private _Empresa As Negocio.Cliente
    Private _CodigoEmpresa As String = ""
    Private _EnderecoEmpresa As Integer = 0

    Private _Pedido As Negocio.Pedido
    Private _CodigoPedido As Integer = 0
    'Private _Sequencia As Integer = 0

    Private _Autorizacao As Integer = 0

    Private _Movimento As Date = Date.Now
    Private _DataBaseCalculo As Date = Date.Now
    Private _DataVencimento As Date = Date.Now

    Private _ClienteRetirada As Negocio.Cliente
    Private _CodigoClienteRetirada As String = ""
    Private _EnderecoClienteRetirada As Integer = 0

    Private _Taxa As Double = 0

    Private _ClienteAutorizante As Negocio.Cliente
    Private _CodigoAutorizante As String = ""
    Private _EnderecoAutorizante As Integer = 0

    Private _PedidoServico As Negocio.Pedido
    Private _CodigoPedidoServico As Integer = 0
    Private _SequenciaPedidoServico As Integer = 0

    Private _QuantidadeAutorizadaFisica As Double = 0
    Private _QuantidadeEntregueFisica As Double = 0
    Private _QuantidadeAutorizadaFiscal As Double = 0
    Private _QuantidadeEntregueFiscal As Double = 0

    Private _Observacao As String = ""

    Private _UsuarioDeInclusao As String
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioDeAlteracao As String
    Private _UsuarioAlteracaoData As DateTime
#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Empresa() As Negocio.Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Negocio.Cliente(_CodigoEmpresa, _EnderecoEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
        End Set
    End Property

    Public ReadOnly Property NomeEmpresa() As String
        Get
            Return Empresa.Nome
        End Get
    End Property

    Public Property Pedido() As Negocio.Pedido
        Get
            If _Pedido Is Nothing And _CodigoPedido > 0 Then _Pedido = New Negocio.Pedido(_CodigoEmpresa, _EnderecoEmpresa, _CodigoPedido)
            Return _Pedido
        End Get
        Set(ByVal value As Negocio.Pedido)
            _Pedido = value
        End Set
    End Property

    Public Property CodigoPedido() As Integer
        Get
            Return _CodigoPedido
        End Get
        Set(ByVal value As Integer)
            _CodigoPedido = value
        End Set
    End Property

    Public ReadOnly Property CodigoClientePedido() As String
        Get
            Return Pedido.Cliente.Codigo
        End Get
    End Property

    Public ReadOnly Property NomeClientePedido() As String
        Get
            Return Pedido.Cliente.Nome & " - " & Pedido.Cliente.Cidade & "/" & Pedido.Cliente.CodigoEstado
        End Get
    End Property


    'Public Property Sequencia() As Integer
    '    Get
    '        Return _Sequencia
    '    End Get
    '    Set(ByVal value As Integer)
    '        _Sequencia = value
    '    End Set
    'End Property

    Public Property Autorizacao() As Integer
        Get
            Return _Autorizacao
        End Get
        Set(ByVal value As Integer)
            _Autorizacao = value
        End Set
    End Property

    Public Property Movimento() As Date
        Get
            Return _Movimento
        End Get
        Set(ByVal value As Date)
            _Movimento = value
        End Set
    End Property

    Public Property DataBaseCalculo() As Date
        Get
            Return _DataBaseCalculo
        End Get
        Set(ByVal value As Date)
            _DataBaseCalculo = value
        End Set
    End Property


    Public Property ClienteRetirada() As Negocio.Cliente
        Get
            If _ClienteRetirada Is Nothing And _CodigoClienteRetirada.Length > 0 Then _ClienteRetirada = New Negocio.Cliente(_CodigoClienteRetirada, _EnderecoClienteRetirada)
            Return _ClienteRetirada
        End Get
        Set(ByVal value As Negocio.Cliente)
            _ClienteRetirada = value
        End Set
    End Property

    Public Property CodigoClienteRetirada() As String
        Get
            Return _CodigoClienteRetirada
        End Get
        Set(ByVal value As String)
            _CodigoClienteRetirada = value
        End Set
    End Property

    Public Property EnderecoClienteRetirada() As Integer
        Get
            Return _EnderecoClienteRetirada
        End Get
        Set(ByVal value As Integer)
            _EnderecoClienteRetirada = value
        End Set
    End Property

    Public ReadOnly Property NomeClienteRetirada() As String
        Get
            Return ClienteRetirada.Nome
        End Get
    End Property


    Public Property Taxa() As Double
        Get
            Return _Taxa
        End Get
        Set(ByVal value As Double)
            _Taxa = value
        End Set
    End Property


    Public Property ClienteAutorizante() As Negocio.Cliente
        Get
            If _ClienteAutorizante Is Nothing And _CodigoAutorizante.Length > 0 Then _ClienteAutorizante = New Negocio.Cliente(_CodigoAutorizante, _EnderecoAutorizante)
            Return _ClienteAutorizante
        End Get
        Set(ByVal value As Negocio.Cliente)
            _ClienteAutorizante = value
        End Set
    End Property

    Public Property CodigoAutorizante() As String
        Get
            Return _CodigoAutorizante
        End Get
        Set(ByVal value As String)
            _CodigoAutorizante = value
        End Set
    End Property

    Public Property EnderecoAutorizante() As Integer
        Get
            Return _EnderecoAutorizante
        End Get
        Set(ByVal value As Integer)
            _EnderecoAutorizante = value
        End Set
    End Property

    Public ReadOnly Property NomeAutorizante() As String
        Get
            Return ClienteAutorizante.Nome
        End Get
    End Property


    Public Property PedidoServico() As Negocio.Pedido
        Get
            If _PedidoServico Is Nothing And _CodigoPedidoServico > 0 Then _PedidoServico = New Negocio.Pedido(_CodigoEmpresa, _EnderecoEmpresa, _CodigoPedidoServico)
            Return _PedidoServico
        End Get
        Set(ByVal value As Negocio.Pedido)
            _PedidoServico = value
        End Set
    End Property

    Public Property CodigoPedidoServico() As Integer
        Get
            Return _CodigoPedidoServico
        End Get
        Set(ByVal value As Integer)
            _CodigoPedidoServico = value
        End Set
    End Property

    Public Property SequenciaPedidoServico() As Integer
        Get
            Return _SequenciaPedidoServico
        End Get
        Set(ByVal value As Integer)
            _SequenciaPedidoServico = value
        End Set
    End Property


    Public Property DataVencimento() As Date
        Get
            Return _DataVencimento
        End Get
        Set(ByVal value As Date)
            _DataVencimento = value
        End Set
    End Property

    Public ReadOnly Property CodigoTitulo() As Integer
        Get
            If _CodigoPedidoServico > 0 Then
                Return PedidoServico.Vencimentos(0).Codigo
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property ValorTitulo() As Double
        Get
            If _QuantidadeAutorizadaFisica > 0 And _Taxa > 0 Then
                '#FimBaseDeCalculo
                'Return (_QuantidadeAutorizadaFisica / Pedido.Itens(0).Produto.BaseCalculo) * _Taxa
                Return _QuantidadeAutorizadaFisica * _Taxa
            Else
                Return 0
            End If
        End Get

    End Property

    Public Property QuantidadeAutorizadaFisica() As Double
        Get
            Return _QuantidadeAutorizadaFisica
        End Get
        Set(ByVal value As Double)
            _QuantidadeAutorizadaFisica = value
        End Set
    End Property

    Public Property QuantidadeEntregueFisica() As Double
        Get
            Return _QuantidadeEntregueFisica
        End Get
        Set(ByVal value As Double)
            _QuantidadeEntregueFisica = value
        End Set
    End Property

    Public Property QuantidadeAutorizadaFiscal() As Double
        Get
            Return _QuantidadeAutorizadaFiscal
        End Get
        Set(ByVal value As Double)
            _QuantidadeAutorizadaFiscal = value
        End Set
    End Property

    Public Property QuantidadeEntregueFiscal() As Double
        Get
            Return _QuantidadeEntregueFiscal
        End Get
        Set(ByVal value As Double)
            _QuantidadeEntregueFiscal = value
        End Set
    End Property

    Public ReadOnly Property SaldoFisico() As Double
        Get
            Return _QuantidadeAutorizadaFisica - _QuantidadeEntregueFisica
        End Get
    End Property

    Public ReadOnly Property SaldoFiscal() As Double
        Get
            Return _QuantidadeAutorizadaFiscal - _QuantidadeEntregueFiscal
        End Get
    End Property

    Public Property Observacao() As String
        Get
            Return _Observacao
        End Get
        Set(ByVal value As String)
            _Observacao = value
        End Set
    End Property

    Public Property UsuarioDeInclusao() As String
        Get
            Return _UsuarioDeInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioDeInclusao = value
        End Set
    End Property

    Public Property UsuarioInclusaoData() As DateTime
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioInclusaoData = value
        End Set
    End Property

    Public Property UsuarioDeAlteracao() As String
        Get
            Return _UsuarioDeAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioDeAlteracao = value
        End Set
    End Property

    Public Property UsuarioAlteracaoData() As DateTime
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioAlteracaoData = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Select Case Me.IUD
            Case "I"
                'Dim n As New Numerador(_CodigoEmpresa, _EnderecoEmpresa, 50)
                '_Autorizacao = n.Sequencia + 1
                'Sqls.Add(n.IncrementarNumeradorSql)

                If (Pedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS Or Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR) And _Taxa > 0 Then
                    Dim nf As New Negocio.NotaFiscal
                    nf.IUD = "I"
                    nf.CodigoEmpresa = _CodigoEmpresa
                    nf.EnderecoEmpresa = _EnderecoEmpresa
                    nf.CodigoPedido = Pedido.Codigo
                    nf.CodigoCliente = Pedido.CodigoCliente
                    nf.EnderecoCliente = Pedido.EnderecoCliente

                    If QuantidadeAutorizadaFisica * Taxa > 5000 Then
                        nf.CodigoOperacao = 70
                        nf.CodigoSubOperacao = 6
                    Else
                        nf.CodigoOperacao = 70
                        nf.CodigoSubOperacao = 7
                    End If
                    nf.DataNota = _Movimento
                    nf.Movimento = _DataVencimento

                    nf.Itens = New ListNotaFiscalXItem(nf)
                    Dim prod As New NotaFiscalXItem(nf)
                    prod.CodigoProduto = 501010029
                    prod.QuantidadeFiscal = 1
                    prod.Unitario = QuantidadeAutorizadaFisica * Taxa
                    prod.ValorTotal = QuantidadeAutorizadaFisica * Taxa
                    prod.CodigoOperacao = nf.CodigoOperacao
                    prod.CodigoSubOperacao = nf.CodigoSubOperacao
                    nf.Itens.Add(prod)
                    nf.AtualizaTotais()
                End If

                Sql = " INSERT INTO AutorizacaoDeRetirada (Empresa_Id, EndEmpresa_id, Pedido_id, Autorizacao_Id, Movimento, ClienteRetirada, EndClienteRetirada, Taxa, " & vbCrLf & _
                      "                                    Autorizante, EndAutorizante, PedidoServico, SequenciaPedidoServico, DataBaseCalculo, DataVencimento,Observacao,UsuarioDeInclusao,UsuarioInclusaoData, QuantidadeFisica, QuantidadeFiscal) " & vbCrLf & _
                      " VALUES ('" & CodigoEmpresa & "'," & _EnderecoEmpresa & "," & _CodigoPedido & "," & _Autorizacao & ",'" & _Movimento.ToString("yyyy-MM-dd") & "','" & _CodigoClienteRetirada & "'," & _EnderecoClienteRetirada & "," & vbCrLf & _
                      "" & Str(_Taxa) & ",'" & _CodigoAutorizante & "'," & _EnderecoAutorizante & ",'" & _CodigoPedidoServico & "'," & _SequenciaPedidoServico & ",'" & _DataBaseCalculo.ToString("yyyy-MM-dd") & "','" & _DataVencimento.ToString("yyyy-MM-dd") & "','" & _Observacao & "','" & _UsuarioDeInclusao & "','" & _UsuarioInclusaoData.ToString("yyyy-MM-dd") & " " & _UsuarioInclusaoData.ToString("HH:mm:ss") & "'," & Str(_QuantidadeAutorizadaFisica) & "," & Str(_QuantidadeAutorizadaFiscal) & ")"
                Sqls.Add(Sql)
            Case "U"

                If (Pedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS Or Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR) And (_Taxa > 0 Or _CodigoPedidoServico > 0) Then
                    Dim nf As New Negocio.NotaFiscal

                    If _Taxa > 0 And _CodigoPedidoServico > 0 Then
                        nf.IUD = "U"
                    ElseIf _Taxa = 0 And _CodigoPedidoServico > 0 Then
                        nf.IUD = "D"
                    ElseIf _Taxa > 0 And _CodigoPedidoServico = 0 Then
                        nf.IUD = "I"
                    End If

                    nf.CodigoAutorizacao = _Autorizacao
                    nf.Autorizacao = Me
                    nf.CodigoEmpresa = _CodigoEmpresa
                    nf.EnderecoEmpresa = _EnderecoEmpresa
                    nf.CodigoPedido = PedidoServico.Codigo
                    nf.CodigoCliente = PedidoServico.CodigoCliente
                    nf.EnderecoCliente = Pedido.EnderecoCliente

                    If QuantidadeAutorizadaFisica * Taxa > 5000 Then
                        nf.CodigoOperacao = 70
                        nf.CodigoSubOperacao = 6
                    Else
                        nf.CodigoOperacao = 70
                        nf.CodigoSubOperacao = 7
                    End If
                    nf.DataNota = _Movimento
                    nf.Movimento = _DataVencimento

                    nf.Itens = New ListNotaFiscalXItem(nf)
                    Dim prod As New NotaFiscalXItem(nf)
                    prod.CodigoProduto = 501010029
                    prod.QuantidadeFiscal = 1
                    prod.Unitario = QuantidadeAutorizadaFisica * Taxa
                    prod.ValorTotal = QuantidadeAutorizadaFisica * Taxa
                    prod.CodigoOperacao = nf.CodigoOperacao
                    prod.CodigoSubOperacao = nf.CodigoSubOperacao
                    nf.Itens.Add(prod)
                    nf.AtualizaTotais()
                End If

                Sql = " UPDATE AutorizacaoDeRetirada SET" & vbCrLf & _
                      "    Movimento              ='" & _Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,ClienteRetirada        ='" & _CodigoClienteRetirada & "'" & vbCrLf & _
                      "   ,EndClienteRetirada     = " & _EnderecoClienteRetirada & vbCrLf & _
                      "   ,Taxa                   = " & Str(_Taxa) & vbCrLf & _
                      "   ,Autorizante            ='" & _CodigoAutorizante & "'" & vbCrLf & _
                      "   ,EndAutorizante         = " & _EnderecoAutorizante & vbCrLf & _
                      "   ,PedidoServico          ='" & _CodigoPedidoServico & "'" & vbCrLf & _
                      "   ,SequenciaPedidoServico = " & IIf(_Taxa = 0, "0", _SequenciaPedidoServico) & vbCrLf & _
                      "   ,DataBaseCalculo        ='" & _DataBaseCalculo.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,DataVencimento         ='" & _DataVencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,Observacao             ='" & _Observacao & "'" & vbCrLf & _
                      "   ,UsuarioDeAlteracao     ='" & _UsuarioDeAlteracao & "'" & vbCrLf & _
                      "   ,UsuarioAlteracaoData   ='" & _UsuarioAlteracaoData.ToString("yyyy-MM-dd") & " " & _UsuarioAlteracaoData.ToString("HH:mm:ss") & "'" & vbCrLf & _
                      "   ,QuantidadeFisica       = " & Str(_QuantidadeAutorizadaFisica) & vbCrLf & _
                      "   ,QuantidadeFiscal       = " & Str(_QuantidadeAutorizadaFiscal) & vbCrLf & _
                      "  WHERE Empresa_Id     ='" & _CodigoEmpresa & "'" & vbCrLf & _
                      "    AND EndEmpresa_Id  = " & _EnderecoEmpresa & vbCrLf & _
                      "    AND Pedido_id      = " & _CodigoPedido & vbCrLf & _
                      "    AND Autorizacao_Id = " & _Autorizacao
                Sqls.Add(Sql)
            Case "D"
                If _QuantidadeEntregueFisica = 0 And _QuantidadeEntregueFiscal = 0 Then
                    Dim nf As New Negocio.NotaFiscal
                    nf.IUD = "D"
                    nf.CodigoAutorizacao = _Autorizacao
                    nf.Autorizacao = Me
                    nf.CodigoEmpresa = _CodigoEmpresa
                    nf.EnderecoEmpresa = _EnderecoEmpresa
                    nf.CodigoPedido = Pedido.Codigo
                    nf.CodigoCliente = Pedido.CodigoCliente
                    nf.EnderecoCliente = Pedido.EnderecoCliente

                    Sql = " DELETE AutorizacaoDeRetirada" & vbCrLf & _
                          "  WHERE Empresa_Id     ='" & _CodigoEmpresa & "'" & vbCrLf & _
                          "    AND EndEmpresa_Id  = " & _EnderecoEmpresa & vbCrLf & _
                          "    AND Pedido_id      = " & _CodigoPedido & vbCrLf & _
                          "    AND Autorizacao_Id = " & _Autorizacao
                    Sqls.Add(Sql)
                End If
        End Select
    End Sub
#End Region

End Class