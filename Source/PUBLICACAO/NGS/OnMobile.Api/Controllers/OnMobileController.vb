Imports System.Net
Imports System.Threading.Tasks
Imports System.Web.Http
Imports OnMobile.Lib.Negocio

Namespace Controllers
    Public Class OnMobileController
        Inherits ApiController
        Public Shared Async Function Sincronizar() As Task
            Try
                Dim Integracao As New IntegracaoOnMobile()

                Await Integracao.IntegracaoWsVendedor() 'OK

                Await Integracao.IntegracaoWsProduto() 'OK

                Await Integracao.IntegracaoWsTabCondicao() 'OK

                Await Integracao.IntegracaoWsTabPreco() 'OK

                Await Integracao.IntegracaoWsCondicao() 'OK

                Await Integracao.IntegracaoWsTipoOperacao() 'OK

                Await Integracao.IntegracaoWsTabPrecoProduto() 'OK

                Await Integracao.IntegracaoWsCliente()

                Await Integracao.IntegracaoWsLoad()

            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Shared Async Function SincronizarPedido() As Task
            Try
                Dim Integracao As New IntegracaoOnMobile()
                Await Integracao.IntegracaoPedidoOnSolft()
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End Function
    End Class
End Namespace