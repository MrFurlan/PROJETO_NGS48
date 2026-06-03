Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class LaudoPesagem

End Class

<Serializable()> _
Public Class LaudosPesagem
    Inherits List(Of LaudoPesagem)

    Public Shared Function Existe(pEmpresa As String, pEndEmpresa As Integer, ByVal Pedido As Integer, ByVal IncluirCancelados As Boolean) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT 1 " & _
                                   "  FROM Pesagem " & _
                                   " WHERE Empresa_Id    ='" & pEmpresa & "'" & _
                                   "   AND EndEmpresa_Id = " & pEndEmpresa & _
                                   "   AND Pedido        = " & Pedido.ToString() & _
                                   "   AND Sequencia_Id  = 0"

            If Not IncluirCancelados Then strSQL &= "AND Situacao <> 2"

            Dim dsPesagem As DataSet = objBanco.ConsultaDataSet(strSQL, "Pesagem")

            If dsPesagem.Tables(0).Rows.Count > 0 Then Return True Else Return False
        Catch ex As Exception
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

End Class