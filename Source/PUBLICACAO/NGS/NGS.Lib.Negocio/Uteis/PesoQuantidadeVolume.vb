Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

Public Class PesoQuantidadeVolume

#Region "Construtor"

    Sub New(ByVal pFatorConversao As Decimal, ByVal pProduto As String, ByVal pQtde As Double, ByVal pPesoQuantidade As String)

        Dim Sql As String

        Sql = " select " & Str(pQtde) & " * PUC.FatorConversao_id AS PesoBruto," & vbCrLf &
              "        " & Str(pQtde) & " * PUC.FatorConversao_id AS PesoLiquido," & vbCrLf &
              "        Round((" & Str(pQtde) & "+ 0.49)/case " & vbCrLf &
              "                       when isnull(QuantidadeNaCaixa,0) = 0" & vbCrLf &
              "                         then 1" & vbCrLf &
              "                         else QuantidadeNaCaixa" & vbCrLf &
              "                      end,0) as Caixas," & vbCrLf &
              "        " & Str(pQtde) & " as Numeracao" & vbCrLf &
              "  FROM Produtos as Prd" & vbCrLf &
              "  INNER JOIN ProdutosxUnidadeDeComercializacao PUC " & vbCrLf &
              "     ON PRO.Produto_Id           = PUC.Produto_id  " & vbCrLf &
              "     AND PRO.Unidade             = PUC.Unidade_id  " & vbCrLf &
              "     AND " & pFatorConversao & " = PUC.FatorConversao_id  " & vbCrLf &
              " WHERE Prd.Produto_Id            = '" & pProduto & "'"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(Sql, "Peso")
        If ds.Tables(0).Rows.Count > 0 Then
            If pPesoQuantidade = "P" Then
                _PesoBruto = pQtde
                _PesoLiquido = pQtde
                _Volumes = 1
                _Numeracao = 1
            Else
                _PesoBruto = ds.Tables(0).Rows(0)("PesoBruto")
                _PesoLiquido = ds.Tables(0).Rows(0)("PesoLiquido")
                _Volumes = ds.Tables(0).Rows(0)("Caixas")
                _Numeracao = ds.Tables(0).Rows(0)("Numeracao")
            End If
        End If
    End Sub
#End Region

#Region "Field"
    Private _PesoBruto As Double = 0
    Private _PesoLiquido As Double = 0
    Private _Volumes As Integer = 0
    Private _Numeracao As Integer = 0
#End Region

#Region "Property"
    Public ReadOnly Property PesoBruto() As Double
        Get
            Return _PesoBruto
        End Get
    End Property

    Public ReadOnly Property PesoLiquido() As Double
        Get
            Return _PesoLiquido
        End Get
    End Property

    Public ReadOnly Property Volumes() As Integer
        Get
            Return _Volumes
        End Get
    End Property

    Public ReadOnly Property Numeracao() As Integer
        Get
            Return _Numeracao
        End Get
    End Property
#End Region

End Class

