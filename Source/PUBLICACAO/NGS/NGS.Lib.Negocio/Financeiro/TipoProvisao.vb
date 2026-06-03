Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListTipoProvisao
    Inherits List(Of TipoProvisao)

#Region "Construtor"
    Public Sub New(Optional ByVal Carregar As Boolean = False)
        If Carregar Then
            Dim objBanco As New AcessaBanco

            Try
                Dim strSql As String = " SELECT Provisao_Id, Descricao " & _
                                       "   FROM Provisoes "

                Dim dsTipoProvisao As DataSet = objBanco.ConsultaDataSet(strSql, "TipoProvisao")

                For Each rom As DataRow In dsTipoProvisao.Tables(0).Rows
                    Dim tipo As New TipoProvisao
                    tipo.Codigo = rom("Provisao_Id")
                    tipo.Descricao = rom("Descricao")
                    Me.Add(tipo)
                Next

            Catch ex As Exception

            End Try



        End If
    End Sub
#End Region

End Class

<Serializable()> _
Public Class TipoProvisao
    Implements IBaseEntity

#Region "Construtores"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Dim objBanco As New AcessaBanco

        Try
            Dim strSql As String = " SELECT Provisao_Id, Descricao " & _
                                   "   FROM Provisoes " & _
                                   "  WHERE Provisao_Id = " & Codigo

            Dim dsTipoProvisao As DataSet = objBanco.ConsultaDataSet(strSql, "TipoProvisao")

            With dsTipoProvisao.Tables(0).Rows
                If .Count > 0 Then
                    Me.Codigo = .Item(0)("Provisao_Id")
                    Me.Descricao = .Item(0)("Descricao")
                End If
            End With

        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "Fields"
    Private _Codigo As Integer
    Private _Descricao As String
#End Region

#Region "Propriedades"

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

#End Region

End Class
