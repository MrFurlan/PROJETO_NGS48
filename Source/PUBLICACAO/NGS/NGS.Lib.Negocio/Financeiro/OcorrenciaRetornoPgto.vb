Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class OcorrenciaRetornoPgto
    Implements IBaseEntity

    Private _codigo As Integer
    Private _descricao As String
    Private _baixaTitulo As Boolean

    Public Property Codigo() As Integer
        Get
            Return _codigo
        End Get
        Set(ByVal value As Integer)
            _codigo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _descricao
        End Get
        Set(ByVal value As String)
            _descricao = value
        End Set
    End Property

    Public Property BaixaTitulo() As Boolean
        Get
            Return _baixaTitulo
        End Get
        Set(ByVal value As Boolean)
            _baixaTitulo = value
        End Set
    End Property

End Class