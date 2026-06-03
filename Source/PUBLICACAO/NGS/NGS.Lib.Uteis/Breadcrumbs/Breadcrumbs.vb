Imports System.Xml.Serialization

<Serializable()> _
Public Class Breadcrumbs

    Private _nodes As List(Of BreadcrumbsNode)
    <XmlElement("BreadcrumbsNode")> _
    Public Property Nodes() As List(Of BreadcrumbsNode)
        Get
            Return _nodes
        End Get
        Set(ByVal value As List(Of BreadcrumbsNode))
            _nodes = value
        End Set
    End Property

End Class