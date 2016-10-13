Imports System.Xml
Imports Aricie.DNN.Services
Imports System.Globalization
Imports System.Xml.Serialization
Imports Aricie.DNN.UI.Attributes

Namespace Configuration
    ''' <summary>
    ''' Component to switch application trust level
    ''' </summary>
    
    Public Class UrlCompressionInfo
        Inherits XmlConfigElementInfo


        Public Sub New()

        End Sub

        <AutoPostBack()> _
        Public Property DoStaticCompression As Boolean

        <AutoPostBack()> _
        Public Property DoDynamicCompression As Boolean

        <AutoPostBack()> _
        <ConditionalVisible("DoDynamicCompression")> _
        Public Property DynamicCompressionBeforeCache As Boolean = True


        Public Overrides Function IsInstalled(ByVal xmlConfig As XmlDocument) As Boolean
            Dim xPath As String = NukeHelper.DefineWebServerElementPath("system.webServer") & "/urlCompression[@doStaticCompression='" & DoStaticCompression.ToString(CultureInfo.InvariantCulture).ToLower() & "'"
            If DoDynamicCompression Then
                xPath &= " and @doDynamicCompression='true'"
                If DynamicCompressionBeforeCache Then
                    xPath &= " and @dynamicCompressionBeforeCache='true'"
                Else
                    xPath &= " and (not(@dynamicCompressionBeforeCache) or @dynamicCompressionBeforeCache='false')"
                End If
            Else
                xPath &= " and (not(@doDynamicCompression) or @doDynamicCompression='false')"
            End If
            xPath &= "]"
            Dim moduleNode As XmlNode = xmlConfig.SelectSingleNode(xPath)
            Return (moduleNode IsNot Nothing)
        End Function


        Public Overrides Sub AddConfigNodes(ByRef targetNodes As NodesInfo, ByVal actionType As ConfigActionType)
            Select Case actionType
                Case ConfigActionType.Install

                    Dim node As New StandardComplexNodeInfo((NukeHelper.DefineWebServerElementPath("system.webServer")), NodeInfo.NodeAction.update, StandardComplexNodeInfo.NodeCollision.save, _
                                                            NukeHelper.DefineWebServerElementPath("system.webServer") & "/urlCompression")
                    node.Children.Add(New UrlCompressionAddInfo(Me))
                    targetNodes.Nodes.Add(node)

                Case ConfigActionType.Uninstall
                    Dim node As New NodeInfo((NukeHelper.DefineWebServerElementPath("system.webServer") & "/urlCompression"), NodeInfo.NodeAction.remove)
                    targetNodes.Nodes.Add(node)
            End Select

        End Sub
    End Class

    ''<XmlType("trust")> _
    ''' <summary>
    ''' Custom Error Add merge node
    ''' </summary>
    <XmlRoot("urlCompression")> _
    Public Class UrlCompressionAddInfo
        Inherits AddInfoBase

        Public Sub New()

        End Sub

        Public Sub New(ByVal objConfig As UrlCompressionInfo)
            Me.Attributes("doStaticCompression") = objConfig.DoStaticCompression.ToString(CultureInfo.InvariantCulture).ToLowerInvariant()
            Me.Attributes("doDynamicCompression") = objConfig.DoDynamicCompression.ToString(CultureInfo.InvariantCulture).ToLowerInvariant()

            If objConfig.DynamicCompressionBeforeCache Then
                Me.Attributes("dynamicCompressionBeforeCache") = objConfig.DynamicCompressionBeforeCache.ToString(CultureInfo.InvariantCulture).ToLowerInvariant()
            End If
        End Sub


    End Class
End Namespace