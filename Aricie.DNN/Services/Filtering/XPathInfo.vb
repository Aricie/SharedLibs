﻿Imports System.Xml.XPath
Imports System.Xml
Imports Aricie.Collections
Imports Aricie.DNN.UI.Attributes
Imports System.ComponentModel
Imports DotNetNuke.UI.WebControls
Imports Aricie.DNN.UI.WebControls.EditControls
Imports Aricie.ComponentModel
Imports Aricie.Services
Imports Aricie.DNN.UI.WebControls
Imports HtmlAgilityPack
Imports System.Xml.Serialization
Imports Aricie.DNN.Services.Flee
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

Namespace Services.Filtering

    
    Public Class HtmlXPathInfo
        Inherits XPathInfo

        <Browsable(False)> _
        Public Overrides Property IsHtmlContent As Boolean
            Get
                Return True
            End Get
            Set(value As Boolean)
                'donothing
            End Set
        End Property

    End Class

    Public Enum XPathMode
        ReturnResults
        UpdateResults
    End Enum

    Public Enum XPathOutputMode
        Selection
        DocumentString
        DocumentNavigable
    End Enum

    Public Enum XPathSelectMode
        SelectionString
        SelectionNodes
    End Enum



    ''' <summary>
    ''' xpath selection helper class
    ''' </summary>
    ''' <remarks></remarks>
    <ActionButton(IconName.Code, IconOptions.Normal)> _
    <DefaultProperty("Expression")> _
    Public Class XPathInfo

        Public Sub New()
        End Sub


        Public Sub New(selectExpression As String, isSingle As Boolean, isHtml As Boolean)
            Me.Expression.Simple = selectExpression
            Me._SingleSelect = isSingle
            Me._IsHtmlContent = isHtml
        End Sub

        ''' <summary>
        ''' XPath selection expression
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Expression As New SimpleOrSimpleExpression(Of CData)("")


        'todo: remove obsolete property
        <Browsable(False)> _
        Public Property SelectExpression() As String
            Get
                Return Nothing
            End Get
            Set(value As String)
                Expression.Simple = value
            End Set
        End Property

        ''' <summary>
        ''' Selection is against Html content using HtmlAgilityPack rather than regular Xml content using System.Xml
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Property IsHtmlContent() As Boolean = True

         <DefaultValue(False)> _
        <ExtendedCategory("XPathSettings")> _
        Public Property UseNamespaceManager As Boolean

         <DefaultValue("")> _
        <ExtendedCategory("XPathSettings")> _
        <ConditionalVisible("UseNamespaceManager")> _
        Public Property DefaultNamespacePrefix As String = ""

         <DefaultValue(False)> _
        <ExtendedCategory("XPathSettings")> _
        Public Property EvaluateExpression As Boolean

        <DefaultValue(DirectCast(XPathOutputMode.Selection, Object))> _
        <JsonConverter(GetType(StringEnumConverter))> _
        <ExtendedCategory("XPathSettings")> _
        <ConditionalVisible("EvaluateExpression", True, True)> _
        Public Property OutputMode As XPathOutputMode = XPathOutputMode.Selection


        <DefaultValue(XPathSelectMode.SelectionString)> _
        <ExtendedCategory("XPathSettings")> _
        <ConditionalVisible("EvaluateExpression", True, True)> _
        <ConditionalVisible("OutputMode", False, True, XPathOutputMode.Selection)> _
        Public Property SelectMode As XPathSelectMode = XPathSelectMode.SelectionString


        ''' <summary>
        ''' Single node selection 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ExtendedCategory("XPathSettings")> _
        <ConditionalVisible("EvaluateExpression", True, True)> _
         <DefaultValue(True)> _
        Public Property SingleSelect() As Boolean = True



        ''' <summary>
        ''' Selection of whole tree
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ConditionalVisible("EvaluateExpression", True, True)> _
        <ExtendedCategory("XPathSettings")> _
        <DefaultValue(False)> _
        Public Property SelectTree() As Boolean

        ''' <summary>
        ''' Sub-selection
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ExtendedCategory("XPathSettings")> _
        <ConditionalVisible("EvaluateExpression", True, True)> _
        <ConditionalVisible("SelectTree", False, True)> _
                <CollectionEditor(DisplayStyle:=CollectionDisplayStyle.Accordion, EnableExport:=True)> _
        Public Property SubSelects() As New SerializableDictionary(Of String, XPathInfo)

        <ExtendedCategory("Filter")> _
        <ConditionalVisible("EvaluateExpression", True, True)> _
        <DefaultValue(False)> _
        Public Property ApplyFilter As Boolean

        <ConditionalVisible("ApplyFilter", False, True)> _
        <ExtendedCategory("Filter")> _
        <ConditionalVisible("EvaluateExpression", True, True)> _
        <DefaultValue(False)> _
        Public Property UpdateNodes As Boolean

        Private _Filter As ExpressionFilterInfo

        <ConditionalVisible("ApplyFilter", False, True)> _
        <ExtendedCategory("Filter")> _
        Public Property Filter As ExpressionFilterInfo
            Get
                If Not Me.ApplyFilter Then
                    Return Nothing
                End If
                If _Filter Is Nothing Then
                    _Filter = New ExpressionFilterInfo
                End If
                Return _Filter
            End Get
            Set(value As ExpressionFilterInfo)
                _Filter = value
            End Set
        End Property


        ''' <summary>
        ''' Simulation data
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ExtendedCategory("Simulation")> _
            <Width(500)> _
            <LineCount(8)> _
            <XmlIgnore()> _
        Public Overridable Property SimulationData() As New CData

        Private _SimulationResult As Object

        <ExtendedCategory("Simulation")> _
        <DefaultValue(False)> _
        Public Property ResultAsXml As Boolean

        ''' <summary>
        ''' Result of the simulation
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ExtendedCategory("Simulation")> _
          <ConditionalVisible("ResultAsXml", True, True)> _
        Public ReadOnly Property BrowsableSimulationResult() As Object
            Get
                Return _SimulationResult
            End Get
        End Property

        <XmlIgnore()> _
        <ExtendedCategory("Simulation")> _
          <ConditionalVisible("ResultAsXml", False, True)> _
        Public ReadOnly Property SimulationResult() As String
            Get
                If _SimulationResult IsNot Nothing Then
                    Return ReflectionHelper.Serialize(_SimulationResult).Beautify().HtmlEncode().HtmlEncode()
                End If
                Return ""
            End Get
        End Property




        <ExtendedCategory("Simulation")> _
        <ActionButton(IconName.Refresh, IconOptions.Normal)> _
        Public Sub RunSimulation(ByVal pe As AriciePropertyEditorControl)
            If Not String.IsNullOrEmpty(Me.SimulationData.Value) Then
                Me._SimulationResult = DoSelect(Me.SimulationData.Value)
                pe.ItemChanged = True
            End If
        End Sub




        ''' <summary>
        ''' Select against a string that will be converted to HTML
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function DoSelect(ByVal source As String) As Object
            Return DoSelect(source, Nothing)
        End Function

        ''' <summary>
        ''' Select against a string that will be converted to HTML
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function DoSelect(ByVal source As String, dataContext As IContextLookup) As Object

            If Not String.IsNullOrEmpty(source) Then
                Dim navigable As IXPathNavigable = Me.GetNavigable(source)
                Return DoSelect(navigable, dataContext)
            End If
            Return Nothing
        End Function

        Public Overloads Function DoSelect(ByVal source As IXPathNavigable) As Object
            Return DoSelect(source, Nothing)
        End Function

        ''' <summary>
        ''' Select against a navigable entity
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function DoSelect(ByVal source As IXPathNavigable, dataContext As IContextLookup) As Object
            If source IsNot Nothing Then
                Dim navigator As XPathNavigator = Nothing
                If UseNamespaceManager Then
                    Select Case source.GetType().Name
                        Case "HtmlDocument"
                            navigator = DirectCast(source, HtmlDocument).DocumentNode.CreateRootNavigator()
                        Case "XmlDocument"
                            navigator = DirectCast(source, XmlDocument).DocumentElement.CreateNavigator()
                    End Select
                Else
                    navigator = source.CreateNavigator()
                End If

                Dim toReturn As Object = Me.SelectNavigate(navigator, dataContext)
                If Not Me.EvaluateExpression Then
                    Select Case OutputMode
                        Case XPathOutputMode.DocumentNavigable
                            toReturn = source
                        Case XPathOutputMode.DocumentString
                            If TypeOf source Is HtmlDocument Then
                                Return DirectCast(source, HtmlDocument).DocumentNode.OuterHtml
                            Else
                                navigator = source.CreateNavigator()
                                navigator.MoveToRoot()
                                Return navigator.OuterXml()
                            End If
                    End Select
                End If
                Return toReturn
            End If
            Return Nothing
        End Function



        Private _CompiledExpression As XPathExpression
        Private _CompiledExpressionSource As String = ""


        Public Overloads Function SelectNavigate(ByVal navigator As XPathNavigator) As Object
            Return SelectNavigate(navigator, Nothing, Nothing)
        End Function


        Public Overloads Function SelectNavigate(ByVal navigator As XPathNavigator, dataContext As IContextLookup) As Object
            Return SelectNavigate(navigator, Nothing, dataContext)
        End Function

        Public Overloads Function SelectNavigate(ByVal navigator As XPathNavigator, objIterator As XPathNodeIterator) As Object
            Return SelectNavigate(navigator, objIterator, Nothing)
        End Function

        ''' <summary>
        ''' Runs selection against a navigator
        ''' </summary>
        ''' <param name="navigator"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function SelectNavigate(ByVal navigator As XPathNavigator, objIterator As XPathNodeIterator, dataContext As IContextLookup) As Object
            Dim selExp As String = Me.Expression.GetValue(dataContext)
            If _CompiledExpression Is Nothing OrElse selExp <> _CompiledExpressionSource Then
                SyncLock Me
                    _CompiledExpressionSource = selExp
                    If UseNamespaceManager Then
                        Dim nsman As New XmlNamespaceManager(navigator.NameTable)
                        For Each nskvp As KeyValuePair(Of String, String) In navigator.GetNamespacesInScope(XmlNamespaceScope.Local)
                            If nskvp.Key.IsNullOrEmpty() Then
                                nsman.AddNamespace(Me.DefaultNamespacePrefix, nskvp.Value)
                            Else
                                nsman.AddNamespace(nskvp.Key, nskvp.Value)
                            End If
                        Next
                        _CompiledExpression = XPathExpression.Compile(selExp, nsman)
                    Else
                        _CompiledExpression = XPathExpression.Compile(selExp)
                    End If

                End SyncLock
            End If
            If Me.EvaluateExpression Then
                Dim toReturn As Object
                If objIterator IsNot Nothing Then
                    toReturn = navigator.Evaluate(_CompiledExpression, objIterator)
                Else
                    toReturn = navigator.Evaluate(_CompiledExpression)
                End If
                If Me.ApplyFilter Then
                    Dim strFilter As String = toReturn.ToString()
                    strFilter = Me.Filter.Process(strFilter, dataContext)
                    toReturn = strFilter
                End If
                Return toReturn
            Else
                Dim results As IEnumerable
                If Me.SingleSelect Then
                    Dim tempResultList As New List(Of XPathNavigator)()
                    Dim singleResult As XPathNavigator = navigator.SelectSingleNode(_CompiledExpression)
                    If singleResult IsNot Nothing Then
                        tempResultList.Add(singleResult)
                    End If
                    results = tempResultList
                Else
                    results = navigator.Select(_CompiledExpression)
                End If
                If Not Me._SelectTree Then
                    Dim multiString As List(Of String) = Nothing
                    Dim resultList As List(Of XPathNavigator) = Nothing
                    If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                        resultList = New List(Of XPathNavigator)
                    Else
                        multiString = New List(Of String)
                    End If
                    Dim multiNodes As New List(Of XPathNavigator)()
                    For Each result As XPathNavigator In results

                        Dim resultValue As String = result.Value
                        If Me.ApplyFilter Then
                            resultValue = Me.Filter.Process(resultValue, dataContext)
                            If Me.UpdateNodes Then
                                If TypeOf result Is HtmlNodeNavigator Then
                                    DirectCast(result, HtmlAgilityPack.HtmlNodeNavigator).CurrentNode.InnerHtml = resultValue
                                Else
                                    result.SetValue(resultValue)
                                End If
                            End If
                        End If
                        If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                            resultList.Add(result)
                        Else
                            multiString.Add(resultValue)
                        End If

                        If SingleSelect Then
                            Exit For
                        End If
                    Next
                    If Not _SingleSelect Then
                        If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                            Return resultList
                        Else
                            Return multiString
                        End If

                    ElseIf (Me.SelectMode = XPathSelectMode.SelectionNodes) Then
                        If resultList.Count > 0 Then
                            Return resultList(0)
                        Else
                            Return Nothing
                        End If

                    ElseIf (Me.SelectMode = XPathSelectMode.SelectionString) Then
                        If multiString.Count > 0 Then
                            Return multiString(0)
                        Else
                            Return String.Empty
                        End If
                    Else
                        Return String.Empty
                    End If
                Else
                    Dim multiDico As List(Of SerializableDictionary(Of String, String)) = Nothing
                    Dim multiDicoNodes As List(Of SerializableDictionary(Of String, Object)) = Nothing
                    If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                        multiDicoNodes = New List(Of SerializableDictionary(Of String, Object))
                    Else
                        multiDico = New List(Of SerializableDictionary(Of String, String))
                    End If
                    For Each result As XPathNavigator In results
                        Dim currentDico As SerializableDictionary(Of String, String) = Nothing
                        Dim currentDicoNodes As SerializableDictionary(Of String, Object) = Nothing
                        If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                            currentDicoNodes = New SerializableDictionary(Of String, Object)
                        Else
                            currentDico = New SerializableDictionary(Of String, String)
                        End If

                        For Each subSelectPair As KeyValuePair(Of String, XPathInfo) In Me._SubSelects
                            Dim objSubResult As Object = subSelectPair.Value.SelectNavigate(result, TryCast(results, XPathNodeIterator), dataContext)
                            If objSubResult IsNot Nothing Then
                                If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                                    currentDicoNodes(subSelectPair.Key) = objSubResult
                                Else
                                    Dim subResult As String = CType(objSubResult, String)
                                    currentDico(subSelectPair.Key) = subResult
                                End If
                            End If
                        Next
                        If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                            multiDicoNodes.Add(currentDicoNodes)
                        Else
                            multiDico.Add(currentDico)
                        End If
                    Next
                    If Not _SingleSelect Then
                        If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                            Return multiDicoNodes
                        Else
                            Return multiDico
                        End If

                    ElseIf multiDico.Count > 0 Then
                        If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                            Return multiDicoNodes(0)
                        Else
                            Return multiDico(0)
                        End If
                    Else
                        If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                            Return New SerializableDictionary(Of String, Object)
                        Else
                            Return New SerializableDictionary(Of String, String)
                        End If
                    End If
                End If
            End If
        End Function





        Public Function GetOutputType() As Type
            If Me.EvaluateExpression Then
                Return GetType(Object)
            Else
                Select Case Me.OutputMode
                    Case XPathOutputMode.DocumentString
                        Return GetType(String)
                    Case XPathOutputMode.DocumentNavigable
                        Return GetType(IXPathNavigable)
                    Case Else
                        If Not Me._SelectTree Then
                            If Not _SingleSelect Then
                                If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                                    Return GetType(List(Of XPathNavigator))
                                Else
                                    Return GetType(List(Of String))
                                End If

                            Else
                                If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                                    Return GetType(XPathNavigator)
                                Else
                                    Return GetType(String)
                                End If
                            End If
                        Else
                            If Not _SingleSelect Then
                                If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                                    Return GetType(List(Of SerializableDictionary(Of String, Object)))
                                Else
                                    Return GetType(List(Of SerializableDictionary(Of String, String)))
                                End If
                            Else
                                If Me.SelectMode = XPathSelectMode.SelectionNodes Then
                                    Return GetType(SerializableDictionary(Of String, Object))
                                Else
                                    Return GetType(SerializableDictionary(Of String, String))
                                End If
                            End If
                        End If
                End Select
            End If
        End Function


        ''' <summary>
        ''' Transforms the parameter string into a navigable xpath object
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetNavigable(ByVal source As String) As IXPathNavigable
            If IsHtmlContent Then
                Dim doc As New HtmlDocument()
                doc.LoadHtml(source)
                Return doc
            Else
                Dim xmlDoc As New XmlDocument()
                xmlDoc.LoadXml(source)
                Return xmlDoc
            End If
        End Function




    End Class

End Namespace


