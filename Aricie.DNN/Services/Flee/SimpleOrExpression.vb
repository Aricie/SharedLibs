﻿Imports Aricie.DNN.UI.Attributes
Imports System.ComponentModel
Imports Aricie.ComponentModel
Imports DotNetNuke.UI.WebControls
Imports Aricie.Services
Imports Aricie.DNN.UI.WebControls
Imports System.Xml.Serialization
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

Namespace Services.Flee
    <Serializable()>
    Public Class SimpleOrExpression(Of T)
        Inherits SimpleOrExpressionBase(Of T)

        Private _expression As  FleeExpressionInfo(Of T)

        Public Sub New()

        End Sub

        Public Sub New(value As T)
            Me.New(value, False)
        End Sub

        <Browsable(False)> _
        <XmlIgnore()> _
        <JsonIgnore()> _
        Public Property TargetSubType As Type


        Public Sub New(value As Object, isExpression As Boolean)
            Me.New(value, isExpression, Nothing)
        End Sub

        Public Sub New(value As Object, isExpression As Boolean, targetSubType As Type)
            If targetSubType IsNot Nothing Then
                Me.TargetSubType = targetSubType
            End If
            If isExpression Then
                Me.Expression = New FleeExpressionInfo(Of T)(DirectCast(value, String))
                Me.Mode = SimpleOrExpressionMode.Expression
            Else
                Me.Simple = DirectCast(value, T)
                Me.Mode = SimpleOrExpressionMode.Simple
            End If
        End Sub


        <ConditionalVisible("Mode", False, True, SimpleOrExpressionMode.Expression)>
        Public Property Expression As FleeExpressionInfo(Of T)
            Get
                If _expression Is Nothing
                    _expression = New FleeExpressionInfo(Of T)()
                End If
                Return _expression
            End Get
            Set
                _expression = value
            End Set
        End Property

        Public  Function ShouldSerializeExpression() As Boolean
            Return Mode = SimpleOrExpressionMode.Expression
        End Function

        Public Overrides Function GetExpression() As SimpleExpression(Of T)
            Return Expression
        End Function

        Public Overrides Function GetSimple() As T
            If TargetSubType IsNot Nothing Then
                Return DirectCast(ReflectionHelper.CreateObject(TargetSubType), T)
            Else
                Return MyBase.GetSimple()
            End If
        End Function
    End Class


    <Serializable()>
    Public MustInherit Class SimpleOrExpressionBase(Of T)
        Inherits SimpleOrExpressionValue(Of T, T)


        Public Function GetValue() As T
            Return GetValue(DnnContext.Current, DnnContext.Current)
        End Function

        Public Shared Function GetValues(sourceCollec As IEnumerable(Of SimpleOrExpressionBase(Of T)), owner As Object, dataContext As IContextLookup) As IEnumerable(Of T)
            Return sourceCollec.Select(Function(objExp) objExp.GetValue(owner, dataContext))
        End Function

        Public Function GetValue(owner As Object, dataContext As IContextLookup) As T
            Select Case Mode
                Case SimpleOrExpressionMode.Simple
                    Return Simple
                Case SimpleOrExpressionMode.Expression
                    Dim toReturn As T = Me.GetExpression.Evaluate(owner, dataContext)
                    If toReturn Is Nothing AndAlso CreateIfNull Then
                        toReturn = Simple
                    End If
                    Return toReturn
            End Select
        End Function

        Public Function GetValue(dataContext As IContextLookup) As T
            Return GetValue(dataContext, dataContext)
        End Function


    End Class




    <Serializable()>
    Public Class SimpleOrExpression(Of TSimple, TExpression)
        Inherits SimpleOrExpressionValue(Of TSimple, TExpression)

        Private _expression As  FleeExpressionInfo(Of TExpression)

        Public Sub New()

        End Sub

        Public Sub New(value As TSimple)
            Me.New(value, False)
        End Sub


        Public Sub New(value As Object, isExpression As Boolean)
            If isExpression Then
                Me.Expression = New FleeExpressionInfo(Of TExpression)(DirectCast(value, String))
                Me.Mode = SimpleOrExpressionMode.Expression
            Else
                Me.Simple = DirectCast(value, TSimple)
                Me.Mode = SimpleOrExpressionMode.Simple
            End If
        End Sub

        <Browsable(False)>
        Public Property Expression As FleeExpressionInfo(Of TExpression)
            Get
                If Mode = SimpleOrExpressionMode.Simple Then
                    Return Nothing
                End If
                If _expression Is Nothing Then
                    _expression = New FleeExpressionInfo(Of TExpression)()
                End If
                Return _expression
            End Get
            Set
                _expression = Value
            End Set
        End Property

        Public Function ShouldSerializeExpression() As Boolean
            Return Mode = SimpleOrExpressionMode.Expression
        End Function

        <XmlIgnore()> _
        <ConditionalVisible("Mode", False, True, SimpleOrExpressionMode.Expression)>
        Public Property EditableExpression As  FleeExpressionInfo(Of TExpression)
            Get
                Return _expression
            End Get
            Set
                _expression = value
            End Set
        End Property

        Public Overrides Function GetExpression() As SimpleExpression(Of TExpression)
            Return Expression
        End Function
    End Class





    <Serializable()>
    Public MustInherit Class SimpleOrExpressionValue(Of TSimple, TExpression)
        Inherits SimpleOrExpressionBase(Of TSimple, TExpression)


        Public Function GetValuePair() As KeyValuePair(Of TExpression, TSimple)
            Return GetValuePair(DnnContext.Current, DnnContext.Current)
        End Function

        Public Shared Function GetValuePairs(sourceCollec As IEnumerable(Of SimpleOrExpressionValue(Of TSimple, TExpression)), owner As Object, dataContext As IContextLookup) As IEnumerable(Of KeyValuePair(Of TExpression, TSimple))
            Return sourceCollec.Select(Function(objExp) objExp.GetValuePair(owner, dataContext))
        End Function

        Public Function GetValuePair(owner As Object, dataContext As IContextLookup) As KeyValuePair(Of TExpression, TSimple)
            Select Case Mode
                Case SimpleOrExpressionMode.Simple
                    Return New KeyValuePair(Of TExpression, TSimple)(Nothing, Simple)
                Case Else
                    Dim exp As TExpression = Me.GetExpression.Evaluate(owner, dataContext)
                    If exp Is Nothing AndAlso CreateIfNull Then
                        Return New KeyValuePair(Of TExpression, TSimple)(Nothing, Simple)
                    End If
                    Return New KeyValuePair(Of TExpression, TSimple)(exp, Nothing)
            End Select
        End Function

    End Class



    <DefaultProperty("FriendlyName")> _
    <Serializable()>
    Public MustInherit Class SimpleOrExpressionBase(Of TSimple, TExpression)

        <XmlIgnore()> _
        <Browsable(False)> _
        Public ReadOnly Property FriendlyName As String
            Get
                If Mode = SimpleOrExpressionMode.Simple Then
                    Return ReflectionHelper.GetFriendlyName(Simple)
                Else
                    Return GetExpression.Expression
                End If
            End Get
        End Property

        '<DefaultValue(DirectCast(SimpleOrExpressionMode.Simple, Object))> _
        <JsonConverter(gettype(StringEnumConverter))> _
        Public Property Mode As SimpleOrExpressionMode

        <DefaultValue(False)> _
        <AutoPostBack()> _
        <ConditionalVisible("Mode", False, True, SimpleOrExpressionMode.Expression)>
        Public Property CreateIfNull As Boolean

        <XmlIgnore()> _
        <Browsable(False)> _
        Public ReadOnly Property DisplaySimple As Boolean
            Get
                Return Mode = SimpleOrExpressionMode.Simple OrElse CreateIfNull
            End Get
        End Property

        Private _Simple As TSimple

        Private Function IsSubType() As Boolean
            If _Simple IsNot Nothing Then
                Return _Simple.GetType() IsNot GetType(TSimple)
            End If
            Return False
        End Function

        <JsonIgnore()> _
        <XmlIgnore()> _
        <SortOrder(100)> _
        <Width(500)> _
        <Required(True)> _
        <ConditionalVisible("DisplaySimple", False, True)>
        Public Property Simple As TSimple
            Get
                If Mode = SimpleOrExpressionMode.Simple AndAlso _Simple Is Nothing Then
                    _Simple = GetSimple()
                End If
                Return _Simple
            End Get
            Set(value As TSimple)
                _Simple = value
            End Set
        End Property
       


        <JsonProperty("Simple")> _
        <XmlElement("Simple")> _
        <Browsable(False)> _
        Public Property SerializedSimple As TSimple
            Get
                If Not IsSubType() Then
                    Dim newSimple = GetSimple()
                    If Mode = SimpleOrExpressionMode.Expression OrElse (newSimple IsNot Nothing AndAlso newSimple.Equals(_Simple)) Then
                        Return Nothing
                    End If
                    Return _Simple
                End If
                Return Nothing
            End Get
            Set(value As TSimple)
                _Simple = value
            End Set
        End Property

        Public Function ShouldSerializeSerializedSimple() As Boolean
            Return Mode = SimpleOrExpressionMode.Simple AndAlso Not IsSubType()
        End Function


        <Browsable(False)> _
        Public Overridable Property Instance As Serializable(Of TSimple)
            Get
                If IsSubType() Then
                    Return New Serializable(Of TSimple)(_Simple)
                End If
                Return Nothing
            End Get
            Set(value As Serializable(Of TSimple))
                Me._Simple = value.Value
            End Set
        End Property


          Public  Function ShouldSerializeInstance() As Boolean
            Return Mode = SimpleOrExpressionMode.Simple AndAlso IsSubType()
        End Function


        <ConditionalVisible("Mode", False, True, SimpleOrExpressionMode.Simple)> _
        <ActionButton(IconName.Refresh, IconOptions.Normal)> _
        Public Sub ResetInstance(ByVal pe As AriciePropertyEditorControl)
            Me.ResetInstance()
            pe.DisplayLocalizedMessage("InstanceReset.Message", DotNetNuke.UI.Skins.Controls.ModuleMessage.ModuleMessageType.GreenSuccess)
            pe.ItemChanged = True
        End Sub

        Public Sub ResetInstance()
            Try
                Me._Simple = Me.GetSimple()
            Catch ex As Exception
                ExceptionHelper.LogException(ex)
            End Try
            
        End Sub

        Public MustOverride Function GetExpression() As SimpleExpression(Of TExpression)

        Public Overridable Function GetSimple() As TSimple
            If GetType(TSimple) IsNot GetType(Object) Then
                Return ReflectionHelper.CreateObject(Of TSimple)()
            End If
            Return Nothing
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Dim typedObj As SimpleOrExpressionBase(Of TSimple, TExpression) = TryCast(obj, SimpleOrExpressionBase(Of TSimple, TExpression))
            If typedObj Is Nothing Then
                Return False
            End If
            Return (Me.Mode = SimpleOrExpressionMode.Simple AndAlso Me.Simple.Equals(typedObj.Simple)) _
                OrElse (Me.Mode = SimpleOrExpressionMode.Expression AndAlso Me.GetExpression().Expression.Equals(typedObj.GetExpression().Expression))

        End Function

    End Class
End Namespace