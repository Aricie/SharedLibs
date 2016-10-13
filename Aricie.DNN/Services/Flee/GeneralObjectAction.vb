Imports System.Reflection
Imports Aricie.DNN.ComponentModel
Imports Aricie.DNN.UI.Attributes
Imports System.ComponentModel
Imports Aricie.DNN.UI.WebControls.EditControls
Imports Aricie.ComponentModel
Imports DotNetNuke.UI.Skins.Controls
Imports DotNetNuke.UI.WebControls
Imports Aricie.Services
Imports Aricie.DNN.UI.WebControls
Imports DotNetNuke.Services.Localization
Imports System.Xml.Serialization
Imports Fasterflect

Namespace Services.Flee


    ''' <summary>
    ''' Non Generic version of the ObjectAction
    ''' </summary>
    ''' <remarks></remarks>
    <DefaultProperty("FriendlyName")> _
    Public Class GeneralObjectAction
        Inherits ObjectAction
        Implements ISelector(Of EventInfo)
        Implements ISelector(Of MethodInfo)
        Implements ISelector(Of PropertyInfo)

        <Browsable(False)> _
        Public ReadOnly Property FriendlyName As String
            Get
                Dim strTypeName As String = "UnTyped"
                If Me.HasType Then
                    strTypeName = ReflectionHelper.GetSimpleTypeName(DotNetType.GetDotNetType())
                End If
                Return String.Format("{1}{0}{2}{0}{3}", UIConstants.TITLE_SEPERATOR, strTypeName, Me.ActionMode.ToString(), MemberName.ToString())
            End Get
        End Property

        Public Property DotNetType As New DotNetType()

        <Browsable(False)> _
        Public ReadOnly Property HasType As Boolean
            Get
                Return DotNetType.GetDotNetType() IsNot Nothing
            End Get
        End Property

        <Browsable(False)> _
        Public ReadOnly Property HasConcreteType As Boolean
            Get
                Return HasType AndAlso Not ReflectionHelper.IsStatic(Me.DotNetType.GetDotNetType())
            End Get
        End Property

        Private _PreviousType As String

        <Browsable(False)> _
        Public Overrides ReadOnly Property ObjectType As String
            Get
               
                Return DotNetType.TypeName
            End Get
        End Property

        
        '<ExtendedCategory("Action")> _
        <ConditionalVisible("HasType", False, True)> _
        Public Property ActionMode As ObjectActionMode


        '<ExtendedCategory("Action")> _
        ''' <summary>
        ''' Gets or sets the property name
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ConditionalVisible("HasType", False, True)> _
        <Editor(GetType(SelectorEditControl), GetType(EditControl))> _
        <Selector("Name", "Name", False, True, "<Select a Member Name>", "", False, True)> _
        <AutoPostBack()> _
        Public Property MemberName() As String = ""
           

        'Private _UpdatePropertyValueSubtype As Boolean = False

        '<ExtendedCategory("Action")> _
        <Editor(GetType(SelectorEditControl), GetType(EditControl))> _
       <ConditionalVisible("HasType", False, True)> _
        <ConditionalVisible("MemberName", True, True, "")> _
       <ProvidersSelector("Key", "Value")> _
        Public Property MemberIndex As Integer = 1

        <XmlIgnore()> _
        <Browsable(False)> _
        Public ReadOnly Property HasParameters As Boolean
            Get
                Dim toReturn As Boolean = False
                Dim objSelectedMember As MemberInfo = Me.SelectedMember
                If objSelectedMember IsNot Nothing Then
                    'If _UpdatePropertyValueSubtype Then
                    '    _UpdatePropertyValueSubtype = False
                    '    Dim objTargetType As Type = DirectCast(objSelectedMember, PropertyInfo).PropertyType
                    '    If Me._Value.TargetSubType IsNot objTargetType Then
                    '        Me._Value.TargetSubType = objTargetType
                    '    End If
                    'End If
                    If TypeOf objSelectedMember Is PropertyInfo Then
                        toReturn = DirectCast(objSelectedMember, PropertyInfo).GetIndexParameters().Length > 0
                    ElseIf TypeOf objSelectedMember Is MethodInfo Then
                        toReturn = DirectCast(objSelectedMember, MethodInfo).GetParameters().Length > 0
                    End If
                    If (Not toReturn) AndAlso Me.Parameters.Instances.Count > 0 Then
                        Me.Parameters.Instances.Clear()
                    End If
                End If
                Return toReturn
            End Get
        End Property

        '<ExtendedCategory("Instance")> _
        Public Property StaticCall As Boolean

        <ConditionalVisible("HasConcreteType")> _
        <ConditionalVisible("StaticCall", True)> _
        Public Overrides Property LockTarget As Boolean
            Get
                Return MyBase.LockTarget
            End Get
            Set(value As Boolean)
                MyBase.LockTarget = value
            End Set
        End Property


        '<ExtendedCategory("Instance")> _
        ''' <summary>
        ''' Instance of the generic type
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ConditionalVisible("HasConcreteType", False, True)> _
        <ConditionalVisible("StaticCall", True, True)> _
        Public Property Instance() As New FleeExpressionInfo(Of Object)

        '<ExtendedCategory("Action")> _
        ''' <summary>
        ''' Parameters for the object
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ConditionalVisible("HasType", False, True)> _
        <ConditionalVisible("HasParameters", False, True)> _
        Public Overridable Property Parameters() As New Variables

        '<ExtendedCategory("Action")>
        <ConditionalVisible("HasType", False, True)> _
        <ConditionalVisible("ActionMode", False, True, ObjectActionMode.SetProperty)> _
        <LabelMode(LabelMode.Top)>
        Public Property Value As SimpleOrExpression(Of Object)
            Get
                If Me.ActionMode = ObjectActionMode.SetProperty AndAlso _Value IsNot Nothing AndAlso _Value.Mode = SimpleOrExpressionMode.Simple _
                    AndAlso (_Value.Simple Is Nothing) Then
                    Me.InitializeSimpleValue()
                End If
                Return _Value
            End Get
            Set(value As SimpleOrExpression(Of Object))
                _Value = value
            End Set
        End Property

        '<Browsable(False)> _
        'Public ReadOnly Property ValueIsSimple As Boolean
        '    Get
        '        Return Value.Mode = SimpleOrExpressionMode.Simple
        '    End Get
        'End Property

        '<ConditionalVisible("ValueIsSimple")> _
        <ConditionalVisible("HasType", False, True)> _
       <ConditionalVisible("ActionMode", False, True, ObjectActionMode.SetProperty)> _
        <ActionButton(IconName.Magic, IconOptions.Normal)> _
        Public Sub InitializeSimpleValue(ByVal pe As AriciePropertyEditorControl)
            If Me.InitializeSimpleValue() Then
                pe.DisplayLocalizedMessage("ValueInitialized.Message", DotNetNuke.UI.Skins.Controls.ModuleMessage.ModuleMessageType.GreenSuccess)
                pe.ItemChanged = True
            End If
        End Sub

        Public Function InitializeSimpleValue() As Boolean
            Dim objSelectedMember As MemberInfo = Me.SelectedMember
            If objSelectedMember IsNot Nothing Then
                Dim objTargetType As Type = DirectCast(objSelectedMember, PropertyInfo).PropertyType
                If Me._Value.Simple IsNot Nothing AndAlso Me._Value.TargetSubType IsNot Nothing AndAlso Not Me._Value.TargetSubType.InheritsOrImplements(objTargetType) Then
                    Me._Value.TargetSubType = objTargetType
                    Me._Value.ResetInstance()
                    Me._Value.Mode = SimpleOrExpressionMode.Simple
                    Return True
                End If
            End If
            Return False
        End Function

        '<ExtendedCategory("Action")> _
        <ConditionalVisible("HasType", False, True)> _
        <ConditionalVisible("ActionMode", False, True, ObjectActionMode.AddEventHandler)> _
        Public Overridable Property DelegateExpression As New FleeExpressionInfo(Of [Delegate])

        Private _CandidateEvent As EventInfo
        Private _Value As New SimpleOrExpression(Of Object)

        <Browsable(False)> _
        Protected ReadOnly Property CandidateEvent As EventInfo
            Get
                If _CandidateEvent Is Nothing Then
                    If Me.ActionMode = ObjectActionMode.AddEventHandler AndAlso Me.SelectedMembers.Count > 0 Then
                        Dim candidateEventMember As MemberInfo = Me.SelectedMembers(0)
                        If candidateEventMember IsNot Nothing Then
                            Dim objCandidateEvent As EventInfo = TryCast(candidateEventMember, EventInfo)
                            If objCandidateEvent IsNot Nothing Then
                                _CandidateEvent = objCandidateEvent
                            Else
                                Throw New Exception(String.Format( _
                                    "Candidate Member {0} could not be converted to event {1} in type {2}", _
                                    candidateEventMember.ToString(), Me.MemberName, Me.DotNetType.GetDotNetType()))
                            End If
                        Else
                            Throw New Exception(String.Format("Event {0} was not found in type {1}", Me.MemberName, ReflectionHelper.GetSafeTypeName(Me.DotNetType.GetDotNetType())))
                        End If
                    End If
                End If
                Return _CandidateEvent
            End Get
        End Property



        '<ExtendedCategory("Action")> _
        <ConditionalVisible("HasParameters", False, True)> _
        <ActionButton(IconName.Key, IconOptions.Normal)> _
        Public Sub SetParameters(ape As AriciePropertyEditorControl)
            If Me.SelectedMember IsNot Nothing Then
                Dim objParameters As ParameterInfo() = {}
                If TypeOf SelectedMember Is PropertyInfo Then
                    objParameters = DirectCast(SelectedMember, PropertyInfo).GetIndexParameters()
                ElseIf TypeOf SelectedMember Is MethodInfo Then
                    objParameters = DirectCast(SelectedMember, MethodInfo).GetParameters()
                End If
                Me.Parameters = Variables.GetFromParameters(objParameters)
                ape.ItemChanged = True
                Dim message As String = Localization.GetString("ParametersCreated.Message", ape.LocalResourceFile)
                ape.DisplayMessage(message, ModuleMessage.ModuleMessageType.GreenSuccess)
            End If
        End Sub



        Private _SelectedMember As MemberInfo


        Public Overrides Function Run(owner As Object, globalVars As IContextLookup) As Object
            Dim toReturn As Object = Nothing
            If Me.Enabled Then
                If Not Me.HasType Then
                    Throw New ApplicationException("GeneralObjectAction has no Type Defined")
                End If
                If _SelectedMember Is Nothing Then
                    _SelectedMember = SelectedMember
                End If
                Dim args As New List(Of Object)

                For Each objParam As KeyValuePair(Of String, Object) In Me.Parameters.EvaluateVariables(owner, globalVars)
                    args.Add(objParam.Value)
                Next
                Select Case Me.ActionMode
                    Case Flee.ObjectActionMode.SetProperty
                        Dim targetProp As PropertyInfo
                        targetProp = DirectCast(_SelectedMember, PropertyInfo)
                        If targetProp.GetIndexParameters.Length = args.Count Then
                            Dim objValue As Object = Me.Value.GetValue(owner, globalVars)
                            If Not (targetProp.GetGetMethod().IsStatic OrElse Me.StaticCall) Then
                                Dim target As Object = Me.Instance.Evaluate(owner, globalVars, DotNetType.GetDotNetType())
                                If target Is Nothing Then
                                    Throw New ApplicationException(String.Format("Expression {0} returns nothing", Me.Instance.Expression))
                                End If
                                If Me.LockTarget Then
                                    SyncLock target
                                        targetProp.SetValue(target, objValue, args.ToArray())
                                    End SyncLock
                                Else
                                    targetProp.SetValue(target, objValue, args.ToArray())
                                End If
                            Else
                                targetProp.SetValue(Nothing, objValue, args.ToArray())
                            End If
                            Return toReturn
                        Else
                            Throw New Exception(String.Format("Property {0} with {1} parameters was not found in type {2}", _
                                                         Me.MemberName, args.Count, ReflectionHelper.GetSafeTypeName(Me.DotNetType.GetDotNetType())))
                        End If
                       
                    Case Flee.ObjectActionMode.CallMethod
                       
                        Dim targetMethod As MethodInfo
                        targetMethod = DirectCast(_SelectedMember, MethodInfo)
                        If targetMethod.GetParameters.Length = args.Count Then
                            If targetMethod.IsStatic OrElse Me.StaticCall Then
                                toReturn = targetMethod.Invoke(Nothing, args.ToArray)
                            Else
                                Dim target As Object = Me.Instance.Evaluate(owner, globalVars, DotNetType.GetDotNetType())
                                If target Is Nothing Then
                                    Throw New ApplicationException(String.Format("Expression {0} returns nothing", Me.Instance.Expression))
                                Else
                                    If Me.LockTarget Then
                                        SyncLock target
                                            toReturn = targetMethod.Invoke(target, args.ToArray)
                                        End SyncLock
                                    Else
                                        toReturn = targetMethod.Invoke(target, args.ToArray)
                                    End If
                                End If
                            End If
                        Else
                            Throw New Exception(String.Format("Method {0} with {1} parameters was not found in type {2}", _
                                                             Me.MemberName, args.Count, ReflectionHelper.GetSafeTypeName(Me.DotNetType.GetDotNetType())))
                        End If

                    Case Flee.ObjectActionMode.AddEventHandler
                        Dim target As Object = Me.Instance.Evaluate(owner, globalVars, DotNetType.GetDotNetType())
                        If target IsNot Nothing Then
                            AddEventHandler(owner, globalVars, target)
                        Else
                            Throw New Exception(String.Format( _
                                "Instance Expression {0} returned a null instance while adding event handler {1}  in type {2}", _
                                Me.Instance.Expression, CandidateEvent.ToString(), Me.DotNetType.GetDotNetType()))
                        End If
                End Select
            End If
            Return toReturn
        End Function


        Protected Overridable Sub AddEventHandler(owner As Object, globalVars As IContextLookup, target As Object)
            Dim objDelegate As [Delegate] = Me.DelegateExpression.Evaluate(owner, globalVars)
            If objDelegate IsNot Nothing Then
                If Me.LockTarget Then
                    SyncLock target
                        CandidateEvent.AddEventHandler(target, objDelegate)
                    End SyncLock
                Else
                    CandidateEvent.AddEventHandler(target, objDelegate)
                End If
            Else
                Throw New Exception(String.Format( _
                    "Delegate Expression {0} returned a null instance while adding event handler {1}  in type {2}", _
                    Me.DelegateExpression.Expression, CandidateEvent.ToString(), Me.DotNetType.GetDotNetType()))
            End If
        End Sub

        <XmlIgnore()> _
        <Browsable(False)> _
        Public ReadOnly Property SelectedMembers As List(Of MemberInfo)
            Get

                Dim toReturn As List(Of MemberInfo) = Nothing
                If Me.DotNetType.GetDotNetType() Is Nothing OrElse Not ReflectionHelper.GetFullMembersDictionary(Me.DotNetType.GetDotNetType()).TryGetValue(Me.MemberName, toReturn) Then
                    toReturn = New List(Of MemberInfo)
                End If
                Return toReturn
            End Get
        End Property

        <XmlIgnore()> _
        <Browsable(False)> _
        Public ReadOnly Property SelectedMember As MemberInfo
            Get
                Dim tmpMembers As List(Of MemberInfo) = SelectedMembers
                If MemberIndex = 0 Then
                    MemberIndex = 1
                End If
                If tmpMembers.Count >= MemberIndex Then
                    Return tmpMembers(MemberIndex - 1)
                End If
                Return Nothing
            End Get
        End Property

        Public Function GetSelector(propertyName As String) As IList Implements ISelector.GetSelector
            Select Case propertyName
                Case "MemberName"
                    Select Case Me.ActionMode
                        Case ObjectActionMode.SetProperty
                            Return DirectCast(GetSelectorProperties(propertyName), IList)
                        Case ObjectActionMode.CallMethod
                            Return DirectCast(GetSelectorMethods(propertyName), IList)
                        Case ObjectActionMode.AddEventHandler
                            Return DirectCast(GetSelectorEvents(propertyName), IList)

                    End Select
                Case "MemberIndex"
                    Dim toReturn As New Dictionary(Of String, Integer)
                    If Not Me.MemberName.IsNullOrEmpty() Then
                        Dim index As Integer = 1
                        For Each objMember As MemberInfo In Me.SelectedMembers
                            toReturn(ReflectionHelper.GetMemberSignature(objMember)) = index
                            index += 1
                        Next
                    End If
                    Return toReturn.ToList()
            End Select
            Return Nothing
        End Function

        Public Function GetSelectorProperties(propertyName As String) As IList(Of PropertyInfo) Implements ISelector(Of PropertyInfo).GetSelectorG
            Dim toReturn As New List(Of PropertyInfo)()
            If Me.HasType Then
                toReturn.AddRange(From objProperty In ReflectionHelper.GetPropertiesDictionary(DotNetType.GetDotNetType()).Values Where objProperty.CanWrite)
                toReturn.Sort(Function(objProp1, objProp2) String.Compare(objProp1.Name, objProp2.Name, StringComparison.InvariantCultureIgnoreCase))
            End If
            Return toReturn
        End Function

        ''' <summary>
        ''' Returns a list of the methods on the generic type
        ''' </summary>
        ''' <param name="propertyName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSelectorMethods(ByVal propertyName As String) As System.Collections.Generic.IList(Of System.Reflection.MethodInfo) Implements ComponentModel.ISelector(Of System.Reflection.MethodInfo).GetSelectorG
            Dim toReturn As New List(Of MethodInfo)
            If Me.HasType Then
                toReturn.AddRange(ReflectionHelper.GetMembersDictionary(Me.DotNetType.GetDotNetType(), True, False).Values.OfType(Of MethodInfo)())
                toReturn.Sort(Function(objMember1, objMember2) String.Compare(objMember1.Name, objMember2.Name, StringComparison.InvariantCultureIgnoreCase))
            End If
            Return toReturn
        End Function



        Public Function GetSelectorEvents(propertyName As String) As IList(Of EventInfo) Implements ISelector(Of EventInfo).GetSelectorG
            Dim toReturn As New List(Of EventInfo)
            If Me.HasType Then
                toReturn = ReflectionHelper.GetMembersDictionary(Me.DotNetType.GetDotNetType(), True, False).Values.OfType(Of EventInfo)().ToList()
                toReturn.Sort(Function(objMember1, objMember2) String.Compare(objMember1.Name, objMember2.Name, StringComparison.InvariantCultureIgnoreCase))
            End If
            Return toReturn
        End Function


        Public Overrides Function GetOutputType() As Type
            Select Case Me.ActionMode
                Case Flee.ObjectActionMode.CallMethod
                    For Each potentialMember As MemberInfo In SelectedMembers
                        If TypeOf potentialMember Is MethodInfo Then
                            Dim targetMethod As MethodInfo = DirectCast(potentialMember, MethodInfo)
                            Return targetMethod.ReturnType
                        End If
                    Next
            End Select
            Return Nothing
        End Function
    End Class
End Namespace