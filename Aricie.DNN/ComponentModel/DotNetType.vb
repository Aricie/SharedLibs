﻿Imports System.ComponentModel
Imports Aricie.DNN.UI.Attributes
Imports DotNetNuke.UI.WebControls
Imports Aricie.Services
Imports Aricie.DNN.UI.WebControls.EditControls
Imports System.Xml.Serialization
Imports Aricie.DNN.UI.WebControls
Imports System.Reflection
Imports System.IO
Imports System.Text
Imports Aricie.DNN.Entities
Imports Aricie.DNN.Services.Flee
Imports Aricie.DNN.Services
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

Namespace ComponentModel
    Public Enum TypeSelector
        CommonTypes
        BrowseHierarchy
        NewType
    End Enum

    <ActionButton(IconName.PuzzlePiece, IconOptions.Normal)> _
    <DefaultProperty("Name")> _
    Public Class DotNetType
        Implements ISelector
        Implements ISelector(Of DotNetType)
        Implements ISelector(Of AssemblyName)
        Implements ISelector(Of String)





        Private _TypeName As String = ""
        Private _Type As Type
        Private _TypeNameSelect As String = ""


        Public Sub New()
        End Sub


        Private Shared _CommonTypes As New Dictionary(Of String, DotNetType)


        Private Function AddCommonType(ByVal strType As String, objDotNetType As DotNetType) As String
            Dim tmpType As Type = Nothing
            Dim tmpDNT As DotNetType = Nothing
            If Not _CommonTypes.TryGetValue(strType, tmpDNT) Then
                If Not String.IsNullOrEmpty(strType) Then
                    If objDotNetType Is Nothing Then
                        tmpType = ReflectionHelper.CreateType(strType, False)
                        If tmpType IsNot Nothing Then
                            strType = ReflectionHelper.GetSafeTypeName(tmpType)
                            tmpDNT = New DotNetType(tmpType)
                        Else
                            tmpDNT = New DotNetType(strType)
                        End If
                    Else
                        tmpDNT = objDotNetType
                    End If

                Else
                    tmpDNT = New DotNetType()
                End If
                SyncLock _CommonTypes
                    _CommonTypes(strType) = tmpDNT
                End SyncLock
            End If
            Return strType
        End Function

        Private Function AddCommonType(ByVal strType As String) As String
            Return AddCommonType(strType, Nothing)
        End Function


        Public Sub New(ByVal typeName As String)
            Me.SetTypeName(typeName)
        End Sub

        Public Sub New(ByVal objType As Type)
            If objType IsNot Nothing AndAlso Not objType.IsGenericParameter Then
                SetType(objType)
            End If
        End Sub

        Public Sub SetType(ByVal objType As Type)
              If objType IsNot Nothing
                Me.SetTypeName(ReflectionHelper.GetSafeTypeName(objType))
              End If
        End Sub

        <XmlIgnore()> _
        <Browsable(False)> _
        Public Overridable ReadOnly Property Name() As String
            Get
                Dim objType As Type = Me.GetDotNetType()

                Return ReflectionHelper.GetSimpleTypeName(objType)
            End Get
        End Property


         <DefaultValue(DirectCast(TypePickerMode.SelectType, Object))> _
         <JsonConverter(gettype(StringEnumConverter))> _
        Public Overridable Property PickerMode As TypePickerMode


        <ConditionalVisible("PickerMode", False, True, TypePickerMode.CustomObject)> _
        Public Property CustomTypeIdentifier As  EnabledFeature(Of NamedIdentifierEntity)
            Get
                If PickerMode <>  TypePickerMode.CustomObject Then
                    Return Nothing
                End If
                if _customTypeIdentifier Is nothing
                    _customTypeIdentifier = new EnabledFeature(Of NamedIdentifierEntity)
                End If
                Return _customTypeIdentifier
            End Get
            Set
                _customTypeIdentifier = value
            End Set
        End Property

        <ConditionalVisible("PickerMode", False, True, TypePickerMode.CustomObject)>
        Public Property CustomObject As  Variables
            Get
                If PickerMode <>  TypePickerMode.CustomObject Then
                    Return Nothing
                End If
                if _CustomObject Is nothing
                    _CustomObject = new Variables()
                End If
                Return _CustomObject
            End Get
            Set
                _CustomObject = value
            End Set
        End Property

        <ConditionalVisible("PickerMode", False, True, TypePickerMode.CustomObject)> _
        <ActionButton(IconName.Refresh, IconOptions.Normal, "CustomTypeReset.Alert")> _
        Public Sub ResetCustomType(ByVal pe As AriciePropertyEditorControl)
            Try
                _Type = Nothing
                pe.DisplayLocalizedMessage("TypeReset.Message", DotNetNuke.UI.Skins.Controls.ModuleMessage.ModuleMessageType.GreenSuccess)

                pe.ItemChanged = True
            Catch ex As Exception
                Dim newEx As New ApplicationException("There was an error trying to create your type. See the complete Stack for more details", ex)
                Throw newEx
            End Try
        End Sub



        Private _TypeSelector As Nullable(Of TypeSelector)

        <ConditionalVisible("PickerMode", False, True, TypePickerMode.SelectType)> _
        <EditOnly()> _
        <XmlIgnore()> _
        Public Property TypeSelector As TypeSelector
            Get
                If Not _TypeSelector.HasValue Then
                    Dim targetType As DotNetType = Nothing
                    If Not String.IsNullOrEmpty(_TypeName) AndAlso (Not _CommonTypes.TryGetValue(_TypeName, targetType) OrElse targetType Is Nothing) Then
                        Return TypeSelector.BrowseHierarchy
                    Else
                        Return TypeSelector.CommonTypes
                    End If
                End If
                Return _TypeSelector.Value
            End Get
            Set(value As TypeSelector)
                _TypeSelector = value
            End Set
        End Property

        <ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        <EditOnly()> _
        <Editor(GetType(SelectorEditControl), GetType(EditControl))> _
        <ConditionalVisible("TypeSelector", False, True, TypeSelector.CommonTypes)> _
        <Selector("Name", "TypeName", False, True, "<Select Type By Name>", "", False, True)> _
        <AutoPostBack> _
        <XmlIgnore()> _
        Public Property CommonType() As String
            Get
                Return _TypeName
            End Get
            Set(value As String)
                If Me._TypeName <> value Then
                    Me.SetTypeName(value)
                End If
            End Set
        End Property



        '<ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        '<EditOnly()> _
        '<ConditionalVisible("TypeSelector", False, True, TypeSelector.CommonTypes)> _
        '<IsReadOnly(True)> _
        <Browsable(False)> _
        Public Property TypeName() As String
            Get
                Return _TypeName
            End Get
            Set(ByVal value As String)
                If _TypeName <> value Then
                    Me.SetTypeName(AddCommonType(value))
                End If
            End Set
        End Property

        <DefaultValue(False)> _
        <ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        <EditOnly()> _
        <ConditionalVisible("TypeSelector", False, True, TypeSelector.BrowseHierarchy)> _
        <XmlIgnore> _
        Public Property MakeArray As Boolean

        <DefaultValue(1)> _
        <ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        <AutoPostBack()> _
        <EditOnly()> _
        <ConditionalVisible("TypeSelector", False, True, TypeSelector.BrowseHierarchy)> _
        <ConditionalVisible("MakeArray")> _
        <XmlIgnore> _
        Public Property Rank As Integer = 1

        Private _ArrayType As DotNetType
        Private _customTypeIdentifier As EnabledFeature(Of NamedIdentifierEntity)
        Private _CustomObject As Variables

        <ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        <EditOnly()> _
        <ConditionalVisible("TypeSelector", False, True, TypeSelector.BrowseHierarchy)> _
        <ConditionalVisible("MakeArray")> _
        <XmlIgnore()> _
        Public Property ArrayType As DotNetType
            Get
                If Not MakeArray Then
                    Return Nothing
                End If
                If _ArrayType Is Nothing Then
                    _ArrayType = New DotNetType()
                End If
                Return _ArrayType
            End Get
            Set(value As DotNetType)
                _ArrayType = value
            End Set
        End Property

        <DefaultValue("")> _
        <ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        <ConditionalVisible("MakeArray", True)> _
        <EditOnly()> _
        <ConditionalVisible("TypeSelector", False, True, TypeSelector.BrowseHierarchy)> _
        <Selector("Name", "FullName", False, True, "<Select an Assembly>", "", False, True)> _
        <Editor(GetType(SelectorEditControl), GetType(EditControl))> _
        <XmlIgnore()> _
        Public Property AssemblyNameSelect As String = ""

        <DefaultValue("")> _
        <ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        <ConditionalVisible("MakeArray", True)> _
        <EditOnly()> _
        <ConditionalVisible("IsGenericParameter", True, True)> _
        <ConditionalVisible("AssemblyNameSelect", True, True, "")> _
        <ConditionalVisible("TypeSelector", False, True, TypeSelector.BrowseHierarchy)> _
        <SelectorAttribute("", "", False, True, "<Select Namespace>", "", False, True)> _
        <AutoPostBack()> _
        <Editor(GetType(SelectorEditControl), GetType(EditControl))> _
        <XmlIgnore()> _
        Public Property NamespaceSelect As String = ""

        <DefaultValue("")> _
        <ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        <ConditionalVisible("MakeArray", True)> _
        <EditOnly()> _
        <AutoPostBack()> _
        <ConditionalVisible("AssemblyNameSelect", True, True, "")> _
        <ConditionalVisible("TypeSelector", False, True, TypeSelector.BrowseHierarchy)> _
        <Editor(GetType(SelectorEditControl), GetType(EditControl))> _
        <Selector("Name", "TypeName", False, True, "<No Typename Selected>", "", False, True)> _
        <XmlIgnore()> _
        Public Property TypeNameSelect As String
            Get
                Return _TypeNameSelect
            End Get
            Set(value As String)
                If _TypeNameSelect <> value Then
                    Me._TypeNameSelect = value
                    Me.SetTypeName(value)
                End If
            End Set
        End Property

        <XmlIgnore()> _
        <Browsable(False)> _
        Public ReadOnly Property IsSelectedGeneric As Boolean
            Get
                'Return Me.TypeNameSelect.Contains("[")
                Return CurrentlySelectedType IsNot Nothing AndAlso CurrentlySelectedType.IsGenericType
            End Get
        End Property

        ' <ConditionalVisible("IsGenericParameter", True, True)> _
        ' <XmlIgnore()> _
        '<ConditionalVisible("TypeSelector", False, True, TypeSelector.BrowseHierarchy)> _
        '<ConditionalVisible("IsSelectedGeneric", False, True, TypeSelector.BrowseHierarchy)> _
        ' Public Property IncludeGenericTypes As Boolean
        <ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        <EditOnly()> _
        <ConditionalVisible("IsSelectedGeneric", False, True)> _
        <ConditionalVisible("AssemblyNameSelect", True, True, "")> _
        <XmlIgnore()> _
        <ConditionalVisible("TypeSelector", False, True, TypeSelector.BrowseHierarchy)> _
        <CollectionEditor(DisplayStyle:=CollectionDisplayStyle.List, EnableExport:=False, _
          ItemsReadOnly:=False, MaxItemNb:=-1, NoAdd:=True, NoDeletion:=True, Ordered:=False, Paged:=False, ShowAddItem:=False)> _
        Public Property GenericTypes As New List(Of DotNetType)

        <XmlIgnore()> _
        <Browsable(False)> _
        Public ReadOnly Property CurrentlySelectedType As Type
            Get
                Dim toReturn As Type = Nothing

                If Me.MakeArray Then
                    toReturn = Me.ArrayType.GetDotNetType()
                    If toReturn IsNot Nothing Then
                        If Rank > 1 Then
                            toReturn = toReturn.MakeArrayType(Rank)
                        Else
                            toReturn = toReturn.MakeArrayType()
                        End If
                    End If
                Else
                    If Not String.IsNullOrEmpty(TypeNameSelect) Then
                        toReturn = Me.GetDotNetType(TypeNameSelect, False)
                        If toReturn IsNot Nothing AndAlso toReturn.IsGenericTypeDefinition Then
                            Dim genParams As Type() = toReturn.GetGenericArguments()
                            Dim passedParams As New List(Of Type)
                            For i As Integer = 0 To genParams.Count - 1
                                Dim genericParam As Type = Nothing
                                If Me.GenericTypes.Count > i Then
                                    genericParam = Me.GenericTypes(i).GetDotNetType()
                                End If
                                If genericParam Is Nothing Then
                                    genericParam = genParams(i)
                                End If
                                passedParams.Add(genericParam)
                            Next
                            Try
                                toReturn = toReturn.MakeGenericType(passedParams.ToArray())
                            Catch ex As Exception
                                ExceptionHelper.LogException(ex)
                                toReturn = Nothing
                            End Try
                        End If
                    End If
                End If

                Return toReturn
            End Get
        End Property

        '<ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        '<ConditionalVisible("TypeSelector", False, True, TypeSelector.BrowseHierarchy)> _
        'Public ReadOnly Property CurrentlySelectedTypeName As String
        '    Get
        '        Return ReflectionHelper.GetSafeTypeName(CurrentlySelectedType)
        '    End Get
        'End Property

        <DefaultValue("")> _
        <ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        <ConditionalVisible("TypeSelector", False, True, TypeSelector.NewType)> _
        <Required(True)> _
        <Width(500)> _
        <LineCount(3)> _
        <AutoPostBack()> _
        <XmlIgnore> _
        Public Property EditableTypeName As String


        '<ConditionalVisible("TypeSelector", False, True, TypeSelector.NewType, TypeSelector.BrowseHierarchy)> _
        '<ActionButton(IconName.Refresh, IconOptions.Normal)> _
        'Public Sub Refresh(ByVal pe As AriciePropertyEditorControl)
        'End Sub


        <ConditionalVisible("Mode", False, True, TypePickerMode.SelectType)> _
        <ConditionalVisible("TypeSelector", False, True, TypeSelector.NewType, TypeSelector.BrowseHierarchy)> _
        <ActionButton("~/images/action_import.gif")> _
        Public Sub ValidateNewType(ByVal pe As AriciePropertyEditorControl)
            Try
                Dim objType As Type = Nothing
                Select Case Me.TypeSelector
                    Case TypeSelector.NewType
                        objType = Me.GetDotNetType(EditableTypeName, True)
                    Case TypeSelector.BrowseHierarchy
                        objType = CurrentlySelectedType
                End Select
                If objType IsNot Nothing Then
                    Me.SetTypeName(AddCommonType(ReflectionHelper.GetSafeTypeName(objType)))
                    Me.TypeSelector = ComponentModel.TypeSelector.CommonTypes
                    pe.DisplayMessage("Type correctly validated", DotNetNuke.UI.Skins.Controls.ModuleMessage.ModuleMessageType.GreenSuccess)
                Else
                    pe.DisplayMessage("Type did not validate", DotNetNuke.UI.Skins.Controls.ModuleMessage.ModuleMessageType.YellowWarning)
                End If
                pe.ItemChanged = True
            Catch ex As Exception
                Dim newEx As New ApplicationException("There was an error trying to create your type. See the complete Stack for more details", ex)
                Throw newEx
            End Try
        End Sub

        <XmlIgnore()> _
        Public Overridable ReadOnly Property FullName() As String
            Get
                Dim objType As Type = Me.GetDotNetType()
                If objType IsNot Nothing Then
                    Return ReflectionHelper.GetSafeTypeName(objType)
                End If
                Return String.Empty
            End Get
        End Property


        Public Function GetDotNetType() As Type
            If _Type Is Nothing Then

                Select Case PickerMode
                    Case TypePickerMode.SelectType
                        If Not String.IsNullOrEmpty(Me._TypeName) Then
                            _Type = Me.GetDotNetType(Me._TypeName, False)
                        End If
                    Case TypePickerMode.CustomObject
                        Try
                            Dim objProps = Me.CustomObject.EvaluateVariables(DnnContext.Current, DnnContext.Current)
                            If Me.CustomTypeIdentifier.Enabled Then
                                _Type = objProps.ToCustomObject(Me.CustomTypeIdentifier.Entity.Name).GetType()
                            Else
                                _Type = objProps.ToCustomObject("").GetType()
                            End If
                        Catch ex As Exception
                            ExceptionHelper.LogException(ex)
                        End Try

                        'AddCommonType(ReflectionHelper.GetSafeTypeName(_Type), Me)
                End Select
            End If

            Return _Type
        End Function

        Public Function GetDotNetType(ByVal strTypeName As String, ByVal throwException As Boolean) As Type
            If Not String.IsNullOrEmpty(strTypeName) Then
                Return ReflectionHelper.CreateType(strTypeName, throwException)
            End If
            Return Nothing
        End Function

        Public Function GetSelector(ByVal propertyName As String) As IList Implements ISelector.GetSelector
            Dim toReturn As IList = Nothing
            Select Case propertyName
                Case "CommonType"
                    toReturn = DirectCast(GetSelectorG(propertyName), List(Of DotNetType))
                Case "AssemblyNameSelect"
                    toReturn = DirectCast(GetSelectorG1(propertyName), List(Of AssemblyName))
                Case "NamespaceSelect"
                    toReturn = DirectCast(GetSelectorG2(propertyName), List(Of String))
                Case "TypeNameSelect"
                    toReturn = DirectCast(GetSelectorG(propertyName), List(Of DotNetType))
            End Select
            Return toReturn
        End Function




        Public Overridable Function GetSelectorG(ByVal propertyName As String) As IList(Of DotNetType) Implements ISelector(Of DotNetType).GetSelectorG
            Dim toReturn As New List(Of DotNetType)
            Select Case propertyName
                Case "CommonType"
                    If _CommonTypes.Count = 0 Then
                        AddCommonType(ReflectionHelper.GetSafeTypeName(GetType(Object)))
                        AddCommonType(ReflectionHelper.GetSafeTypeName(GetType(String)))
                        AddCommonType(ReflectionHelper.GetSafeTypeName(GetType(Integer)))
                    End If
                    toReturn.AddRange(From tmpType In _CommonTypes.Values.Distinct() Where tmpType IsNot Nothing)
                Case "TypeNameSelect"
                    If Not String.IsNullOrEmpty(AssemblyNameSelect) Then
                        Dim objAssembly As Assembly = Assembly.Load(New AssemblyName(AssemblyNameSelect))
                        If objAssembly IsNot Nothing Then
                            Try
                                toReturn.AddRange(From objType In objAssembly.GetTypes() _
                                                  Where objType.Namespace = Me.NamespaceSelect _
                                                  AndAlso Not String.IsNullOrEmpty(objType.AssemblyQualifiedName) _
                                                  From objTypeOrChild In New Type() {objType} _
                                                      .Union(objType.GetNestedTypes() _
                                                          .Where(Function(objNestedType) Not objNestedType.AssemblyQualifiedName.IsNullOrEmpty()))
                                                  Select New DotNetType(objTypeOrChild))
                            Catch ex As ReflectionTypeLoadException
                                Dim sb As New StringBuilder()
                                For Each exSub As Exception In ex.LoaderExceptions
                                    sb.AppendLine(exSub.Message)
                                    Dim exFileNotFound As FileNotFoundException = TryCast(exSub, FileNotFoundException)
                                    If exFileNotFound IsNot Nothing Then
                                        If Not String.IsNullOrEmpty(exFileNotFound.FusionLog) Then
                                            sb.AppendLine("Fusion Log:")
                                            sb.AppendLine(exFileNotFound.FusionLog)
                                        End If
                                    End If
                                    sb.AppendLine()
                                Next
                                'Display or log the error based on your application.
                                Dim errorMessage As String = sb.ToString()
                                Dim exc As New ApplicationException("Assembly Load Exception. Fusion Log: " & errorMessage, ex)
                                ExceptionHelper.LogException(exc)
                            End Try

                        End If
                    End If
            End Select
            Return toReturn.OrderBy(
                Function(objDotNetType)
                    If objDotNetType.GetDotNetType() IsNot Nothing Then
                        Return objDotNetType.GetDotNetType().Name
                    End If
                    Return ""
                End Function).ToList()
        End Function

        'Public Overrides Function Equals(ByVal obj As Object) As Boolean
        '    If TypeOf (obj) Is DotNetType Then
        '        Return Me.TypeName.Equals(DirectCast(obj, DotNetType).TypeName)
        '    End If
        '    Return False
        'End Function

        'Public Overrides Function GetHashCode() As Integer
        '    Return Me.TypeName.GetHashCode()
        'End Function

        Public Overridable Function GetSelectorG1(ByVal propertyName As String) As IList(Of AssemblyName) Implements ISelector(Of AssemblyName).GetSelectorG
            Dim toReturn As New List(Of AssemblyName)
            For Each objAssembly As Assembly In AppDomain.CurrentDomain.GetAssemblies()
                Dim toAdd As AssemblyName = objAssembly.GetName()
                'Dim objTypes As Type() = objAssembly.GetTypes()
                If Not (toAdd.Name.StartsWith("Anonymously Hosted") OrElse toAdd.Name.StartsWith("App_") OrElse toAdd.Name.StartsWith("Microsoft.GeneratedCode")) Then
                    toReturn.Add(toAdd)
                End If
            Next
            toReturn.Sort(Function(objAssemblyName1, objAssemblyName2) String.Compare(objAssemblyName1.FullName, objAssemblyName2.FullName, StringComparison.InvariantCultureIgnoreCase))
            Return toReturn
        End Function

        Public Overridable Function GetSelectorG2(ByVal propertyName As String) As IList(Of String) Implements ISelector(Of String).GetSelectorG
            Dim toReturnSet As New HashSet(Of String)
            If Not String.IsNullOrEmpty(AssemblyNameSelect) Then
                Dim objAssembly As Assembly = Assembly.Load(New AssemblyName(AssemblyNameSelect))
                If objAssembly IsNot Nothing Then
                    Try
                        For Each objType As Type In objAssembly.GetTypes()
                            If Not String.IsNullOrEmpty(objType.Namespace) Then
                                If Not toReturnSet.Contains(objType.Namespace) Then
                                    toReturnSet.Add(objType.Namespace)
                                End If
                            End If
                        Next
                    Catch ex As ReflectionTypeLoadException
                        Dim sb As New StringBuilder()
                        For Each exSub As Exception In ex.LoaderExceptions
                            sb.AppendLine(exSub.Message)
                            Dim exFileNotFound As FileNotFoundException = TryCast(exSub, FileNotFoundException)
                            If exFileNotFound IsNot Nothing Then
                                If Not String.IsNullOrEmpty(exFileNotFound.FusionLog) Then
                                    sb.AppendLine("Fusion Log:")
                                    sb.AppendLine(exFileNotFound.FusionLog)
                                End If
                            End If
                            sb.AppendLine()
                        Next
                        'Display or log the error based on your application.
                        Dim errorMessage As String = sb.ToString()
                        Dim exc As New ApplicationException("Assembly Load Exception. Fusion Log: " & errorMessage, ex)
                        ExceptionHelper.LogException(exc)
                    End Try
                End If
            End If
            Dim toReturn As List(Of String) = toReturnSet.ToList()
            toReturn.Sort()
            Return toReturn
        End Function


        Private Sub SetTypeName(ByVal value As String)
            Me._TypeName = value
            Me._EditableTypeName = Me._TypeName
            Me._Type = Nothing
            Dim objType As Type = Me.GetDotNetType(_TypeName, False)
            If objType IsNot Nothing Then
                If objType.IsArray Then
                    Me.MakeArray = True
                    Me.Rank = objType.GetArrayRank()
                    Me.ArrayType = New DotNetType(objType.GetElementType())
                Else
                    Me.AssemblyNameSelect = objType.Assembly.GetName().FullName
                    Me.NamespaceSelect = objType.Namespace
                    If objType.IsGenericType Then
                        Me._TypeNameSelect = ReflectionHelper.GetSafeTypeName(objType.GetGenericTypeDefinition())
                        PrepareGenericType(objType)
                    Else
                        Me._TypeNameSelect = ReflectionHelper.GetSafeTypeName(objType)
                    End If
                End If
            End If
        End Sub

        Private Sub PrepareGenericType(ByVal objType As Type)
            Dim argTypes As Type() = objType.GetGenericArguments()
            If argTypes.Count > 0 Then
                Me.GenericTypes.Clear()
                For Each argType As Type In argTypes
                    Me.GenericTypes.Add(New DotNetType(argType))
                Next
            End If
        End Sub


        'Public Overrides Function Equals(obj As Object) As Boolean
        '    If obj IsNot Nothing Then
        '        If TypeOf obj Is DotNetType Then
        '            Return Me.TypeName = DirectCast(obj, DotNetType).TypeName
        '        End If
        '    End If
        '    Return False
        'End Function

        'Public Overrides Function GetHashCode() As Integer
        '    Return Me.TypeName.GetHashCode()
        'End Function

    End Class

   


    Public Class DotNetType(Of TVariable)
        Inherits DotNetType
        Implements IGenericizer(Of TVariable)
        Implements IProviderConfig(Of IGenericizer(Of TVariable))


        Private _GenericVariableType As Type
        Private _TargetTypes As List(Of DotNetType)

        Public Sub New()

        End Sub

        Public Sub New(ByVal targetType As DotNetType)
            Me.TypeName = targetType.TypeName
        End Sub

        Public Sub New(ByVal genericVariableType As Type, ByVal ParamArray targetTypes As DotNetType())
            Me.TypeName = ReflectionHelper.MakeGenerics(genericVariableType, targetTypes.Select(Function(objDotNetType) objDotNetType.GetDotNetType())).AssemblyQualifiedName
        End Sub


        Public Function GetProvider() As Object Implements IProviderConfig.GetProvider
            Return GetTypedProvider()
        End Function

        Public Function GetTypedProvider() As IGenericizer(Of TVariable) Implements IProviderConfig(Of IGenericizer(Of TVariable)).GetTypedProvider
            Return Me
        End Function

        Public Sub SetConfig(ByVal config As IProviderConfig) Implements IProvider.SetConfig
            ReflectionHelper.MergeObjects(config, Me)
        End Sub

        Public Property Config As DotNetType(Of TVariable) Implements IProvider(Of DotNetType(Of TVariable)).Config
            Get
                Return Me
            End Get
            Set(value As DotNetType(Of TVariable))
                ReflectionHelper.MergeObjects(value, Me)
            End Set
        End Property



        Public Function GetNewProviderSettings() As TVariable Implements IProvider(Of DotNetType(Of TVariable), TVariable).GetNewProviderSettings
            'If Me.GenericVariableType IsNot Nothing Then
            '    Dim genType As Type = Me.Config.GenericVariableType
            '    If Me.TargetTypes.Count > 0 Then
            '        Dim targetTypes As New List(Of Type)
            '        For Each targetDotNetType As DotNetType In Me.Config.TargetTypes
            '            targetTypes.Add(targetDotNetType.GetDotNetType)
            '        Next
            '        genType = Me.Config.GenericVariableType.MakeGenericType(targetTypes.ToArray)
            '    End If
            '    Return ReflectionHelper.CreateObject(Of TVariable)(genType.FullName)
            'End If
            Return ReflectionHelper.CreateObject(Of TVariable)(Me.TypeName)
        End Function

        'Public Class GenericsProvider
        '    Implements IGenericizer(Of TVariable)

        '    Private _Config As DotNetType(Of TVariable)
        '    'Private _Settings As TVariable


        '    Public Sub SetConfig(ByVal config As IProviderConfig) Implements IProvider.SetConfig
        '        Me._Config = DirectCast(config, DotNetType(Of TVariable))
        '    End Sub

        '    'Public Sub SetSettings(ByVal settings As IProviderSettings) Implements IProvider.SetSettings
        '    '    Me._Settings = DirectCast(settings, TVariable)
        '    'End Sub

        '    Public Property Config() As DotNetType(Of TVariable) Implements IProvider(Of DotNetType(Of TVariable)).Config
        '        Get
        '            Return Me._Config
        '        End Get
        '        Set(ByVal value As DotNetType(Of TVariable))
        '            Me._Config = value
        '        End Set
        '    End Property

        '    Public Function GetNewProviderSettings() As TVariable Implements IProvider(Of DotNetType(Of TVariable), TVariable).GetNewProviderSettings
        '        If Me.Config.GenericVariableType IsNot Nothing Then
        '            Dim genType As Type = Me.Config.GenericVariableType
        '            If Me.Config.TargetTypes.Count > 0 Then
        '                Dim targetTypes As New List(Of Type)
        '                For Each targetDotNetType As DotNetType In Me.Config.TargetTypes
        '                    targetTypes.Add(targetDotNetType.GetDotNetType)
        '                Next
        '                genType = Me.Config.GenericVariableType.MakeGenericType(targetTypes.ToArray)
        '            End If
        '            Return ReflectionHelper.CreateObject(Of TVariable)(genType.FullName)
        '        End If
        '        Return ReflectionHelper.CreateObject(Of TVariable)(Me.Config.TypeName)

        '    End Function

        '    'Public Property Settings() As TVariable Implements IProvider(Of DotNetType(Of TVariable), TVariable).Settings
        '    '    Get
        '    '        Return Me._Settings
        '    '    End Get
        '    '    Set(ByVal value As TVariable)
        '    '        Me._Settings = value
        '    '    End Set
        '    'End Property
        'End Class


    End Class

   


End Namespace