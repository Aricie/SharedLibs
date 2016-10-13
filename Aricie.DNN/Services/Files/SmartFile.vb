﻿Imports Aricie.DNN.ComponentModel
Imports System.ComponentModel
Imports Aricie.DNN.UI.Attributes
Imports System.Web.Configuration
Imports Aricie.ComponentModel
Imports Aricie.Cryptography
Imports Aricie.Security.Cryptography
Imports Aricie.DNN.Security.Cryptography
Imports DotNetNuke.UI.Skins.Controls
Imports DotNetNuke.UI.WebControls
Imports System.Xml.Serialization
Imports Aricie.DNN.UI.WebControls
Imports DotNetNuke.Services.FileSystem
Imports Aricie.Services
Imports System.Text
Imports System.Globalization
Imports DotNetNuke.Security.Permissions
Imports DotNetNuke.Entities.Users
Imports System.Xml
Imports DotNetNuke.Services.Localization
Imports System.IO
Imports Aricie.DNN.UI.WebControls.EditControls

Namespace Services.Files

    <Flags()>
    Public Enum SmartFileState
        Clear = 0
        Signed = 1
        Compressed = 2
        Encrypted = 4
    End Enum


    Public Enum PayLoadFormat
        None
        UTF8String
        Bytes
        Base64String
    End Enum

    <Serializable()>
    Public Class SmartFile
        'Inherits SmartFileInfo

        Private _encrypter As IEncrypter
        Protected _SaltBytes As Byte() = {}
        Private _PayLoad As Byte()
        Private _DNNFile As DotNetNuke.Services.FileSystem.FileInfo


        Public Sub New()
        End Sub

        Public Sub New(key As EntityKey)
            Me.Key = key
        End Sub

        Public Sub New(key As EntityKey, encrypter As IEncrypter)
            Me.New(key)
            Me._encrypter = encrypter
        End Sub

        <Browsable(False)>
        Public ReadOnly Property DNNFile As DotNetNuke.Services.FileSystem.FileInfo
            Get
                Return _DNNFile
            End Get
        End Property


        <ExtendedCategory("Key")>
        <IsReadOnly(True)>
        Public Property Key As EntityKey

        <ExtendedCategory("Content")>
        <IsReadOnly(True)>
        Public Property State As SmartFileState = SmartFileState.Clear



        <Browsable(False)>
        Public ReadOnly Property Signed As Boolean
            Get
                Return (State And SmartFileState.Signed) = SmartFileState.Signed
            End Get
        End Property

        <Browsable(False)>
        Public ReadOnly Property Compressed As Boolean
            Get
                Return (State And SmartFileState.Compressed) = SmartFileState.Compressed
            End Get
        End Property

        <Browsable(False)>
        Public ReadOnly Property Encrypted As Boolean
            Get
                Return (State And SmartFileState.Encrypted) = SmartFileState.Encrypted
            End Get
        End Property




        <ExtendedCategory("Content")>
        Public ReadOnly Property HasEncrypter As Boolean
            Get
                Return _encrypter IsNot Nothing
            End Get
        End Property

        Public Sub SetEncrypter(ByVal encrypter As IEncrypter)
            _encrypter = encrypter
        End Sub


        <Browsable(False)>
        <XmlIgnore()>
        Public ReadOnly Property SaltBytes As Byte()
            Get
                Return _SaltBytes
            End Get
        End Property

        <ExtendedCategory("Content")>
        <ConditionalVisible("Encrypted", False, True)>
        <IsReadOnly(True)>
        Public Property Salt As String
            Get
                If Me.Encrypted Then
                    Return Convert.ToBase64String(_SaltBytes)
                End If
                Return ""
            End Get
            Set(value As String)
                _SaltBytes = Convert.FromBase64String(value)
            End Set
        End Property


        <LineCount(20)>
        <Width(500)>
        <Editor(GetType(CustomTextEditControl), GetType(EditControl))>
        <ExtendedCategory("Content")>
        <ConditionalVisible("EditPayLoadFormat", True, True, PayLoadFormat.None)>
        <XmlIgnore()>
        Public Property EditPayLoad As String
            Get
                Return EditPayload(EditPayLoadFormat)
            End Get
            Set(value As String)
                EditPayload(EditPayLoadFormat) = value
            End Set
        End Property

        <ExtendedCategory("Content")>
        <XmlIgnore()>
        Public Property EditPayLoadFormat As PayLoadFormat = PayLoadFormat.None
        '<ExtendedCategory("Content")> _
        'Public Property ShowPayLoad As Boolean

        ' <AutoPostBack()> _
        '<ExtendedCategory("Content")> _
        '       <XmlIgnore()> _
        ' Public Property ShowLineBreaks As Boolean


        <Browsable(False)>
        Public Property EditPayload(format As PayLoadFormat) As String
            Get
                Select Case format
                    Case PayLoadFormat.UTF8String
                        Return Encoding.UTF8.GetString(_PayLoad)
                    Case PayLoadFormat.Bytes
                        Return BitConverter.ToString(_PayLoad)
                    Case PayLoadFormat.Base64String
                        Return Convert.ToBase64String(_PayLoad)
                    Case Else
                        Return String.Empty
                End Select
            End Get
            Set(value As String)
                If format <> PayLoadFormat.None Then
                    Dim newValue As Byte() = {}
                    Select Case format
                        Case PayLoadFormat.UTF8String
                            newValue = Encoding.UTF8.GetBytes(value)
                        Case PayLoadFormat.Bytes
                            'newValue = Encoding.UTF8.GetString(Common.StringToByteArray(value))
                            newValue = Common.BitConverterStringToByteArray(value)
                        Case PayLoadFormat.Base64String
                            newValue = Convert.FromBase64String(value)
                    End Select
                    _PayLoad = newValue
                End If
            End Set
        End Property

        <XmlIgnore()>
        <Browsable(False)>
        Public Property PayLoad As Byte()
            Get
                Return _PayLoad
            End Get
            Set(value As Byte())
                _PayLoad = value
            End Set
        End Property



        <Browsable(False)>
        Public Property StoreBinPayload As Byte()
            Get
                If (Me.Compressed OrElse Me.Encrypted) Then
                    Return _PayLoad
                End If
                Return Nothing
            End Get
            Set(value As Byte())
                Me._PayLoad = value
            End Set
        End Property


        <Browsable(False)>
        Public Property StorePlainPayload As CData
            Get
                If Not (Me.Compressed OrElse Me.Encrypted) Then
                    Return Encoding.UTF8.GetString(_PayLoad)
                End If
                Return Nothing
            End Get
            Set(value As CData)
                Me._PayLoad = Encoding.UTF8.GetBytes(value)
            End Set
        End Property


        <Browsable(False)>
        <XmlIgnore()>
        Public Property PayLoadAsXmlDocument As XmlDocument
            Get
                Using inputStream As New MemoryStream(Me._PayLoad)
                    Dim doc = New XmlDocument()
                    doc.Load(inputStream)
                    Return doc
                End Using
            End Get
            Set(value As XmlDocument)
                Using outStream As New MemoryStream()
                    value.Save(outStream)
                    _PayLoad = outStream.ToArray()
                End Using
            End Set
        End Property


        <Browsable(False)>
        Public ReadOnly Property Size As Integer
            Get
                Return Me._PayLoad.Length
            End Get
        End Property

        <ExtendedCategory("Content")>
        Public ReadOnly Property FormattedSize As String
            Get
                Return BytesToString(Size)
            End Get
        End Property

        <ExtendedCategory("Content")>
        Public ReadOnly Property MD5Checksum As String
            Get
                Return Me._PayLoad.Hash(HashProvider.MD5)
            End Get
        End Property

        <ExtendedCategory("Content")>
        Public ReadOnly Property Sha256Checksum As String
            Get
                Return Me._PayLoad.Hash(HashProvider.SHA256)
            End Get
        End Property

        <ExtendedCategory("Content")>
        <XmlIgnore()>
        Public Property UseCustomEncryption As Boolean

        <ExtendedCategory("Content")>
        <ConditionalVisible("UseCustomEncryption", False, True)>
        <XmlIgnore()>
        Public Property CustomEncryption As New EncryptionInfo()


#Region "Public Methods"



        <ExtendedCategory("Content")>
        <ConditionalVisible("Signed", True, True)>
        <ConditionalVisible("Encrypted", True, True)>
        <ConditionalVisible("Compressed", True, True)>
        <ActionButton(IconName.Key, IconOptions.Normal)>
        Public Overloads Sub Sign(ape As AriciePropertyEditorControl)
            Me.Sign()
            ape.ItemChanged = True
            Dim message As String = Localization.GetString("SmartFileSigned.Message", ape.LocalResourceFile)
            ape.DisplayMessage(message, ModuleMessage.ModuleMessageType.GreenSuccess)
        End Sub


        '<ConditionalVisible("HasEncrypter", False, True)> _
        <ExtendedCategory("Content")>
        <ConditionalVisible("Signed", False, True)>
        <ConditionalVisible("Encrypted", True, True)>
        <ConditionalVisible("Compressed", True, True)>
        <ActionButton(IconName.CheckSquareO, IconOptions.Normal)>
        Public Overloads Sub Verify(ape As AriciePropertyEditorControl)
            Dim message As String
            Dim messageType As ModuleMessage.ModuleMessageType
            If Me.Verify() Then
                message = Localization.GetString("SignatureVerified.Message", ape.LocalResourceFile)
                messageType = ModuleMessage.ModuleMessageType.GreenSuccess
            Else
                message = Localization.GetString("SignatureFailedToVerify.Message", ape.LocalResourceFile)
                messageType = ModuleMessage.ModuleMessageType.RedError
            End If
            ape.DisplayMessage(message, messageType)
        End Sub

        <ExtendedCategory("Content")>
        <ConditionalVisible("Signed", False, True)>
        <ConditionalVisible("Encrypted", True, True)>
        <ConditionalVisible("Compressed", True, True)>
        <ActionButton(IconName.Eraser, IconOptions.Normal)>
        Public Overloads Sub RemoveSignature(ape As AriciePropertyEditorControl)
            Me.RemoveSignature()
            ape.ItemChanged = True
            Dim message As String = Localization.GetString("SignatureRemoved.Message", ape.LocalResourceFile)
            Dim messageType As ModuleMessage.ModuleMessageType = ModuleMessage.ModuleMessageType.GreenSuccess
            ape.DisplayMessage(message, messageType)
        End Sub

        <ExtendedCategory("Content")>
        <ConditionalVisible("Compressed", True, True)>
        <ConditionalVisible("Encrypted", True, True)>
        <ActionButton(IconName.Compress, IconOptions.Normal)>
        Public Sub Compress(ape As AriciePropertyEditorControl)
            Me.Compress()
            If Me.EditPayLoadFormat = PayLoadFormat.UTF8String Then
                Me.EditPayLoadFormat = PayLoadFormat.Base64String
            End If
            ape.ItemChanged = True
            Dim message As String = Localization.GetString("SmartFileCompressed.Message", ape.LocalResourceFile)
            ape.DisplayMessage(message, ModuleMessage.ModuleMessageType.GreenSuccess)
        End Sub

        <ExtendedCategory("Content")>
        <ConditionalVisible("Compressed", False, True)>
        <ConditionalVisible("Encrypted", True, True)>
        <ActionButton(IconName.Expand, IconOptions.Normal)>
        Public Sub Decompress(ape As AriciePropertyEditorControl)
            Me.Decompress()
            ape.ItemChanged = True
            Dim message As String = Localization.GetString("SmartFileDecompressed.Message", ape.LocalResourceFile)
            ape.DisplayMessage(message, ModuleMessage.ModuleMessageType.GreenSuccess)
        End Sub

        <ExtendedCategory("Content")>
        <ConditionalVisible("Encrypted", False, True)>
        <ActionButton(IconName.Unlock, IconOptions.Normal)>
        Public Sub Decrypt(ape As AriciePropertyEditorControl)
            Me.Decrypt()
            ape.ItemChanged = True
            Dim message As String = Localization.GetString("SmartFileDecrypted.Message", ape.LocalResourceFile)
            ape.DisplayMessage(message, ModuleMessage.ModuleMessageType.GreenSuccess)
        End Sub

        <ExtendedCategory("Content")>
        <ConditionalVisible("Encrypted", True, True)>
        <ActionButton(IconName.Lock, IconOptions.Normal)>
        Public Sub Encrypt(ape As AriciePropertyEditorControl)
            Me.Encrypt()
            If Me.EditPayLoadFormat = PayLoadFormat.UTF8String Then
                Me.EditPayLoadFormat = PayLoadFormat.Base64String
            End If
            ape.ItemChanged = True
            Dim message As String = Localization.GetString("SmartFileEncrypted.Message", ape.LocalResourceFile)
            ape.DisplayMessage(message, ModuleMessage.ModuleMessageType.GreenSuccess)
        End Sub



        Public Sub Decrypt()
            Try
                If Me.Encrypted Then
                    Dim newPayLoad As Byte()
                    If UseCustomEncryption Then
                        newPayLoad = Me.CustomEncryption.Decrypt(Me._PayLoad, Me._SaltBytes)
                    ElseIf _encrypter IsNot Nothing Then
                        newPayLoad = _encrypter.Decrypt(Me._PayLoad, Me._SaltBytes)
                    Else
                        newPayLoad = Me._PayLoad.Decrypt(GetDefaultKey(), {}, Me._SaltBytes)
                    End If
                    Me.DecryptInternal(newPayLoad)
                End If
            Catch ex As Exception
                Throw New ApplicationException("Value was encrypted with a distinct key or encrypted bytes were corrupted.", ex)
            End Try
        End Sub


        Public Sub Encrypt()
            If Not Me.Encrypted Then
                Dim objSalt As Byte() = Nothing
                Dim newPayload As Byte()
                If UseCustomEncryption Then
                    newPayload = Me.CustomEncryption.DoEncrypt(Me.PayLoad, objSalt)
                    Me.Encrypt(newPayload, objSalt)
                ElseIf _encrypter IsNot Nothing Then
                    newPayload = _encrypter.Encrypt(Me._PayLoad, objSalt)
                    Me.Encrypt(newPayload, objSalt)
                Else
                    newPayload = _PayLoad.Encrypt(GetDefaultKey(), {}, objSalt)
                    Me.EncryptInternal(newPayload, objSalt)
                End If
            End If
        End Sub



        Public Overloads Sub Sign()
            If Me.Encrypted OrElse Me.Compressed Then
                Throw New ApplicationException("Encrypted or compressed Content cannot be Signed")
            End If
            If Not Me.Signed Then
                SyncLock Me
                    Dim objEncrypter As IEncrypter
                    If UseCustomEncryption Then
                        objEncrypter = Me.CustomEncryption
                    ElseIf Me._encrypter IsNot Nothing Then
                        objEncrypter = _encrypter
                    Else
                        Throw New ApplicationException("Cannot sign a smart file without an ecrypter")
                    End If
                    Dim tempDoc As XmlDocument = Me.PayLoadAsXmlDocument
                    objEncrypter.Sign(tempDoc)
                    Me.PayLoadAsXmlDocument = tempDoc
                    Me.State = Me.State Or SmartFileState.Signed
                End SyncLock
            End If
        End Sub

        Public Overloads Function Verify() As Boolean
            If Not Me.Signed Then
                Throw New ApplicationException("Unsigned document cannot be verified")
            Else
                SyncLock Me
                    Dim objEncrypter As IEncrypter
                    If UseCustomEncryption Then
                        objEncrypter = Me.CustomEncryption
                    ElseIf Me._encrypter IsNot Nothing Then
                        objEncrypter = _encrypter
                    Else
                        Throw New ApplicationException("Cannot verify a smart file without an ecrypter")
                    End If
                    Return objEncrypter.Verify(Me.PayLoadAsXmlDocument)
                End SyncLock
            End If
        End Function




        Public Overloads Sub RemoveSignature()
            If Not Me.Signed Then
                Throw New ApplicationException("Unsigned document cannot have signature removed")
            Else
                SyncLock Me
                    Dim tempDoc As XmlDocument = Me.PayLoadAsXmlDocument
                    CryptoHelper.RemoveSignatureFromXmlDocument(tempDoc)
                    Me.PayLoadAsXmlDocument = tempDoc
                    Me.State = (Me.State And Not SmartFileState.Signed)
                End SyncLock
            End If
        End Sub

        Public Sub Compress()
            If Me.Encrypted Then
                Throw New ApplicationException("Encrypted Content cannot be compressed")
            End If
            If Not Me.Compressed Then
                SyncLock Me
                    _PayLoad = _PayLoad.Compress(CompressionMethod.Gzip)
                    Me.State = Me.State Or SmartFileState.Compressed
                End SyncLock
            End If
        End Sub

        Public Sub Decompress()
            If Me.Encrypted Then
                Throw New ApplicationException("Encrypted Content cannot be decompressed")
            End If
            If Me.Compressed Then
                SyncLock Me

                    _PayLoad = _PayLoad.Decompress(CompressionMethod.Gzip)
                    Me.State = (Me.State And Not SmartFileState.Compressed)
                End SyncLock
            End If
        End Sub


        Public Function GetDefaultKey() As Byte()
            Return Encoding.UTF8.GetBytes(New MachineKeySection().DecryptionKey())
        End Function


        'Public Function GetKey() As String
        '    If UseCustomKey Then
        '        Return Me.CustomKey
        '    Else
        '        Return New MachineKeySection().DecryptionKey
        '    End If
        'End Function

        Public Sub Decrypt(newPayload As Byte())
            SyncLock Me
                Me.DecryptInternal(newPayload)
            End SyncLock
        End Sub

        Public Sub Encrypt(ByVal newPayLoad As Byte(), ByVal objSalt As Byte())
            SyncLock Me
                Me.EncryptInternal(newPayLoad, objSalt)
            End SyncLock
        End Sub

        Public Overridable Sub UnWrap()
            Me.Decrypt()
            Me.Decompress()
            If Me.Signed Then
                If Not Me.Verify() Then
                    Throw New ApplicationException("Smart file integrity signature does not verify against the provided encryption")
                End If
            End If
        End Sub

        Public Sub Wrap(settings As SmartFileInfo)
            If settings.Sign AndAlso Not Me.Signed AndAlso Not Me.Compressed AndAlso Not Me.Encrypted AndAlso Me._encrypter IsNot Nothing Then
                Me.Sign()
            End If
            If settings.Compress AndAlso Not Me.Compressed AndAlso Not Me.Encrypted Then
                Me.Compress()
            End If
            If settings.Encrypt Then
                Me.Encrypt()
            End If
        End Sub



#End Region




#Region "Shared Methods"


        Public Shared Function GetFileInfo(key As EntityKey, settings As SmartFileInfo) As DotNetNuke.Services.FileSystem.FileInfo
            Dim objPath As String = settings.GetPath(key)
            Dim folderPath As String = settings.GetFolderPath(key)
            Dim objFolderInfo As FolderInfo = ObsoleteDNNProvider.Instance.GetFolderFromPath(key.PortalId, folderPath)
            If objFolderInfo IsNot Nothing Then
                Dim fileName As String = System.IO.Path.GetFileName(objPath)
                Return ObsoleteDNNProvider.Instance.GetFile(objFolderInfo, fileName)
            End If
            Return Nothing
        End Function

        'Public Shared Function LoadAndRead(Of T As New)(key As EntityKey, settings As SmartFileInfo) As T
        '    Dim toReturn As T
        '    Dim objSmartFile As SmartFile(Of T) = LoadSmartFile(Of T)(key, settings)
        '    If objSmartFile IsNot Nothing Then
        '        toReturn = objSmartFile.Value
        '    End If
        '    Return toReturn
        'End Function

        Public Shared Sub DeleteSmartFile(Of T As New)(key As EntityKey, settings As SmartFileInfo)
            Dim objFileInfo As DotNetNuke.Services.FileSystem.FileInfo = GetFileInfo(key, settings)
            ObsoleteDNNProvider.Instance.DeleteFile(objFileInfo)
        End Sub


        Public Shared Function LoadSmartFile(Of T As New)(key As EntityKey, settings As SmartFileInfo) As SmartFile(Of T)
            Dim objFileInfo As DotNetNuke.Services.FileSystem.FileInfo = GetFileInfo(key, settings)
            Dim toReturn As SmartFile(Of T) = LoadSmartFile(Of T)(objFileInfo)
            If toReturn IsNot Nothing Then
                If Not settings.CheckSmartFile(toReturn) Then
                    Throw New ApplicationException(String.Format("smart file for key {0} at path {1} didn't match security settings", key.ToString(), objFileInfo.PhysicalPath))
                End If
                toReturn.SetEncrypter(settings.Encryption)
            End If
            Return toReturn
        End Function

        Private Const SmartFilesDicoKey As String = "SmartFilesDico"
        Private Const SmartFilesToSaveKey As String = "SmartFilesToSave"

        Public Shared Function GetSmartFileInfo(ByVal globalVars As IContextLookup) As SmartFileInfo
            Return TryCast(globalVars.Items("SmartFileInfo"), SmartFileInfo)
        End Function

        Public Shared Sub SetSmartFileInfo(ByVal globalVars As IContextLookup, value As SmartFileInfo) 
            globalVars.Items("SmartFileInfo") = value
        End Sub

        Public Shared Function LoadSmartFile(key As EntityKey, ByVal globalVars As IContextLookup, setToSave As Boolean, Optional settings As SmartFileInfo =  Nothing) As SmartFile
            If settings Is Nothing
                settings = GetSmartFileInfo(globalVars)
            End If
            Dim toReturn As SmartFile = LoadSmartFile(key,settings)
            Dim dicoSmartFiles As Dictionary(Of EntityKey, SmartFile)
            Dim dicoObj As Object= Nothing
            If Not globalVars.Items.TryGetValue(SmartFilesDicoKey, dicoObj) Then
                dicoObj = New Dictionary(Of EntityKey, SmartFile)()
                globalVars.Items(SmartFilesDicoKey) = dicoObj
            End If
            dicoSmartFiles = DirectCast(dicoObj, Dictionary(Of EntityKey, SmartFile))
            dicoSmartFiles(key) = toReturn
            If setToSave Then
                Dim smartFilesToSave As HashSet(Of EntityKey)
                Dim objSmartFilesToSave As Object = Nothing
                If Not globalVars.Items.TryGetValue(SmartFilesToSaveKey, objSmartFilesToSave) Then
                    objSmartFilesToSave = New HashSet(Of EntityKey)
                    globalVars.Items(SmartFilesToSaveKey) = objSmartFilesToSave
                End If
                smartFilesToSave = DirectCast(objSmartFilesToSave, HashSet(Of EntityKey))
                smartFilesToSave.Add(key)
            End If
            Return toReturn
        End Function

        Public Shared Sub SaveSmartFiles(ByVal globalVars As IContextLookup)
            Dim smartFilesToSave As HashSet(Of EntityKey)
            Dim objSmartFilesToSave As Object= Nothing
            If globalVars.Items.TryGetValue(SmartFilesToSaveKey, objSmartFilesToSave) Then
                smartFilesToSave = DirectCast(objSmartFilesToSave, HashSet(Of EntityKey))
                Dim dicoSmartFiles As Dictionary(Of EntityKey, SmartFile)
                Dim dicoObj As Object= Nothing
                If globalVars.Items.TryGetValue(SmartFilesDicoKey, dicoObj) Then
                    dicoObj = New Dictionary(Of EntityKey, SmartFile)()
                    dicoSmartFiles = DirectCast(dicoObj, Dictionary(Of EntityKey, SmartFile))
                    For Each objKey As EntityKey In smartFilesToSave
                        Dim objSmartFile As SmartFile= Nothing
                        If dicoSmartFiles.TryGetValue(objKey, objSmartFile) Then
                            SmartFile.SaveSmartFile(objSmartFile, GetSmartFileInfo(globalVars))
                        End If
                    Next
                End If
            End If
        End sub

        Public Shared Function LoadSmartFile(key As EntityKey, settings As SmartFileInfo) As SmartFile
            Dim objFileInfo As DotNetNuke.Services.FileSystem.FileInfo = GetFileInfo(key, settings)
            Dim toReturn As SmartFile = LoadSmartFile(objFileInfo)
            If toReturn IsNot Nothing Then
                If Not settings.CheckSmartFile(toReturn) Then
                    Throw New ApplicationException(String.Format("smart file for key {0} at path {1} didn't match security settings", key.ToString(), objFileInfo.PhysicalPath))
                End If
                toReturn.SetEncrypter(settings.Encryption)
            End If
            Return toReturn
        End Function

        Public Shared Function LoadSmartFile(Of T As New)(objFileInfo As DotNetNuke.Services.FileSystem.FileInfo) As SmartFile(Of T)
            If objFileInfo IsNot Nothing Then
                Dim content As Byte() = ObsoleteDNNProvider.Instance.GetFileContent(objFileInfo)
                Dim toReturn As SmartFile(Of T) = ReflectionHelper.Deserialize(Of SmartFile(Of T))(content)
                toReturn._DNNFile = objFileInfo
                Return toReturn
            End If
            Return Nothing
        End Function

        Public Shared Function LoadSmartFile(objFileInfo As DotNetNuke.Services.FileSystem.FileInfo) As SmartFile
            If objFileInfo IsNot Nothing Then
                Dim content As Byte() = ObsoleteDNNProvider.Instance.GetFileContent(objFileInfo)
                Dim toReturn As SmartFile = ReflectionHelper.Deserialize(Of SmartFile)(content)
                toReturn._DNNFile = objFileInfo
                Return toReturn
            End If
            Return Nothing
        End Function

        Public Shared Function SaveSmartFile(value As SmartFile, settings As SmartFileInfo) As Boolean
            If value IsNot Nothing Then
                value.Wrap(settings)
                Dim objFolderPäth As String = settings.GetFolderPath(value.Key)
                Dim objFolderInfo As FolderInfo = ObsoleteDNNProvider.Instance.GetFolderFromPath(value.Key.PortalId, objFolderPäth)
                If objFolderInfo Is Nothing Then
                    Dim permissionUserId As Integer = -1
                    If (settings.GrantUserView OrElse settings.GrantUserEdit) AndAlso value.Key.UserName <> "" Then
                        Dim objUser As UserInfo = DotNetNuke.Entities.Users.UserController.GetUserByName(value.Key.PortalId, value.Key.UserName)
                        If objUser IsNot Nothing Then
                            permissionUserId = objUser.UserID
                        End If
                    End If
                    CreateSecureFoldersRecursive(value.Key.PortalId, objFolderPäth, permissionUserId, settings)
                    objFolderInfo = ObsoleteDNNProvider.Instance.GetFolderFromPath(value.Key.PortalId, objFolderPäth)
                End If
                If objFolderInfo IsNot Nothing Then
                    Dim objFileInfo As DotNetNuke.Services.FileSystem.FileInfo = GetFileInfo(value.Key, settings)
                    If objFileInfo Is Nothing Then
                        Dim objPath As String = settings.GetPath(value.Key)
                        objFileInfo = New DotNetNuke.Services.FileSystem.FileInfo() With {.FileName = System.IO.Path.GetFileName(objPath), .ContentType = "text/xml", .FolderId = objFolderInfo.FolderID, .PortalId = value.Key.PortalId, .StorageLocation = 2}
                    End If
                    Dim content As Byte()
                    content = ReflectionHelper.SerializeToBytes(value, False)
                    Dim result As Integer = ObsoleteDNNProvider.Instance.AddOrUpdateFile(objFileInfo, content, False)
                    If result < 0 Then
                        Throw New ApplicationException("Save failed, DNN returned " & result.ToString(CultureInfo.InvariantCulture))
                    End If
                    'End If
                Else
                    Throw New ApplicationException("Could not access Smart File Storage")
                End If
            Else
                Throw New ApplicationException("Cannot save null smartfile")
            End If

            Return True
        End Function

        Protected Shared Sub CreateSecureFoldersRecursive(portalId As Integer, path As String, permissionUserId As Integer, settings As SmartFileInfo)
            Dim objFolderInfo As FolderInfo = ObsoleteDNNProvider.Instance.GetFolderFromPath(portalId, path)
            If objFolderInfo Is Nothing Then
                Dim parentPath As String = path.TrimEnd("/"c)
                If parentPath.Contains("/"c) Then
                    parentPath = parentPath.Substring(0, parentPath.LastIndexOf("/"c))
                    CreateSecureFoldersRecursive(portalId, parentPath, -1, settings)
                End If
                Dim folder As New FolderInfo() With {
                    .PortalID = portalId,
                    .FolderPath = path,
                    .StorageLocation = 2,
                    .IsProtected = False,
                    .IsCached = False
                }
                ObsoleteDNNProvider.Instance.AddFolder(folder)
                folder = ObsoleteDNNProvider.Instance.GetFolderFromPath(folder.PortalID, folder.FolderPath)
                If folder IsNot Nothing Then
                    If permissionUserId > 0 Then
                        Dim fpc As New FolderPermissionController()
                        If settings.GrantUserView Then
                            Dim fp As New FolderPermissionInfo()
                            With fp
                                .UserID = permissionUserId
                                .FolderID = folder.FolderID
                                .FolderPath = folder.FolderPath
                                .PermissionCode = "VIEW"
                                .PermissionKey = "VIEW"
                            End With
                            fpc.AddFolderPermission(fp)
                        End If
                        If settings.GrantUserEdit Then
                            Dim fp As New FolderPermissionInfo()
                            With fp
                                .UserID = permissionUserId
                                .FolderID = folder.FolderID
                                .FolderPath = folder.FolderPath
                                .PermissionCode = "EDIT"
                                .PermissionKey = "EDIT"
                            End With
                            fpc.AddFolderPermission(fp)
                        End If
                    End If
                End If
            End If
        End Sub


#End Region


#Region "Private Methods"


        Private Sub EncryptInternal(ByVal newPayLoad As Byte(), ByVal objSalt As Byte())
            SyncLock Me
                Me._PayLoad = newPayLoad
                Me._SaltBytes = objSalt
                Me.State = (Me.State Or SmartFileState.Encrypted)
            End SyncLock
        End Sub

        Private Sub DecryptInternal(newPayload As Byte())
            SyncLock Me
                Me._PayLoad = newPayload
                Me.State = (Me.State And Not SmartFileState.Encrypted)
                Me._SaltBytes = {}
            End SyncLock
        End Sub


#End Region

    End Class


    <Serializable()>
    Public Class SmartFile(Of T)
        Inherits SmartFile

        Private _TypedValue As T

        Public Sub New()
        End Sub


        Public Sub New(key As EntityKey, value As T, settings As SmartFileInfo)
            MyBase.New(key, settings.Encryption)
            Me.TypedValue = value
            Me.Wrap(settings)
        End Sub

        <XmlIgnore()>
        <ExtendedCategory("Value")>
        Public Property ShowValue As Boolean

        <ExtendedCategory("Value")>
        <XmlIgnore()>
        <ConditionalVisible("ShowValue", False, True)>
        Public Property TypedValue As T
            Get
                If _TypedValue Is Nothing Then
                    Me.UnWrap()
                    If Not Me.PayLoad.Length = 0 Then

                        _TypedValue = Aricie.Services.ReflectionHelper.Deserialize(Of T)(Me.PayLoad)
                    End If
                End If
                Return _TypedValue
            End Get
            Set(value As T)
                _TypedValue = value
                Me.UpdatePayload()
            End Set
        End Property

        <ExtendedCategory("Value")>
        <ActionButton(IconName.Refresh, IconOptions.Normal)>
        Public Sub UpdatePayload(ape As AriciePropertyEditorControl)
            Me.UpdatePayload()
            ape.ItemChanged = True
            Dim message As String = Localization.GetString("SmartFileDecompressed.Message", ape.LocalResourceFile)
            ape.DisplayMessage(message, ModuleMessage.ModuleMessageType.GreenSuccess)
        End Sub

        Public Sub UpdatePayload()
            If _TypedValue IsNot Nothing Then
                PayLoad = ReflectionHelper.SerializeToBytes(Me._TypedValue, True)
            Else
                Me.PayLoad = {}
            End If
        End Sub

    End Class

End Namespace