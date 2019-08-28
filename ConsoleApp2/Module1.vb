Module Module1

    Private sessionID$ = ""
    Private sdkID$ = ""
    Private webURL$

    Sub Main()

        ActivateSession()
        TestCreateAndLogin()

    End Sub

    Private Function SearchInLDAP(ByVal directoryName As String, ByVal username As String) As CxPortal.CxDomainUser
        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()

        Dim resp As CxPortal.CxWSResponseDomainUserList

        resp = CxProxy.GetAllUsersFromUserDirectory(sessionID, directoryName, username, CxPortal.CxWSSearchPatternOption.Contains)
        If resp.IsSuccesfull Then
            Console.WriteLine("Get All users from user directory successful")

            For Each user As CxPortal.CxDomainUser In resp.UserList
                Console.WriteLine("username: {0}", user.Username)
                Console.WriteLine("UPN: {0}", user.UPN)
                Console.WriteLine("FirstName: {0}", user.FirstName)
                Console.WriteLine("LastName: {0}", user.LastName)
                Console.WriteLine("Email: {0}", user.Email)

                Return user

            Next

        End If

        Return Nothing
    End Function

    Private Sub TestCreateAndLogin()

        Dim username As String = "MyActiveDirectory\psmith"
        Dim password As String = "Cx!123456"

        'Dim allGroups As List(Of CxSDKns.Group)
        Dim teamNames As String = "CxServer\SP\APAC\DevOps"

        Dim ldapuserPair As String() = username.Split(New Char() {"\"c, " "c}, StringSplitOptions.RemoveEmptyEntries)
        Dim user As CxPortal.CxDomainUser = SearchInLDAP(ldapuserPair(0), ldapuserPair(1))

        If user Is Nothing Then
            Console.WriteLine("Failed")
            Return
        End If

        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()

        Dim resP As CxPortal.CxWSBasicRepsonse

        Dim userData As CxPortal.UserData = New CxPortal.UserData

        With userData
            .UserName = username

            Dim rolE As New CxPortal.Role
            rolE.ID = 0 'Scanner
            .RoleData = rolE

            .FirstName = user.FirstName
            .LastName = user.LastName
            .Email = user.Email

            Call CXgetGroups(allGroups)
            'Call CXgetUsers(allUsers)

            Dim GL As List(Of CxPortal.Group) = getGroupList(teamNames)
            .GroupList = GL.ToArray

            .Password = ""

            .DateCreated = Now
            .LastLoginDate = Now
            .LimitAccessByIPAddress = False

            .IsActive = True
            .UserPreferedLanguageLCID = 1033
            .willExpireAfterDays = 365

        End With

        resP = CxProxy.AddNewUser(sessionID, userData, CxPortal.CxUserTypes.LDAP)

        If resP.IsSuccesfull Then

            Console.WriteLine("Successful User Creation")

            Dim creD = New CxPortal.Credentials()
            creD.User = username
            creD.Pass = password

            Dim lResp As CxPortal.CxWSResponseLoginData
            lResp = CxProxy.Login(creD, 1033)

            If lResp.IsSuccesfull Then
                Console.WriteLine("Login successful")
            Else
                Console.WriteLine("Login failed")
            End If

        Else
            Console.WriteLine("Failed to create: {0}", resP.ErrorMessage)

        End If

    End Sub

    Private Sub Test()

        Dim CxSDKProxy = New CxSDKns.CxSDKWebServiceSoapClient
        Dim resP As CxSDKns.CxWSResponseGroupList

        resP = CxSDKProxy.GetAssociatedGroupsList(sessionID)
        If resP.IsSuccesfull = False Then
            Console.WriteLine("Error: Could not load Team Data - " + resP.ErrorMessage)
            Exit Sub
        Else
            For Each G As CxSDKns.Group In resP.GroupList
                'allGroups.Add(G)
                Console.WriteLine("*********************")
                Console.WriteLine("Groupname: {0}", G.GroupName)
                Console.WriteLine("Guid: {0}", G.Guid)
                Console.WriteLine("ID: {0}", G.ID)
                Console.WriteLine("Path: {0}", G.Path)
                Console.WriteLine("FullPath: {0}", G.FullPath)
                Console.WriteLine("Type: {0}", G.Type)
                Console.WriteLine("*********************")
            Next
        End If

    End Sub

    Public Sub CXgetUsers(ByRef allUsers As CxPortal.CxWSResponseUserData, Optional ByVal forceREFRESH As Boolean = False)
        Static alreadyGotUsers As Boolean = False

        If alreadyGotUsers = True And forceREFRESH = False Then
            Console.WriteLine("Already loaded users")
            Exit Sub
        End If

        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()
        Dim CxSDKProxy = New CxSDKns.CxSDKWebServiceSoapClient()

        allUsers = New CxPortal.CxWSResponseUserData

        Console.WriteLine("Retrieving list of users..")
        allUsers = CxProxy.GetAllUsers(sessionID)

        If allUsers.IsSuccesfull = False Then
            Console.WriteLine("ERROR: Could not pull Users - " + allUsers.ErrorMessage)
            Exit Sub
        End If

        Console.WriteLine(Trim(Str(allUsers.UserDataList.Count)) + " users loaded")
        alreadyGotUsers = True

    End Sub

    Private Sub MySub()

        'CXcreateHierarchy("cxserver\sp\apac\devops\team2")

        Dim ldapserverid As Integer = CXGetConfiguredLdapServerId("MyActiveDirectory")
        If ldapserverid <> -1 Then
            Console.WriteLine("ldapserverid {0}", ldapserverid)
        End If

        Dim ldapgrouppair1 As CxPortal.CxWSLdapGroup = CXGetLdapServerGroups(ldapserverid, "APAC-RND-Scanner-All")
        Dim ldapgrouppair2 As CxPortal.CxWSLdapGroup = CXGetLdapServerGroups(ldapserverid, "APAC-RND-Reviewer-All")

        Dim ldapgroups(1) As CxPortal.CxWSLdapGroupMapping
        ldapgroups(0) = New CxPortal.CxWSLdapGroupMapping()
        With ldapgroups(0)
            .LdapServerId = ldapserverid
            .LdapGroup = ldapgrouppair1
        End With
        ldapgroups(1) = New CxPortal.CxWSLdapGroupMapping()
        With ldapgroups(1)
            .LdapServerId = ldapserverid
            .LdapGroup = ldapgrouppair2
        End With

        CXcreateHierarchy("CxServer\SP\APAC\RND\Team3", ldapgroups)

        'Dim node As CxPortal.HierarchyGroupNode = searchWithinHierarchyTree(CXgetHierarchyGroupTree(), "CxServer\SP\APAC\RND\Team3", CxPortal.GroupType.Team)
        'printGroupNode(node)


    End Sub

    Private Function getGroupList(teamNames$) As List(Of CxPortal.Group)
        getGroupList = New List(Of CxPortal.Group)

        Dim tName() As String = Split(teamNames, ",")

        Dim tCtr As Integer
        For tCtr = 0 To UBound(tName)
            Dim G As New CxPortal.Group
            G.Guid = getGUIDofTEAM(LTrim(tName(tCtr)))
            Console.WriteLine("CONSOLE: team->" + tName(tCtr))
            Console.WriteLine("CONSOLE: team->" + printGroup("group", G))
            getGroupList.Add(G)
        Next

    End Function

    Private allUsers As CxPortal.CxWSResponseUserData
    Private allGroups As List(Of CxSDKns.Group)

    Public Sub CXgetGroups(ByRef allGroups As List(Of CxSDKns.Group), Optional ByVal forceREFRESH As Boolean = False)
        Static alreadyGotGroups As Boolean = False

        If alreadyGotGroups = True And forceREFRESH = False Then
            Console.WriteLine("Already loaded projects")
            Exit Sub
        End If

        Dim CxSDKProxy = New CxSDKns.CxSDKWebServiceSoapClient

        Console.WriteLine("Retrieving list Of Teams/Groups..")
        allGroups = New List(Of CxSDKns.Group)

        Dim resP As CxSDKns.CxWSResponseGroupList

        resP = CxSDKProxy.GetAssociatedGroupsList(sessionID)
        If resP.IsSuccesfull = False Then
            Console.WriteLine("Error: Could not load Team Data - " + resP.ErrorMessage)
            Exit Sub
        Else
            For Each G In resP.GroupList
                allGroups.Add(G)
            Next
        End If

        Console.WriteLine(Trim(Str(allGroups.Count)) + " groups loaded")

        alreadyGotGroups = True
        '        For Each G In allGroups
        '        TextBox1.Text += G.Team.GroupName + " - " + G.Team.ID + " - " + G.Team.Guid + vbCrLf
        '        Next

        '        addLOG(Trim(Str(allGroups.Count)) + " groups loaded")




    End Sub

    Private Function getGUIDofTEAM(teamName$, Optional ByVal teamNameOnly As Boolean = False) As String
        getGUIDofTEAM = ""
        Dim a$ = ""
        Dim b$ = ""
        b$ = LCase(teamName)
        If teamNameOnly Then b = LCase(stripToFilename(b))
        For Each G In allGroups
            a$ = LCase(G.GroupName)
            If teamNameOnly Then
                a = LCase(stripToFilename(a))
            End If
            If a = b Then
                If IsNothing(G.Guid) = True Then Return G.ID Else Return G.Guid
                Exit Function
            End If
        Next
    End Function

    Public Function stripToFilename(ByVal fileN$) As String
        'C:\Program Files\Checkmarx\Checkmarx Jobs Manager\Results\WebGoat.NET.Default 2014-10.9.2016-19.59.35.pdf
        stripToFilename = ""

        Do Until InStr(fileN, "\") = 0
            fileN = Mid(fileN, InStr(fileN, "\") + 1)
        Loop

        stripToFilename = fileN

    End Function

    Private Function CXGetLdapServerGroups(ByVal ldapServerId As Integer, ByVal groupName As String) As CxPortal.CxWSLdapGroup
        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()

        Dim resP As CxPortal.CxWSResponseLDAPServerGroups
        resP = CxProxy.GetLdapServerGroups(sessionID, ldapServerId, groupName, CxPortal.CxWSSearchPatternOption.Contains)

        If resP.IsSuccesfull Then

            Console.WriteLine("groups len:", resP.Groups.Length)

            'For Each group As CxPortal.CxWSLdapGroup In resP.Groups
            'Console.WriteLine("Group Name: {0}", group.Name)
            'Console.WriteLine("Group Dn: {0}", group.DN)
            'Next

            If resP.Groups.Length <> 0 Then
                Return resP.Groups(0)

                'Dim ldapGroupPair As CxPortal.CxWSLdapGroup = New CxPortal.CxWSLdapGroup
                'With ldapGroupPair
                '.Name = group[0].Name
                '.DN = "CN=APAC-RND-Scanner-All,OU=APAC,OU=Sites,DC=mycheckmarx,DC=com"
                'End With

            End If

        Else
            Console.WriteLine("Err: " + resP.ErrorMessage)
        End If
        Return Nothing
    End Function


    Private Function CXGetConfiguredLdapServerId(ByVal ldapName As String) As Integer

        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()

        Dim ldapPairs As CxPortal.CxWSResponseIdNamePairList
        ldapPairs = CxProxy.GetConfiguredLdapServerNames(sessionID, False)

        If ldapPairs.IsSuccesfull Then
            For Each item As CxPortal.CxWSIdNamePair In ldapPairs.Items()
                If item.Name.Equals(ldapName) Then
                    Return item.Id
                End If
            Next

        End If
        Return -1
    End Function


    Public Function CXsetTeamLDAPGroupsMapping(ByVal teamId As String, ByVal ldapGroups As CxPortal.CxWSLdapGroupMapping())

        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()

        Console.WriteLine("Mapping: {0}", ldapGroups)


        Dim resP As CxPortal.CxWSBasicRepsonse
        resP = CxProxy.SetTeamLdapGroupsMapping(sessionID, teamId, ldapGroups)

        If resP.IsSuccesfull = True Then
            Console.WriteLine("CONSOLE: SetTeamLdapGroupsMapping rv: " + resP.IsSuccesfull.ToString)
        Else
            Console.WriteLine("CONSOLE: SetTeamLdapGroupsMapping rv: " + resP.ErrorMessage)
        End If

    End Function


    Public Function CXcreateHierarchy(ByVal toHavePath As String, ByVal ldapGroups As CxPortal.CxWSLdapGroupMapping())
        transverseHierarchy(CXgetHierarchyGroupTree(), toHavePath, ldapGroups)
    End Function

    Private Function transverseHierarchy(ByVal hierarchy As CxPortal.HierarchyGroupNode(), ByVal toHavePath As String, ByVal ldapGroups As CxPortal.CxWSLdapGroupMapping()) As Boolean
        For Each node As CxPortal.HierarchyGroupNode In hierarchy
            printGroupNode(node)
            Dim path As String = getNodePath(node)

            If (path.Equals(toHavePath)) Then 'Path exists, no need to search further
                Console.WriteLine("Path Exists: {0}", toHavePath)
                Return True

            ElseIf (toHavePath.StartsWith(path)) Then
                Console.WriteLine("Starting path exists for {0}", toHavePath)
                Console.WriteLine("#Child nodes: {0}", node.Childs.Length)

                Dim isExist As Boolean = transverseHierarchy(node.Childs, toHavePath, ldapGroups)
                Console.WriteLine("transverse inner: {0}", isExist)

                If (Not isExist) Then 'No child, create company/teams

                    Console.WriteLine("No more child to transverse, start creating path")
                    Dim newGroups As String()
                    newGroups = getNewPathElements(path, toHavePath)
                    Console.WriteLine("Missing no. of groups: {0}", newGroups.Length)

                    Dim newParentGroup As CxPortal.Group = node
                    For Each newGroup As String In newGroups
                        newParentGroup = createNewCompanyOrTeam(newParentGroup, newGroup)
                    Next

                    ' Add ldap groups if specified
                    If ldapGroups IsNot Nothing AndAlso ldapGroups.Count <> 0 Then
                        CXsetTeamLDAPGroupsMapping(newParentGroup.Guid, ldapGroups)
                    End If

                End If

                Return True

            End If
        Next
    End Function

    Private Function getNewPathElements(ByVal path As String, ByVal toHavePath As String) As String()
        Console.WriteLine("***********************************")
        Console.WriteLine("extract missing groups {0} from {1}", toHavePath, path)
        getNewPathElements = toHavePath.Substring(path.Length).Split(New Char() {"\"c}, StringSplitOptions.RemoveEmptyEntries)
    End Function

    Private Function getNodePath(ByVal node As CxPortal.HierarchyGroupNode) As String
        ' Server is always empty, use groupname instead
        If (String.IsNullOrEmpty(node.FullPath)) Then
            getNodePath = node.GroupName
        Else
            getNodePath = node.FullPath
        End If
    End Function

    Private Function printGroupNode(ByVal node As CxPortal.HierarchyGroupNode)
        Console.WriteLine("***********************************")
        Console.WriteLine("GroupId: " + node.ID)
        Console.WriteLine("GUID: " + node.Guid)
        Console.WriteLine("GroupName: " + node.GroupName)
        Console.WriteLine("Path: " + node.Path)
        Console.WriteLine("FullPath: " + node.FullPath)
        Console.WriteLine("Childs Len: {0}", node.Childs.Length)
        Console.WriteLine("GroupType: {0}", node.Type.ToString)
        Console.WriteLine("***********************************")
    End Function

    Private Function CXgetHierarchyGroupTree() As CxPortal.HierarchyGroupNode()
        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()
        Dim groupNodes As CxPortal.CxWSResponseHierarchyGroupNodes
        groupNodes = CxProxy.GetHierarchyGroupTree(sessionID)

        If groupNodes.IsSuccesfull Then
            Return groupNodes.HierarchyGroupNodes
        End If

    End Function

    Private Function searchWithinHierarchyTree(ByVal hierarchy As CxPortal.HierarchyGroupNode(), ByVal searchPath As String, ByVal groupType As CxPortal.GroupType) As CxPortal.HierarchyGroupNode

        Console.WriteLine("searchWithinHierarchyTree for {0}", searchPath)

        ' Get ancestry tree, traversed via node.FullPath
        For Each node As CxPortal.HierarchyGroupNode In hierarchy
            printGroupNode(node)
            Dim rtnpath As String = getNodePath(node)

            If rtnpath.Equals(searchPath) AndAlso node.Type.Equals(groupType) Then
                Console.WriteLine("Found")
                Return node

            ElseIf (searchPath.StartsWith(rtnpath)) Then
                Console.WriteLine("Starting path exists for {0}", searchPath)
                Console.WriteLine("#Child nodes: {0}", node.Childs.Length)

                Return searchWithinHierarchyTree(node.Childs, searchPath, groupType)

            End If

        Next
        Return Nothing
    End Function


    Private Function createNewCompanyOrTeam(ByVal node As CxPortal.Group, ByVal newCompanyOrTeamName As String) As CxPortal.Group

        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()
        Dim LD(0) As CxPortal.CxWSLdapGroupMapping

        Dim resP As CxPortal.CxWSBasicRepsonse

        'printGroup("parent_node", node)

        ' Assumption that Service Provider will always exist
        If (node.Type.Equals(CxPortal.GroupType.SP)) Then ' Parent is Service provider, create company

            resP = CxProxy.CreateNewCompany(sessionID, node.Guid, newCompanyOrTeamName, 0, 0, 0, False, LD)
            Console.WriteLine("createNewCompany rv: {0}", resP.IsSuccesfull)

            If resP.IsSuccesfull Then

                Dim resG As CxPortal.CxWSResponseTeamData
                ' search under SP
                resG = CxProxy.GetServiceProviderCompanies(sessionID, node.Guid)

                If resG.IsSuccesfull Then
                    For Each item As CxPortal.TeamData In resG.TeamDataList()

                        printGroup("company_node", item.Company)

                        ' No need to verified service provider as already under SP
                        If item.Company.GroupName.Equals(newCompanyOrTeamName) Then

                            If String.IsNullOrEmpty(item.Company.FullPath) Then
                                item.Company.FullPath = node.FullPath + "\" + newCompanyOrTeamName
                            End If

                            Return item.Company
                        End If
                    Next
                End If
            End If

        Else 'Parent is Company/Team, create team

            resP = CxProxy.CreateNewTeam(sessionID, node.Guid, newCompanyOrTeamName, LD)
            Console.WriteLine("createNewTeam rv: {0}", resP.IsSuccesfull)

            If resP.IsSuccesfull = True Then

                Dim team As CxPortal.HierarchyGroupNode = searchWithinHierarchyTree(CXgetHierarchyGroupTree(), node.FullPath + "\" + newCompanyOrTeamName, CxPortal.GroupType.Team)

                If team IsNot Nothing Then
                    Return team
                Else
                    Console.WriteLine("problem finding team")
                End If


            End If

            End If
        Return Nothing
    End Function



    Private Function printGroup(ByVal label As String, ByVal node As CxPortal.Group)
        Console.WriteLine("************ " + label + " *************")
        Console.WriteLine("GroupId: {0}", node.ID)
        Console.WriteLine("GUID: {0}", node.Guid)
        Console.WriteLine("GroupName: {0}", node.GroupName)
        Console.WriteLine("Path: {0}", node.Path)
        Console.WriteLine("FullPath: {0}", node.FullPath)
        Console.WriteLine("GroupType: {0}", node.Type.ToString)
        Console.WriteLine("***********************************")

    End Function

    Public Function ActivateSession() As String
        ActivateSession = "True"

        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()
        Dim CxSDKProxy = New CxSDKns.CxSDKWebServiceSoapClient()

        Dim username As String
        Dim password As String

        username = "administrator"
        password = "Cx!123456"

        Dim creD = New CxPortal.Credentials()
        creD.User = username
        creD.Pass = password

        Dim cred2 = New CxSDKns.Credentials()
        cred2.User = username
        cred2.Pass = password

        Dim lResp As CxPortal.CxWSResponseLoginData
        Dim lResp2 As CxSDKns.CxWSResponseLoginData

        lResp = CxProxy.Login(creD, 1033)
        lResp2 = CxSDKProxy.Login(cred2, 1033)


        Console.WriteLine("Private SOAP:" + lResp.IsSuccesfull.ToString + ":" + lResp.ErrorMessage)
        Console.WriteLine("Public SOAP:" + lResp2.IsSuccesfull.ToString + ":" + lResp.ErrorMessage)

        If lResp.IsSuccesfull = True Then sessionID = lResp.SessionId

        If lResp2.IsSuccesfull = True Then sdkID = lResp2.SessionId

        creD.Pass = ""
        cred2.Pass = ""

        If sessionID = "" Or sdkID = "" Then
            Console.WriteLine("Cannot obtain sessionID! API calls will fail.")
            ActivateSession = "ERROR: Portal- " + lResp.ErrorMessage + " / SDK - " + lResp2.ErrorMessage
        Else
            Console.WriteLine("Obtained Session IDs For PORTAL And SDK APIs")
            ActivateSession = "True"
        End If

        Dim PP As New CxPortal.CxWsResponseSystemSettings
        PP = CxProxy.GetSystemSettings(sessionID)

        If PP.IsSuccesfull = True Then
            webURL = PP.SysSettings.WebServer
            webURL = Replace(webURL, "https://", "")
            webURL = Replace(webURL, "http://", "")
        Else
            webURL = "Limited Access"
        End If
    End Function


End Module
