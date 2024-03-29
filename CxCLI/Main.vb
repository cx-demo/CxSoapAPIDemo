﻿Imports System.Net
Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Web
Imports System.ServiceModel
Imports System.Diagnostics
'Imports Microsoft.Office.Interop
'Imports System.Net.Mail
Imports System.Threading
Imports System.IO.Compression
Imports System.Linq
Imports System.Collections.Generic
Imports System.Xml.Linq
Imports System.ComponentModel

Module Module1

    Private sessionID$
    Private sdkID$
    Private allUsers As CxPortal.CxWSResponseUserData
    Private allScanS As CxSDKns.CxWSResponseScansDisplayData


    ' Private showingADdropdown As Boolean

    Private allProjScans As CxPortal.CxWSResponseProjectScannedDisplayData

    Private usersDisplayData As AdvancedList(Of userDisplayRows)
    'Private projDisplayData As AdvancedList(Of projDisplayProps)
    'Private allProjects As List(Of CxPortal.ProjectDisplayData)

    'Private customFields As CxPortal.CxWSCustomField()
    'Private custFieldValues As Dictionary(Of String, String)

    'Public publicProjIDs As Collection

    '    Private resultStates As CxPortal.ResultState()

    Private allTeamProperties As AdvancedList(Of teamProps)

    Private currDomainList As List(Of CxPortal.CxDomainUser)
    'Private currReportRequests As List(Of getReportArgs)

    'Private allProjConfigs As List(Of CxPortal.ProjectConfiguration)


    ' Private WithEvents projScans As batchOfScans
    ' Private vulnsWaitingToDerive As List(Of Long)

    Private CXldapServers As List(Of CxPortal.CxWSLdapServerConfiguration)

    Private isLDAPactive As Boolean
    Private allGroups As List(Of CxSDKns.Group)
    Private allRoles As List(Of CxPortal.Role)
    ' Private projectsWithScans As CxPortal.CxWSResponseProjectsScansList
    Private allPresets As CxPortal.CxWSResponsePresetList
    ' Private allStates As List(Of CxPortal.ResultState)

    'Private exeAction As String
    'Private exeActionTimeoutMins As Integer
    'Private dashResults As dashInfo
    'Private buildingDash As Boolean
    Private WithEvents CxWrap As CxWrapper
    'Public guiActive As Boolean

    'Private backCompleteUsers As Boolean
    'Private backCompleteProj As Boolean
    'Private backCompleteGroups As Boolean
    'Private backCompleteScans As Boolean


    '    Private scanDetailWaiting As Boolean

    'Private usersBeingAdded As Collection
    'Private usersBeingEdited As Collection
    'Private groupsBeingAdded As Collection

    'Private runningUserEdits As backgroundUserArgs

    'Private tlNumBulkCreates As Integer

    Private exeAction$

    Delegate Sub StringArgReturningVoidDelegate([text] As String)
    Public loggingEnabled As Boolean

    Sub Main()

        Dim filenameArg$ = ""
        exeAction = ""

        loggingEnabled = False

        Dim doLogging$ = argPROP("loggingenabled")
        If LCase(doLogging) = "true" Then
            loggingEnabled = True
            addLOG("------------------------------------------")
        End If

        '        startSession()
        '        Call CxWrap.CXgetGroups(allGroups)
        '        Call CxWrap.CXgetUsers(allUsers)
        '        Call addUser("application", "testcxcli33", "Scanner", "TestCxCLI22", "TestLast", "CxServer\SP\Company\Users", "test333@cx.com", "", "", "", "", "", "", "", "", "Password12345!")
        '        Exit Sub



        For Each Arg In My.Application.CommandLineArgs
            Call addLOG("Argument: " & Arg)
            'If InStr(Arg, "mailto") Then exeAction = "Send File to Specified Recipient"
            'If InStr(Arg, "mailcustomfield") Then exeAction = "Send Mail Custom Field"
            'If InStr(Arg, "importxls") Then exeAction = "Import Users from Spreadsheet"
            'If InStr(Arg, "userrptxls") Then exeAction = "Report Users to Spreadsheet"
            'If InStr(Arg, "unzip") Then exeAction = "Unzip File"
            'If InStr(Arg, "cleanzip") Then exeAction = "Clean Zip File"
            'If InStr(Arg, "zipsource") Then exeAction = "Zip Source"
            'If InStr(Arg, "scanfolderofzips") Then exeAction = "Scan Folder of Zips"
            If InStr(Arg, "disableusers") Then exeAction = "Disable Users from File"
            If InStr(Arg, "deleteusers") Then exeAction = "Delete Users from File"
            If InStr(Arg, "edituser") Then exeAction = "CLI Edit User"
            If InStr(Arg, "adduser") Then exeAction = "Add User"
            If InStr(Arg, "enableusers") Then exeAction = "Enable Users from File"


            '6/26
            If InStr(Arg, "showteams") Then exeAction = "Show all Teams"
            If InStr(Arg, "swapteams") Then exeAction = "Move Users to a new Team"
            If InStr(Arg, "showusers") Then exeAction = "Show Users"



            If InStr(Arg, "userexpire") Then exeAction = "Show Expiring Users"
            If InStr(Arg, "setuserexpire") Then exeAction = "Set New Expiration for Users"

            If InStr(Arg, "userexist") Then exeAction = "Determine if User Exists"


            If InStr(Arg, "encrypt") Then exeAction = "Encrypt Text"
            If InStr(Arg, "help") Then exeAction = "help"

            If InStr(Arg, "getpresets") Then exeAction = "Get List of Presets"
            If InStr(Arg, "getpresetdef") Then exeAction = "Get Preset Details"
            If InStr(Arg, "addpreset") Then exeAction = "Add Preset"

            If InStr(Arg, "cancelscans") Then exeAction = "Cancel Scans"
            If InStr(Arg, "showscans") Then exeAction = "Show Scans"
            If InStr(Arg, "numscans") Then exeAction = "Number of Scans"

            'If InStr(Arg, "filediff") Then exeAction = "Perform DIFF analysis of File Listings"
            'If InStr(Arg, "addsamlxls") Then exeAction = "Add SAML users from XLS"

            'Pedric: Added Aug 2019
            If InStr(Arg, "login") Then exeAction = "User Login"
            If InStr(Arg, "createhierarchy") Then exeAction = "Create Hierarchy"

        Next


        If Len(exeAction) Then addLOG("CLI Action - " + exeAction)

        Select Case exeAction

    '        Case "Send Mail Custom Field"
            '            Call emailFromCustomField(filenameArg, argPROP("mailcustomfield"))
            '            exeActionTimeoutMins = 5
            '            End
            '         Case "Report Users to Spreadsheet"
            '             Call CxWrap.CXgetGroups(allGroups)
            '             Call CxWrap.CXgetUsers(allUsers)
            '             Call usersReport()
            '             End
            '         Case "Add SAML users from XLS"
            '             Call CxWrap.CXgetGroups(allGroups)
            '             Call CxWrap.CXgetUsers(allUsers)
            '             Call importSAMLfromXLS(filenameArg)
            '             End

            '          Case "Import Users from Spreadsheet"
            '              Call CxWrap.CXgetGroups(allGroups)
            '              Call importXLS(filenameArg)
            '              End
            '          Case "Send File to Specified Recipient"
            '              Call sendEmailAttachment(filenameArg, argPROP("mailto"), "Full PDF: " + stripToFilename(filenameArg), "Sending Full PDF: " + stripToFilename(filenameArg))
            '              End

            Case "Encrypt Text"
                Dim pW3DES As New Simple3Des("2&#263gdjSiUEYkadhEII276#*763298")
                addLOG("CONSOLE:encrypted_text=" + pW3DES.Encode(argPROP("text", True)))
                pW3DES = Nothing
                End


            Case "Cancel Scans"
                addLOG("CONSOLE:Current Scans")
                Call cancelScans
                End

            Case "Show Scans", "Number of Scans"
                Dim numOnly As Boolean = False
                If exeAction = "Number of Scans" Then numOnly = True

                Dim minS As Long
                If Val(argPROP("mins")) > 0 Then minS = Val(argPROP("mins"))
                Dim echoS = ""
                echoS = "CONSOLE:Show Scans of TYPE:" + argPROP("type")
                If minS Then echoS += " since " + DateAdd(DateInterval.Minute, Val(argPROP("mins")) * -1, Date.Now)
                addLOG(echoS)
                Call showScans(argPROP("type"), argPROP("mins"), numOnly)
                End

            Case "Get List of Presets"
                allPresets = New CxPortal.CxWSResponsePresetList
                startSession()
                Call CxWrap.CXgetPresetList(allPresets)

                For Each P In allPresets.PresetList
                    addLOG("CONSOLE:" + P.ID.ToString + ":" + P.PresetName)
                Next
                GC.Collect()

                End


            Case "Add User"
                startSession()
                Call CxWrap.CXgetGroups(allGroups)
                Call CxWrap.CXgetUsers(allUsers)
                Call addUser(argPROP("usertype", True), argPROP("username", True), argPROP("role", True), argPROP("firstname", True), argPROP("lastname", True), argPROP("team", True), argPROP("email", True), argPROP("jobtitle", True), argPROP("country", True), argPROP("phone", True), argPROP("cellphone", True), argPROP("langlcid", True), argPROP("audituser", True), argPROP("activeuser", True), argPROP("expiredays", True), argPROP("password", True))
                End


            Case "Add Preset"
                allPresets = New CxPortal.CxWSResponsePresetList
                startSession()
                ' Call CxWrap.CXgetPresetList(allPresets)  'not necessary as import will load
                Dim PD As New CxPortal.CxPresetDetails
                Dim allQ As Object
                PD.name = argPROP("name", True)
                allQ = Split(argPROP("queries"), ",")
                Dim numQ As Integer = 0
                Dim listOfNums(UBound(allQ)) As Long

                For Each Q In allQ
                    listOfNums(numQ) = Q
                    numQ += 1
                Next
                PD.queryIds = listOfNums

                Dim resP$
                resP = CxWrap.importPreset(PD)
                If resP = "True" Then
                    addLOG("CONSOLE:Preset added:" + PD.id.ToString + " " + PD.name)
                Else
                    addLOG("CONSOLE:ERROR:" + resP)
                End If

            Case "Get Preset Details"
                allPresets = New CxPortal.CxWSResponsePresetList
                startSession()
                Call CxWrap.CXgetPresetList(allPresets)

                Dim getPid As Long
                getPid = Val(argPROP("id", False))

                Dim PD As New CxPortal.CxPresetDetails

                For Each P In allPresets.PresetList
                    If P.ID = getPid Then
                        Dim resP$
                        resP = CxWrap.getPresetDetails(getPid, PD)
                        If resP <> "True" Then
                            addLOG("CONSOLE:ERROR - " + resP)
                            End
                        End If

                        addLOG("CONSOLE:NAME:" + PD.name)
                        addLOG("CONSOLE:ID:" + PD.id.ToString)
                        addLOG("CONSOLE:OWNER:" + PD.owner)
                        addLOG("CONSOLE:OWNING_TEAM:" + PD.owningteam)
                        addLOG("CONSOLE:IS_DUPLICATE:" + PD.IsDuplicate.ToString)
                        addLOG("CONSOLE:IS_PUBLIC:" + PD.isPublic.ToString)
                        addLOG("CONSOLE:UPDATEABLE:" + PD.isUserAllowToUpdate.ToString)
                        addLOG("CONSOLE:DELETABLE:" + PD.isUserAllowToDelete.ToString)

                        Dim allQ$ = ""
                        For Each Q In PD.queryIds
                            allQ += Q.ToString + ","
                        Next
                        allQ = Mid(allQ, 1, Len(allQ) - 1)

                        addLOG("CONSOLE:QUERIES:" + allQ)
                    End If

                Next
                GC.Collect()
                End

            Case "Help", "help", "-h", "/h"
                dumpInstructs()
                End


            Case "Determine if User Exists"
                startSession()
                Call CxWrap.CXgetUsers(allUsers)
                Dim C As New CLIArgs
                C.matchOn = argPROP("match")
                C.uData = argPROP("user")
                Dim U As CxPortal.UserData

                U = returnUserUsingMatch(C.matchOn, C.uData)
                If U.UserName = "NOTFOUND" Then
                    addLOG("CONSOLE:FALSE")
                Else
                    addLOG("CONSOLE:TRUE")
                End If
                GC.Collect()

                End

            Case "User Login"
                'startSession()
                CxWrap = New CxWrapper
                Call CxWrap.CXlogin(argPROP("user", True), argPROP("pass", True))
                GC.Collect()
                End

            Case "Create Hierarchy"
                startSession()

                Dim ldapserver As String = argPROP("ldapservername", True)
                Dim ldapgroupsArg As String = argPROP("ldapgroups", True)

                ' Get Configured LDAP Server ID
                Dim ldapserverid As Integer = CxWrap.CXGetConfiguredLdapServerId(ldapserver)
                If ldapserverid <> -1 Then
                    Console.WriteLine("ldapserverid {0}", ldapserverid)
                End If

                Dim ldapgroups As String() = ldapgroupsArg.Split(New Char() {","c, " "c}, StringSplitOptions.RemoveEmptyEntries)

                Dim ldapgroupmapping As New List(Of CxPortal.CxWSLdapGroupMapping)
                For Each ldapgroup As String In ldapgroups

                    Dim ldapgrouppair As CxPortal.CxWSLdapGroup = CxWrap.CXGetLdapServerGroups(ldapserverid, ldapgroup)
                    Dim ldapgroupMap As CxPortal.CxWSLdapGroupMapping = New CxPortal.CxWSLdapGroupMapping()
                    With ldapgroupMap
                        .LdapServerId = ldapserverid
                        .LdapGroup = ldapgrouppair
                    End With

                    ldapgroupmapping.Add(ldapgroupMap)
                Next


                Call CxWrap.CXcreateHierarchy(argPROP("hierarchy", True), ldapgroupmapping.ToArray())
                GC.Collect()
                End

            Case "Disable Users from File"
                startSession()
                Call CxWrap.CXgetUsers(allUsers)
                Call CXdisableUserFromFile(argPROP("file"), argPROP("match"))
                GC.Collect()

                End

            Case "Delete Users from File"
                startSession()
                Call CxWrap.CXgetUsers(allUsers)
                Call CXdisableUserFromFile(argPROP("file"), argPROP("match"), False, True)
                GC.Collect()

                End

            Case "Enable Users from File"
                startSession()
                Call CxWrap.CXgetUsers(allUsers)
                Call CXdisableUserFromFile(argPROP("file"), argPROP("match"), True)
                GC.Collect()

                End

            Case "Show Expiring Users"
                startSession()
                Call CxWrap.CXgetUsers(allUsers)
                Dim numDays As Integer = CInt(argPROP("days"))
                Dim numExpired As Integer = 0
                Dim a$ = ""

                For Each U In allUsers.UserDataList
                    numExpired = U.willExpireAfterDays
                    If numExpired <= 0 Then a$ = " *EXPIRED*" Else a$ = ""
                    If numExpired <= numDays Then addLOG("CONSOLE:" + U.UserName + " " + U.Email + " expires in " + U.willExpireAfterDays.ToString + " days" + a)
                Next

                GC.Collect()

                End

            Case "Set New Expiration for Users"
                startSession()
                Call CxWrap.CXgetUsers(allUsers)
                Dim numDays As Integer = CInt(argPROP("days"))
                Dim newDays As Long = DateDiff(DateInterval.Day, Today, CDate(argPROP("newdate")))
                addLOG("CONSOLE:For users expiring in <" + numDays.ToString + " days, set to " + newDays.ToString + " days (" + argPROP("newdate") + ")")
                Dim numExpired As Integer = 0

                For Each U In allUsers.UserDataList
                    numExpired = U.willExpireAfterDays
                    If numExpired <= numDays Then
                        U.willExpireAfterDays = newDays
                        addLOG("CONSOLE:" + U.UserName + " " + U.Email + " now expires in " + U.willExpireAfterDays.ToString + " days: " + CxWrap.CXeditUser(U))
                    End If
                Next

                GC.Collect()

                End


            Case "Show Users"
                startSession()

                Call CxWrap.CXgetGroups(allGroups)
                Call CxWrap.CXgetUsers(allUsers)

                Dim teamID$ = getGUIDofTEAM(argPROP("team", True))
                If LCase(argPROP("team")) = "all" Then teamID = "all"

                For Each G In allGroups
                    If G.ID <> teamID And teamID <> "all" Then GoTo nextGroup
                    For Each U In allUsers.UserDataList
                        For Each GG In U.GroupList
                            If GG.ID = G.ID Then
                                addLOG("CONSOLE:Team:" + G.GroupName + " GUID:" + G.ID + " USERNAME:" + U.UserName + " EMAIL:" + U.Email)

                            End If
                        Next
                    Next

nextGroup:
                Next



            Case "Show all Teams"
                startSession()

                Call CxWrap.CXgetGroups(allGroups)
                Call CxWrap.CXgetUsers(allUsers)

                Dim numU As Integer = 0

                For Each G In allGroups
                    numU = 0
                    For Each U In allUsers.UserDataList
                        For Each GG In U.GroupList
                            If GG.ID = G.ID Then numU += 1
                        Next
                    Next
                    addLOG("CONSOLE:Team:" + G.GroupName + " GUID:" + G.ID + " USERS:" + numU.ToString)
                Next

            Case "Move Users to a new Team"
                startSession()

                Call CxWrap.CXgetGroups(allGroups)
                Call CxWrap.CXgetUsers(allUsers)

                Dim fromTeam$ = argPROP("from", True)
                Dim toTeam$ = argPROP("to", True)

                Dim fromID$ = getGUIDofTEAM(fromTeam)

                addLOG("CONSOLE:Moving all users that are part of " + fromTeam + " into " + toTeam)
                Dim usersToNewGroup As New Collection

                For Each U In allUsers.UserDataList
                    '                    addLOG("CONSOLE:" + U.UserName + " # Teams:" + U.GroupList.Count.ToString)
                    Dim groupExist As Boolean = False

                    For Each GGG In U.GroupList
                        If LCase(GGG.ID) = LCase(fromID) Then
                            'addLOG("CONSOLE:User is part of " + fromTeam)
                            usersToNewGroup.Add(U.UserName)
                            groupExist = True
                        End If
                    Next

                    If groupExist Then
                        Dim C As New CLIArgs
                        C.matchOn = "username"
                        C.uData = U.UserName
                        C.newVal = toTeam
                        C.editCmd = "addgroup"
                        Call editUserCLI(C)
                    End If
                Next
                Call CxWrap.CXgetUsers(allUsers, True)

                For Each usrName In usersToNewGroup
                    Dim C As New CLIArgs
                    C.matchOn = "username"
                    C.uData = usrName
                    C.newVal = fromTeam
                    C.editCmd = "subtractgroup"
                    Call editUserCLI(C)

                Next

            Case "CLI Edit User"
                startSession()

                Call CxWrap.CXgetGroups(allGroups)
                Call CxWrap.CXgetUsers(allUsers)
                Dim C As New CLIArgs

                addLOG("CLI Edit User")
                C.matchOn = argPROP("match")
                C.uData = argPROP("user")

                addLOG("CLI EDIT USER" + vbCrLf + C.uData + " match on " + C.matchOn)

                If Len(argPROP("addgroup")) Then
                    C.newVal = argPROP("addgroup")
                    C.editCmd = "addgroup"
                    Call editUserCLI(C)
                End If

                If Len(argPROP("subtractgroup")) Then
                    C.newVal = argPROP("subtractgroup")
                    C.editCmd = "subtractgroup"
                    Call editUserCLI(C)
                End If

                If Len(argPROP("changerole")) Then
                    C.newVal = argPROP("changerole")
                    C.editCmd = "changerole"
                    Call editUserCLI(C)
                End If

                If Len(argPROP("status")) Then
                    C.editCmd = "status"
                    C.newVal = argPROP("status")
                    'either enable,disable,delete
                    Call editUserCLI(C)
                End If

                GC.Collect()

                End

                '     Case "Scan Folder of Zips"
                '         Call scanFolderOfZips(filenameArg, argPROP("team"), argPROP("preset"))
                '         End

                '          Case "Unzip File"
                '              Dim targetFolder$ = ""
                '              If My.Application.CommandLineArgs.Count > 2 Then
                '                  targetFolder = My.Application.CommandLineArgs(2)
                '              End If
                '              Call unzipFile(filenameArg, targetFolder)
                '              End
                '          Case "Clean Zip File"
                '              Dim targetFile$ = My.Application.CommandLineArgs(2)
                '              Call cleanZIP(filenameArg, targetFile)
                '              End
                '           Case "Zip Source"
                '               Dim sourceFolder$ = My.Application.CommandLineArgs(0)
                '               Call zipAllSRC(sourceFolder)
                '               End
                '            Case "Perform DIFF analysis of File Listings"
                '                Call performDIFF(argPROP("file1"), argPROP("file1type"), argPROP("file2"), argPROP("file2type"), filenameArg)
                '                End
        End Select

    End Sub


    Public Sub showScans(typeOfScan$, pastMins$, Optional ByVal numOnly As Boolean = False)
        startSession()

        Dim calcDate As Boolean = False
        Dim minS As Long
        If Len(pastMins) Then
            minS = Val(pastMins)
            calcDate = True
        End If

        Dim datesAfter As Date
        datesAfter = DateAdd(DateInterval.Minute, Val(argPROP("mins")) * -1, Date.Now)

        Dim tlScans As Long = 0

        Select Case typeOfScan

            Case "queued"
                Dim SS As CxPortal.CxWSResponseExtendedScanStatus()

                SS = CxWrap.getScansInQueue


                For Each S In SS
                    If CXconvertDTportal(S.TimeBeginWorking) >= datesAfter Or calcDate = False Then
                        tlScans += 1
                    End If
                Next

                addLOG("CONSOLE:TOTAL NUM SCANS:" + Trim(Str(tlScans)))

                If numOnly = True Then Exit Sub

                For Each S In SS
                    If CXconvertDTportal(S.TimeBeginWorking) >= datesAfter Or calcDate = False Then
                        addLOG("CONSOLE:STATUS:" + S.CurrentStatus.ToString + " STAGE:" + S.CurrentStage.ToString + " RUNID:" + S.RunId.ToString)
                    End If
                Next


            Case "completed"

                Dim eM$ = ""
                Call CxWrap.CXgetScans(allScanS,, eM)

                If Len(eM) Then
                    addLOG("CONSOLE:ERROR:" + eM)
                    Exit Sub
                End If

                Dim S As CxSDKns.ScanDisplayData

                For Each S In allScanS.ScanList
                    If CXconvertDT(S.FinishedDateTime) >= datesAfter Or calcDate = False Then
                        tlScans += 1
                    End If
                Next

                addLOG("CONSOLE:TOTAL NUM SCANS:" + Trim(Str(tlScans)))

                If numOnly = True Then Exit Sub

                For Each S In allScanS.ScanList
                    If CXconvertDT(S.FinishedDateTime) >= datesAfter Or calcDate = False Then
                        addLOG("CONSOLE:PROJECT:" + S.ProjectName + " SCANID:" + S.ScanID.ToString + " LOC:" + S.LOC.ToString + " COMPLETED:" + CXconvertDT(S.FinishedDateTime).ToString + " COMMENTS:" + S.Comments)
                    End If
                Next


            Case "failed"
                Dim allFailed As New CxPortal.CxWSResponseFailedScansDisplayData
                Dim eS$
                eS = CxWrap.CXgetFailedScans(allFailed)
                If eS <> "TRUE" Then
                    addLOG("CONSOLE:ERROR:" + eS)
                    Exit Sub
                End If

                Dim S As CxPortal.FailedScansDisplayData

                For Each S In allFailed.FailedScansList
                    If New Date(S.CreatedOn).ToString >= datesAfter Or calcDate = False Then
                        tlScans += 1
                    End If
                Next

                addLOG("CONSOLE:TOTAL NUM SCANS:" + Trim(Str(tlScans)))

                If numOnly = True Then Exit Sub

                For Each S In allFailed.FailedScansList
                    If New Date(S.CreatedOn).ToString >= datesAfter Or calcDate = False Then
                        addLOG("CONSOLE:PROJECT:" + S.ProjectName + " DETAILS:" + S.Details + " LOC:" + S.LOC.ToString + " COMPLETED:" + New Date(S.CreatedOn).ToString + " COMMENTS:" + S.Comments)
                    End If
                Next



        End Select


    End Sub

    Public Sub cancelScans()
        startSession()

        Dim SS As CxPortal.CxWSResponseExtendedScanStatus()

        SS = CxWrap.getScansInQueue

        addLOG("CONSOLE:CURRENT SCANS")
        For Each S In SS
            addLOG("CONSOLE:STATUS:" + S.CurrentStatus.ToString + " STAGE:" + S.CurrentStage.ToString + " RUNID:" + S.RunId.ToString)
        Next

        For Each S In SS
            If S.CurrentStatus.ToString = "Queued" Then addLOG("CONSOLE:CANCEL " + S.RunId.ToString + ":" + CxWrap.cancelScanID(S.RunId.ToString))
        Next

    End Sub
    Private Sub getLDAPservers()
        CXldapServers = New List(Of CxPortal.CxWSLdapServerConfiguration)
        Dim getLDAP$ = ""
        getLDAP = CxWrap.CxGetLdapServers(CXldapServers)

        addLOG("GETLDAP: " + getLDAP)

        If CXldapServers.Count Then
            '            allADusers = New List(Of CxPortal.CxDomainUser)
            isLDAPactive = True
            addLOG("LDAP configured @" + CXldapServers(0).Name + " -type " + Trim(Str(CXldapServers(0).DirectoryType)))
        Else
            addLOG("NO LDAP discovered")
            isLDAPactive = False
        End If
    End Sub

    Private Sub addLDAPUserFromCLI(U As CxPortal.UserData)
        Dim currDomainList As New List(Of CxPortal.CxDomainUser)
        currDomainList = CxWrap.CXgetLDAPUsers(CXldapServers(0).Name, U.UserName)
        If currDomainList.Count = 0 Then
            addLOG("CONSOLE:Cannot find in LDAP: " + U.UserName)
            Exit Sub
        End If
        Dim auditNDXopt As Integer = 0

        Dim useLDAP As Boolean
        Dim nameNoLDAP$ = ""
        nameNoLDAP = stripToFilename(U.UserName)
        Dim emailNDX As Integer = 0
        Dim lastNameNDX As Integer = 0
        Dim phoneNDXopt As Integer = 0
        Dim cellNDXopt As Integer = 0
        Dim activeNDXopt As Integer = 0
        Dim lcidNDXopt As Integer = 0

        Dim expireNDXopt As Integer = 0
        Dim jobNDXopt As Integer = 0

        If currDomainList.Count > 1 Then
            addLOG("CONSOLE:WARNING: MULTIPLE USERS FOUND IN LDAP FOR " + nameNoLDAP + " - Choosing first, email " + currDomainList(0).Email)
        End If

        U.Email = currDomainList(0).Email
        U.FirstName = currDomainList(0).FirstName
        U.LastName = currDomainList(0).LastName


        'set up role data
        U.RoleData = New CxPortal.Role
        If U.RoleData.Name = "Scanner" Then U.RoleData.ID = 0
        If U.RoleData.Name = "Reviewer" Then U.RoleData.ID = 1

        U.IsActive = True

        '--------------Dim tempBOOL$ = ""----------------------------------------------
        Dim tempBOOL$ = ""
        'audit or not
        If auditNDXopt = 0 Then
            U.AuditUser = False

        End If

        'phone
        If phoneNDXopt Then
            ' U.Phone = CStr(big3D(curRow, phoneNDXopt - 1))
        End If

        'jobtitle
        If jobNDXopt Then
            'U.JobTitle = CStr(big3D(curRow, jobNDXopt - 1))
        End If

        'cell
        If cellNDXopt Then
            'U.CellPhone = CStr(big3D(curRow, cellNDXopt - 1))
        End If

        Dim tDate$ = ""
        Dim a$ = ""

        'expire
        If expireNDXopt = 0 Then
            a$ = defaultS.userExpire
            If Len(a) = 0 Then a = "365"
            tDate$ = ""
            If a = "EOY" Then
                tDate = "12/31/" + Trim(Str(Today.Year))
            End If
            If a = "EOL" Then
                tDate = defaultS.lastDayOfLicense
            End If

            If Len(tDate) Then a = Trim(Str(DateDiff(DateInterval.Day, Today, CDate(tDate))))

            U.willExpireAfterDays = Val(a)
        Else
            'tDate$ = CStr(big3D(curRow, expireNDXopt - 1))
            If InStr(tDate, "/") Then a = Trim(Str(DateDiff(DateInterval.Day, Today, CDate(tDate)))) Else a = tDate
            If Val(a) = 0 Then a = "365"
            U.willExpireAfterDays = Val(a)
        End If

        'dont need to build unsubscribedGroups to add users
        Dim currUser$ = ""
        Dim G(1000) As CxPortal.Group
        Dim numGroups As Integer = 0
        'now groups



        If numGroups Then
            Array.Resize(G, numGroups)
            U.GroupList = G
            '                addLOG("Added " + Trim(Str(numGroups)) + " to profile of " + nameNoLDAP)
        Else
            addLOG("ERROR: No groups to define for user " + nameNoLDAP + " - Users must belong to at least 1 group.")
            '    GoTo nextRow
        End If


        Dim addArgs As New backgroundUserArgs
        addArgs.addORedit = "add"
        addArgs.isLDAP = useLDAP
        addArgs.U = U
        ' addArgs.changeActiveState = changeActiveState

        If editORaddUser(addArgs) = False Then

        End If
    End Sub


    Private Function startSession() As Boolean
        startSession = False
        CxWrap = New CxWrapper
        Dim getSession$ = CxWrap.ActivateSession
        'here some change to trigger a push

        'here some change to trigger a push
        'here some change to trigger a push
        'here some change to trigger a push
        'here some change to trigger a push

        addLOG("Activating Session:" + getSession)
        If getSession <> "True" Then
            addLOG(getSession)
            End
        Else
            startSession = True
        End If

    End Function



    Private Sub dumpInstructs()
        addLOG("CONSOLE:CxCLI allows SOAP interaction with CxServer via Windows Command Line")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:Submit KEY/VALUE (K/V) pairs in the form key=value, example user=mhorty")
        addLOG("CONSOLE:Optional - use quotes around values to preserve text with spaces")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:For all calls, loggingenabled=true details activity inside currutil_log.txt")
        addLOG("CONSOLE: ")

        '        addLOG("CONSOLE: ")

        addLOG("CONSOLE:COMMAND       K/V Pairs            DETAIL")
        addLOG("CONSOLE:help                               Produces this help file")
        addLOG("CONSOLE:encrypt       text                 TEXT=text to encrypt - Returns encrypted data (ie for CxPassword in config.txt)")
        addLOG("CONSOLE:userexist     user,match           MATCH=username/mail, USER=user data - Returns true if user exists")

        addLOG("CONSOLE:getpresets                         Get list of Presets by ID")
        addLOG("CONSOLE:getpresetdef  id                   id=[PresetID] - Returns details of Preset")
        addLOG("CONSOLE:addpreset     name,queries         Name of Preset,Queries separated by commas to Add Preset")
        addLOG("CONSOLE:cancelscans                        Cancels all active scans in Queued Status")
        addLOG("CONSOLE:showscans     type,mins            TYPE=queued/failed/completed, MINS(optional)=Summarize for past X minutes")
        addLOG("CONSOLE:numscans      type,mins            TYPE=queued/failed/completed, MINS(optional)=Summarize for past X minutes")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:showteams                          List all teams with # of users assigned to that team")
        addLOG("CONSOLE:swapteams     from,to              Move users FROM (CxServer\SP\Company\TeamName) team TO a different team")
        addLOG("CONSOLE:showusers     team                 TEAM=Team name or 'ALL' for all users")
        addLOG("CONSOLE: ")

        addLOG("CONSOLE: ")
        addLOG("CONSOLE:login         user,pass            USER=username, PASS=password - Returns true if login successful")
        addLOG("CONSOLE:createhierarchy     hierarchy,ldapservername,ldapgroups      hierarchy=CxServer\SP\APAC\RND\Team1 ldapservername=MyActiveDirectory ldapgroups=APAC-RND-Scanner-All,APAC-RND-Reviewer-All")
        addLOG("CONSOLE: ")


        addLOG("CONSOLE:enableusers   file,match           MATCH=username/mail, FILE=text file of users, 1 per line")
        addLOG("CONSOLE:disableusers  file,match           MATCH=username/mail, FILE=text file of users, 1 per line")
        addLOG("CONSOLE:deleteusers   file,match           MATCH=username/mail, FILE=text file of users, 1 per line")
        addLOG("CONSOLE:userexpire    days                 DAYS=show users expiring within X days")
        addLOG("CONSOLE:setuserexpire days,newdate         DAYS=Users expiring within X days,NEWDATE=MM/DD/YYYY")
        addLOG("CONSOLE:edituser      addgroup,match       MATCH=username/mail, ADDGROUP=group(s) to add user to")
        addLOG("CONSOLE:edituser      subtractgroup,match  MATCH=username/mail, SUBTRACTGROUP=group(s) to remove user from")
        addLOG("CONSOLE:edituser      changerole,match     MATCH=username/mail, CHANGEROLE=Scanner or Reviewer - Change role of user")
        addLOG("CONSOLE:edituser      status,match         MATCH=username/mail, STATUS=enable/disable/delete - Change status of user")
        addLOG("CONSOLE:adduser       ** ADDUSER HAS MANY PARAMETERS **")
        addLOG("CONSOLE:---The following parameters are REQUIRED:")
        addLOG("CONSOLE:   usertype      - either APPLICATION,LDAP or SAML")
        addLOG("CONSOLE:   username      - Username of user *without* prefix (eg LDAP\ or SAML\)")
        addLOG("CONSOLE:   role          - ServerManager,CompanyManager,SPManager,Scanner,Reviewer")
        addLOG("CONSOLE:   firstname     - First name")
        addLOG("CONSOLE:   lastname      - Last name")
        addLOG("CONSOLE:   team          - Fully qualified team name(s) eg CxServer\SP\Company\Team1,CxServer\SP\Company\Team2")
        addLOG("CONSOLE:   email         - User email address")
        addLOG("CONSOLE:   password      - User password (required only if APPLICATION user type")
        addLOG("CONSOLE:---The following parameters are OPTIONAL:")
        addLOG("CONSOLE:   activeuser    - Boolean, Determines if user is active [DEFAULT=TRUE]")
        addLOG("CONSOLE:   audituser     - Boolean, Determines if user has CxAudit permission [DEFAULT=FALSE]")
        addLOG("CONSOLE:   expiredays    - Number of days before user expires [DEFAULT=365]")
        addLOG("CONSOLE:   langlcid      - Language type [DEFAULT=1033 (English/US)]")
        addLOG("CONSOLE:   country       - Country")
        addLOG("CONSOLE:   phone         - Phone")
        addLOG("CONSOLE:   cellphone     - Cellphone")



        'optional
        ' "jobtitle"
        '        dataName = "Country" : dataNDX = returnUserInfo(rowOfHeaders, dataName$) : If dataNDX = 0 Then missingData += dataName + " "
        '        dataName = "Phone" : dataNDX = returnUserInfo(rowOfHeaders, dataName$) : If dataNDX = 0 Then missingData += dataName + " "
        '        dataName = "Cellphone" : dataNDX = returnUserInfo(rowOfHeaders, dataName$) : If dataNDX = 0 Then missingData += dataName + " "
        '        dataName = "LanguageLCID" : dataNDX = returnUserInfo(rowOfHeaders, dataName$) : If dataNDX = 0 Then missingData += dataName + " "
        '        dataName = "AuditUser" : dataNDX = returnUserInfo(rowOfHeaders, dataName$) : If dataNDX = 0 Then missingData += dataName + " "
        '        dataName = "ActiveUser" : dataNDX = returnUserInfo(rowOfHeaders, dataName$) : If dataNDX = 0 Then missingData += dataName + " "
        '        dataName = "ExpirationDays" : dataNDX = returnUserInfo(rowOfHeaders, dataName$) : If dataNDX = 0 Then missingData += dataName + " "

        'On Error GoTo errorcatch
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:*******                           USAGE EXAMPLES                            ******* ")

        addLOG("CONSOLE: ")
        addLOG("CONSOLE:Example, encrypt your password:")
        addLOG("CONSOLE:CMD>cxcli encrypt text=passwordtext")
        addLOG("CONSOLE:encrypted_text=yADi4Mkw3cUqC8mtiUOyh/dZ5TuzCl5i4Gx0hmVftw8=")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:Example, check user's existence:")
        addLOG("CONSOLE:CMD>cxcli userexist user=mhorty match=username")
        addLOG("CONSOLE:TRUE")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")

        addLOG("CONSOLE:Example, user login:")
        addLOG("CONSOLE:CMD>cxcli login user=mhorty pass=password")
        addLOG("CONSOLE:Login successful for mhorty")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")

        addLOG("Example,Show all users expiring within a certain number of days:")
        addLOG("CONSOLE:--------------------------------------------------------")
        addLOG("CONSOLE:C:\>cxcli userexpire days=365")
        addLOG("CONSOLE:miketyson mt@cx.com expires in 76 days")
        addLOG("CONSOLE:jerryseinfeld jf@cx.com expires in -119 days *EXPIRED*")
        addLOG("CONSOLE:janedoe jd@cx.com expires in -119 days *EXPIRED*")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")

        addLOG("Example,Set *new* expiration for all users expiring within a certain number of days:")
        addLOG("CONSOLE:C:\>cxcli setuserexpire days=-100 newdate=4/15/2019")
        addLOG("CONSOLE:For users expiring in <-100 days, set to 4 days (4/15/2019)")
        addLOG("CONSOLE:jerryseinfeld jf@cx.com now expires in 4 days: True")
        addLOG("CONSOLE:janedoe jd@cx.com now expires in 4 days: True")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")

        addLOG("Example,Move users from team ..\Nested3 to team ..\ProjectComponent1")
        addLOG("CONSOLE:C:\>cxcli swapteams from=CxServer\SP\Company\TestAutomation\Nested3 to=CxServer\SP\Company\ProjectName\ProjectComponent1")
        addLOG("CONSOLE:Moving all users that are part of CxServer\SP\Company\TestAutomation\Nested3 into CxServer\SP\Company\ProjectName\ProjectComponent1")
        addLOG("CONSOLE:addgroup=cxserver\sp\company\projectname\projectcomponent1 miketyson:mt@cx.com")
        addLOG("CONSOLE:Submitting Group change - TRUE")
        addLOG("CONSOLE:addgroup=cxserver\sp\company\projectname\projectcomponent1 SAML\testuser1:testuser@cx.com")
        addLOG("CONSOLE:Submitting Group change - TRUE")
        addLOG("CONSOLE:subtractgroup=cxserver\sp\company\testautomation\nested3 miketyson:mt@cx.com")
        addLOG("CONSOLE:Submitting Group change - TRUE")
        addLOG("CONSOLE:subtractgroup=cxserver\sp\company\testautomation\nested3 SAML\testuser1:testuser@cx.com")
        addLOG("CONSOLE:Submitting Group change - TRUE")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")

        addLOG("Example,Show all teams and the number of users assigned to them")
        addLOG("CONSOLE:C:\>cxcli showteams")
        addLOG("CONSOLE:Team:CxServer GUID:00000000-1111-1111-b111-989c9070eb11 USERS:2")
        addLOG("CONSOLE:Team:CxServer\SP GUID:11111111-2222-448d-b029-989c9070eb22 USERS:1")
        addLOG("CONSOLE:Team:CxServer\SP\Company GUID:22222222-2222-448d-b029-989c9070eb22 USERS:3")
        addLOG("CONSOLE:Team:CxServer\SP\Company\CompanyA_TeamA GUID:f0df3df3-00ea-41f2-b4fe-c87436e526fa USERS:0")
        addLOG("CONSOLE:Team:CxServer\SP\Company\CompanyA_TeamA\BookstoreParent GUID:c24b74d3-13a3-4a89-b67e-51599bd18697 USERS:0")
        addLOG("CONSOLE:Team:CxServer\SP\Company\CompanyA_TeamA\BookstoreParent\TeamA GUID:e89cb529-7a49-405d-b497-032af56b05d4 USERS:1")
        addLOG("CONSOLE:Team:CxServer\SP\Company\CompanyA_TeamA\BookstoreParent\TeamA\SomeNestedTeam GUID:d90ae49e-0e3a-4ef9-a8e4-dc3b99a85637 USERS:0")
        addLOG("CONSOLE:Team:CxServer\SP\Company\CompanyA_TeamA\BookstoreParent\TeamB GUID:bbd8a7d2-a5df-4959-8a6d-8ae42d8fc58b USERS:1")
        addLOG("CONSOLE:Team:CxServer\SP\Company\ProjectName GUID:802e408e-70cc-4052-bfd1-ece0c7f1333f USERS:0")
        addLOG("CONSOLE:Team:CxServer\SP\Company\ProjectName\ProjectComponent1 GUID:4988d81d-5a49-4682-81a5-6cf1b4abfe90 USERS:0")
        addLOG("CONSOLE:Team:CxServer\SP\Company\ProjectName\ProjectComponent2 GUID:e93c963e-ce72-44eb-b039-b9a03c4ccd72 USERS:3")
        addLOG("CONSOLE:Team:CxServer\SP\Company\TestAutomation GUID:ba3c66b9-c417-45ea-ad32-add388ffad97 USERS:1")
        addLOG("CONSOLE:Team:CxServer\SP\Company\TestAutomation\Nested1 GUID:78fb498d-715a-4bb8-94ef-564ae1437e7a USERS:2")
        addLOG("CONSOLE:Team:CxServer\SP\Company\TestAutomation\Nested2 GUID:b604a370-827c-4a38-8307-8e0d6e4d6b7f USERS:1")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")

        addLOG("Example,Show all users assigned to a specific team")

        addLOG("CONSOLE:C:\>cxcli showusers team=CxServer\SP\Company\Users")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli EMAIL:testcxcli@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli11 EMAIL:testcxcli11@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli22 EMAIL:test222@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli33 EMAIL:test333@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli44 EMAIL:testcxcli44@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli55 EMAIL:testcxcli55@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli555 EMAIL:testcxcli555@cx.com")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")

        addLOG("Example,Show all users by team")
        addLOG("CONSOLE:C:\>cxcli showusers team=ALL")
        addLOG("CONSOLE:Team:CxServer GUID:00000000-1111-1111-b111-989c9070eb11 USERNAME:mhorty EMAIL:admin@cx.com")
        addLOG("CONSOLE:Team:CxServer GUID:00000000-1111-1111-b111-989c9070eb11 USERNAME:testcxcli66 EMAIL:testcxcli66@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP GUID:11111111-2222-448d-b029-989c9070eb22 USERNAME:tnusertest EMAIL:tnu@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company GUID:22222222-2222-448d-b029-989c9070eb22 USERNAME:jerryseinfeld EMAIL:jf@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company GUID:22222222-2222-448d-b029-989c9070eb22 USERNAME:janedoe EMAIL:jd@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company GUID:22222222-2222-448d-b029-989c9070eb22 USERNAME:SAML\testuser2 EMAIL:testuser2@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\CompanyA_TeamA\BookstoreParent\TeamA GUID:e89cb529-7a49-405d-b497-032af56b05d4 USERNAME:ghi111 EMAIL:ghi111@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\CompanyA_TeamA\BookstoreParent\TeamB GUID:bbd8a7d2-a5df-4959-8a6d-8ae42d8fc58b USERNAME:ghi111 EMAIL:ghi111@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\ProjectName\ProjectComponent1 GUID:4988d81d-5a49-4682-81a5-6cf1b4abfe90 USERNAME:miketyson EMAIL:mt@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\ProjectName\ProjectComponent1 GUID:4988d81d-5a49-4682-81a5-6cf1b4abfe90 USERNAME:SAML\testuser1 EMAIL:testuser@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\ProjectName\ProjectComponent2 GUID:e93c963e-ce72-44eb-b039-b9a03c4ccd72 USERNAME:miketyson EMAIL:mt@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\ProjectName\ProjectComponent2 GUID:e93c963e-ce72-44eb-b039-b9a03c4ccd72 USERNAME:SAML\testuser1 EMAIL:testuser@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\ProjectName\ProjectComponent2 GUID:e93c963e-ce72-44eb-b039-b9a03c4ccd72 USERNAME:def111 EMAIL:def111@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\TestAutomation GUID:ba3c66b9-c417-45ea-ad32-add388ffad97 USERNAME:abc111 EMAIL:abc111@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\TestAutomation2\Nested1 GUID:4a2af321-2fb6-410e-a5dd-5dc380539a4a USERNAME:def111 EMAIL:def111@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli EMAIL:testcxcli@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli11 EMAIL:testcxcli11@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli22 EMAIL:test222@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli33 EMAIL:test333@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli44 EMAIL:testcxcli44@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli55 EMAIL:testcxcli55@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\Company\Users GUID:22222222-2222-448d-b029-989c9070eb23 USERNAME:testcxcli555 EMAIL:testcxcli555@cx.com")
        addLOG("CONSOLE:Team:CxServer\SP\CompanyB GUID:96c13d34-816f-4db7-ba65-17506de3142a USERNAME:SAML\testuser3 EMAIL:testuser8@cx.com")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")



        addLOG("CONSOLE:Examples, enable/disable/delete users from file:")
        addLOG("CONSOLE:------------------------------------------------ ")
        addLOG("CONSOLE:CMD>cxcli disableusers file=" + Chr(34) + "c:\Folder Name\some_file_1_username_per_line.txt" + Chr(34) + " match=username")

        addLOG("CONSOLE:Matching username on  3 entries")
        addLOG("CONSOLE:User SAML\testuser1: testuser@cx.com  disabled")
        addLOG("CONSOLE:User tnusertest: tnu@cx.com  disabled")
        addLOG("CONSOLE:User miketyson: mt@cx.com  disabled")
        addLOG("CONSOLE: ")

        addLOG("CONSOLE:CMD>CxCLI enableusers file=" + Chr(34) + "c:\Folder Name\some_file_1_email_per_line.txt" + Chr(34) + " match=mail")
        addLOG("CONSOLE:Matching mail on  3 entries")
        addLOG("CONSOLE:User SAML\testuser1: testuser@cx.com  enabled")
        addLOG("CONSOLE:User tnusertest: tnu@cx.com  enabled")
        addLOG("CONSOLE:User miketyson: mt@cx.com  enabled")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:Examples, various functions of edituser:")
        addLOG("CONSOLE:---------------------------------------- ")
        addLOG("CONSOLE:-------Roles")
        addLOG("CONSOLE:CMD>CxCLI edituser user=mt@cx.com match=mail changerole=Reviewer")
        addLOG("CONSOLE:changerole = reviewer miketyson: mt@cx.com")
        addLOG("CONSOLE:No change necessary")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:CMD>cxcli edituser user=miketyson match=username changerole=Scanner")
        addLOG("CONSOLE:changerole = scanner miketyson: mt@cx.com")
        addLOG("CONSOLE:Changing role from Scanner To Scanner")
        addLOG("CONSOLE:Submitting Role Change:   True")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:------Groups")
        addLOG("CONSOLE:CMD>CxCLI edituser user=def111@cx.com match=mail addgroup=TeamA")
        addLOG("CONSOLEaddgroup = teama def111:def111@cx.com")
        addLOG("CONSOLE:Submitting change - True")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:CMD>cxcli edituser user=def111@cx.com match=mail subtractgroup=TeamA,ProjectComponent1")
        addLOG("CONSOLE:subtractgroup = teama,projectcomponent1 def111:def111@cx.com")
        addLOG("CONSOLE:Submitting change - True")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:CMD>CxCLI edituser user=def111@cx.com match=mail addgroup=NonExistTeam")
        addLOG("CONSOLE:addgroup = nonexistteam def111def111@cx.com")
        addLOG("CONSOLE:Submitting change - ERROR: def111 -General error occurred. Please refer to application admin - Could Not update groups")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:------Status")
        addLOG("CONSOLE:CMD>CxCLI edituser user=502752445 match=username status=delete")
        addLOG("CONSOLE:status = delete 502752445:Service.account@cust.com")
        addLOG("CONSOLE:User 502752445:Service.account@cust.com  deleted")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:CMD>CxCLI edituser user=def111@cx.com match=mail status=enable")
        addLOG("CONSOLE:status = enable def111:def111@cx.com")
        addLOG("CONSOLE:User def111:def111@cx.com  enabled")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE:CMD>CxCLI edituser user=def111@cx.com match=mail status=disable")
        addLOG("CONSOLE:status = disable def111:def111@cx.com")
        addLOG("CONSOLE:User def111:def111@cx.com  disabled")
        addLOG("CONSOLE: ")
        addLOG("CONSOLE: ")

        addLOG("CONSOLE:Examples, add users:")
        addLOG("CONSOLE:c:\>cxcli adduser username=testcxcli55 firstname=Test55 lastname=CxCLI role=Scanner email=testcxcli55@cx.com team=CxServer\SP\Company\Users usertype=application password=Password12345!")
        addLOG("CONSOLE:DEFAULT audituser=false")
        addLOG("CONSOLE:DEFAULT activeuser=true")
        addLOG("CONSOLE:DEFAULT langlcid=1033 (English)")
        addLOG("CONSOLE:DEFAULT expiredays=365")
        addLOG("CONSOLE:Adding user application testcxcli55 testcxcli55@cx.com Test55 CxCLI 0(Scanner) CxServer\SP\Company\Users 1 1033 False True 365")
        addLOG("CONSOLE:User added: True")
        addLOG("CONSOLE:")
        addLOG("CONSOLE:c:\>cxcli adduser username=testcxcli66 firstname=Test56 lastname=CxCLI role=ServerManager email=testcxcli66@cx.com team=CxServer usertype=application password=Password12345!")
        addLOG("CONSOLE:DEFAULT audituser=false")
        addLOG("CONSOLE:DEFAULT activeuser=true")
        addLOG("CONSOLE:DEFAULT langlcid=1033 (English)")
        addLOG("CONSOLE:DEFAULT expiredays=365")
        addLOG("CONSOLE:Adding user application testcxcli66 testcxcli66@cx.com Test56 CxCLI 5(ServerManager) CxServer 1 1033 False True 365")
        addLOG("CONSOLE:User added: True")

    End Sub





    ' ------- editing user CLI for Salesforce

    Private Sub editUserCLI(cArgs As CLIArgs)
        Dim U As CxPortal.UserData
        U = returnUserUsingMatch(cArgs.matchOn, cArgs.uData)

        If IsNothing(U) = True Or U.UserName = "" Or U.UserName = "NOTFOUND" Then
            addLOG("CONSOLE:Cannot find User with '" + cArgs.uData + "' set as " + cArgs.matchOn + "  - can match on either username or mail")
            Exit Sub
        End If

        cArgs.editCmd = LCase(cArgs.editCmd)
        cArgs.newVal = LCase(cArgs.newVal)
        addLOG("CONSOLE:" + cArgs.editCmd + "=" + cArgs.newVal + " " + U.UserName + ":" + U.Email)

        Select Case cArgs.editCmd
            Case "addgroup"
                Call addTeamsByCSV(cArgs.newVal, U)
                ' For Each G In U.GroupList
                'addLOG("CONSOLE:Group:" + G.GroupName + " GUID:" + G.Guid + " ID:" + G.ID + " PATH:" + G.Path + " FULLPATH:" + G.FullPath)
                'Next
                Dim addArgs As New backgroundUserArgs
                addArgs.addORedit = "edit"
                addArgs.U = U
                addArgs.changeActiveState = False
                addArgs.unSubscribed = buildUnsubscribedList(U)
                addLOG("CONSOLE:Submitting Group change - " + editORaddUser(addArgs))

            Case "subtractgroup"
                Call subtractTeamsByCSV(cArgs.newVal, U)
                Dim addArgs As New backgroundUserArgs
                addArgs.addORedit = "edit"
                addArgs.U = U
                addArgs.changeActiveState = False
                addArgs.unSubscribed = buildUnsubscribedList(U)
                addLOG("CONSOLE:Submitting Group change - " + editORaddUser(addArgs))

            Case "changerole"
                Dim a$ = cArgs.newVal
                Dim changeS$ = ""
                Dim origR$ = U.RoleData.Name
                If LCase(a) = "scanner" Then
                    If LCase(U.RoleData.Name) <> "scanner" Then
                        changeS = "Scanner"
                        U.RoleData = New CxPortal.Role
                        U.RoleData.Name = "Scanner"
                        U.RoleData.ID = 0
                    End If
                End If
                If LCase(a) = "reviewer" Then
                    If LCase(U.RoleData.Name) <> "reviewer" Then
                        changeS = "Reviewer"
                        U.RoleData = New CxPortal.Role
                        U.RoleData.Name = "Reviewer"
                        U.RoleData.ID = 1
                    End If
                End If



                If Len(changeS) Then
                    addLOG("CONSOLE:Changing role from " + origR$ + " to " + changeS)
                    Call addLOG("CONSOLE:Submitting Role Change: " + CxWrap.CXeditUser(U))
                Else
                    addLOG("CONSOLE:No change necessary")
                End If


            Case "status"
                If cArgs.newVal = "disable" Then Call CXdisableUser(cArgs.matchOn, cArgs.uData)
                If cArgs.newVal = "enable" Then Call CXdisableUser(cArgs.matchOn, cArgs.uData, True)
                If cArgs.newVal = "delete" Then Call CXdisableUser(cArgs.matchOn, cArgs.uData,, True)


        End Select


    End Sub

    Private Sub editExpiration(ByRef U As CxPortal.UserData, ByVal numDaysToExpire As Integer)
        U.willExpireAfterDays = numDaysToExpire
        Call addLOG("CONSOLE:" + U.UserName + " " + U.Email + " setting to expire in " + numDaysToExpire.ToString + " days " + CxWrap.CXeditUser(U))
    End Sub


    Private Function editORaddUser(Args As backgroundUserArgs) As String
        editORaddUser = "TRUE"
        Dim actTXT$ = "EDIT:" + Trim(Str(Args.U.ID)) + "," + Args.U.UserName + "," + Args.U.LastName + "," + Args.U.FirstName + " # Groups: " + Args.U.GroupList.Count.ToString

        addLOG(actTXT)
        Dim resultOfCall$ = ""

        Select Case Args.addORedit
            Case "edit"
                resultOfCall$ = CxWrap.CXeditUser(Args.U)
                If resultOfCall$ = "True" Then
                    addLOG(actTXT)
                Else
                    resultOfCall = "ERROR: " + Args.U.UserName + " - " + resultOfCall$
                    addLOG(resultOfCall)
                End If


                Dim tGrps(100) As CxPortal.Group
                Dim numGrps As Integer = 0
                Dim K As Integer = 0


                addLOG("EditOrAdd: " + Args.U.GroupList.Count.ToString + " Groups total")

                For Each G As CxPortal.Group In Args.U.GroupList
                    tGrps(numGrps) = G
                    numGrps += 1
                Next

                Dim resP$ = ""

                For K = 0 To numGrps - 1
                    Args.U.GroupList() = {tGrps(K)}

                    resP = CxWrap.CXeditUserGroups(Args.U, Args.unSubscribed)
                    If resP = "True" Then
                        'addLOG("Updated group " + Args.U.GroupList(0).GroupName)
                    Else
                        editORaddUser = "ERROR: " + Args.U.UserName + " - " + resP + " - Could not update groups"
                        addLOG(editORaddUser)
                    End If
                Next

                If Args.changeActiveState = True Then

                    Dim succesS$
                    succesS = CxWrap.CXsetUserActivationState(Args.U.ID, Args.U.IsActive)

                    Dim actionStr$ = " disable"
                    If Args.U.IsActive = True Then actionStr = " enable"
                    If succesS = True Then

                        addLOG("User " + Args.U.UserName + ": " + Args.U.FirstName + " " + Args.U.LastName + actionStr + "d")

                    Else
                        editORaddUser = "ERROR: Could not" + actionStr + " User " + Args.U.UserName + ":  " + Args.U.FirstName + " " + Args.U.LastName
                        addLOG(editORaddUser)

                    End If
                End If

            Case "add"
                actTXT = "ADD:" + Args.U.UserName + "," + Args.U.LastName + "," + Args.U.FirstName
                With Args.U
                    .DateCreated = Now
                    .LastLoginDate = Now
                    .LimitAccessByIPAddress = False
                    '                    .Password = "Password123!"
                End With

                Dim someResult$

                someResult = CxWrap.CXaddUser(Args.U, Args.isLDAP)

                If someResult = "True" Then
                    addLOG("CONSOLE:SUCCESS: " + actTXT)
                Else
                    editORaddUser = "CONSOLE:ERROR:" + someResult + " - " + actTXT
                    addLOG(editORaddUser)
                End If



                'ignore this - have to obtain user ID in order to change active state
                Args.changeActiveState = False

                If Args.changeActiveState = True Then

                    Dim succesS$
                    succesS = CxWrap.CXsetUserActivationState(Args.U.ID, Args.U.IsActive)

                    Dim actionStr$ = " disable"
                    If Args.U.IsActive = True Then actionStr = " enable"
                    If succesS = True Then
                        addLOG("User " + Args.U.UserName + ": " + Args.U.FirstName + " " + Args.U.LastName + actionStr + "d")

                    Else
                        editORaddUser = "ERROR: Could not" + actionStr + " User " + Args.U.UserName + ":  " + Args.U.FirstName + " " + Args.U.LastName
                        addLOG(editORaddUser)

                    End If
                End If


        End Select


    End Function



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



    Private Function returnGroupOfGUID(ByVal guiD$) As CxPortal.Group
        returnGroupOfGUID = New CxPortal.Group
        For Each G In allGroups
            If G.ID = guiD Then
                With returnGroupOfGUID
                    .FullPath = G.FullPath
                    .Guid = G.Guid
                    .GroupName = stripToFilename(G.GroupName)
                    .ID = G.ID
                    .Type = G.Type
                    .Path = G.Path
                End With
                Exit Function
            End If
        Next
    End Function

    Private Function doesTeamExist(ByRef grpS() As CxPortal.Group, ByRef gName$) As Boolean
        doesTeamExist = False
        For Each G In grpS
            If IsNothing(G) = False Then
                If LCase(stripToFilename(G.GroupName)) = LCase(stripToFilename(gName)) Then
                    doesTeamExist = True
                    Exit Function
                End If
            End If
        Next
    End Function

    Private Function buildUnsubscribedList(ByRef U As CxPortal.UserData) As CxPortal.Group()
        'when submitting user you must submit user with U.GROUPLIST along with a separate list of groups they are *not* assigned to
        'deemed too high of an loe by SFDC - built edit capability borrowing from usermgmt control and big3d code
        '
        '
        ' Assumption is U has all groups after adding or subtracting

        Dim unsubscribedGroups(10000) As CxPortal.Group

        Dim G(1000) As CxPortal.Group
        Dim numGroups As Integer = 0
        For Each userTeam In U.GroupList
            If IsNothing(userTeam) = False Then
                If doesTeamExist(G, userTeam.GroupName) = False Then
                    G(numGroups) = userTeam
                    numGroups += 1
                End If
            End If
        Next
        Array.Resize(G, numGroups)
        U.GroupList = G

        Dim numUnsub As Integer = 0 ' keep track of groups user is not subscribed to
        For Each allG In allGroups
            If doesTeamExist(G, allG.GroupName) = False Then
                Dim NG As New CxPortal.Group
                NG = returnGroupOfGUID(allG.ID)
                unsubscribedGroups(numUnsub) = NG
                numUnsub += 1
            End If
        Next

        Array.Resize(unsubscribedGroups, numUnsub)

        Return unsubscribedGroups

    End Function


    Private Sub addTeamsByCSV(ByVal csV$, ByRef U As CxPortal.UserData)

        'this adds to existing teams of user
        addLOG("Adding to USER " + U.UserName + ": " + csV)
        addLOG("User assigned to " + U.GroupList.Count.ToString + " teams before edit")
        'adding group GUIDs
        Dim numGroups As Integer
        Dim a$ = ""
        Dim grpString() As Object

        grpString = Split(csV, ",")
        numGroups = UBound(grpString)

        addLOG("Adding to " + U.ID.ToString + ":" + U.UserName + " " + csV)

        Dim G(1000) As CxPortal.Group

        Dim nG As Integer
        For nG = 0 To UBound(grpString)
            a$ = grpString(nG)
            If Len(a) Then
                G(numGroups) = returnGroupOfGUID(getGUIDofTEAM(a)) ' New CxPortal.Group
                '  G(numGroups).ID = getGUIDofTEAM(a, True)
                numGroups += 1
            End If
        Next nG


        'this contains all of the groups from the parameter
        Dim numAdded As Integer = 0


        ' need to add from user to catch groups that were already assigned to
        For Each userTeam In U.GroupList
            If doesTeamExist(G, userTeam.GroupName) = False Then
                G(numGroups) = userTeam
                numGroups += 1
            End If
        Next

        Array.Resize(G, numGroups)

        U.GroupList = G ' users groups now aligned with new and original adds

        addLOG("User assigned to " + U.GroupList.Count.ToString + " teams after edit")

    End Sub


    Private Sub subtractTeamsByCSV(ByVal csV$, ByRef U As CxPortal.UserData)

        'this adds to existing teams of user
        addLOG("Subtracting teams from USER " + U.UserName + ": " + csV)
        addLOG("User assigned to " + U.GroupList.Count.ToString + " teams before edit")
        'adding group GUIDs
        Dim numGroups As Integer = 0
        Dim a$ = ""
        Dim grpString() As Object

        grpString = Split(csV, ",")

        Dim G(100) As CxPortal.Group
        ' g(100) contains list of groups to remove from user

        Dim nG As Integer
        For nG = 0 To UBound(grpString)
            a$ = grpString(nG)
            If Len(a) Then
                G(numGroups) = returnGroupOfGUID(getGUIDofTEAM(a, True)) ' New CxPortal.Group
                numGroups += 1
            End If
        Next nG

        ReDim Preserve G(numGroups - 1)

        Dim userTeam As CxPortal.Group

        Dim newList(numGroups + U.GroupList.Count - 1) As CxPortal.Group

        numGroups = 0 ' init var for use in this loop
        For nG = 0 To U.GroupList.Count - 1
            userTeam = U.GroupList(nG)
            If doesTeamExist(G, userTeam.GroupName) = False Then 'team needs to be removed from profile
                newList(numGroups) = userTeam
                numGroups += 1
            End If
        Next

        ReDim Preserve newList(numGroups - 1)


        U.GroupList = newList ' users groups now aligned with new and original adds

        addLOG("User assigned to " + U.GroupList.Count.ToString + " teams after edit")

    End Sub





    Private Function returnRoleID(rolE$) As String
        rolE = LCase(rolE)
        returnRoleID = ""
        Select Case rolE
            Case "scanner"
                Return "0"
            Case "reviewer"
                Return "1"
            Case "company manager", "companymanager"
                Return "2"
            Case "sp manager", "spmanager"
                Return "4"
            Case "server manager", "servermanager"
                Return "5"
        End Select
    End Function

    Private Function returnRoleString(id As Integer) As String
        returnRoleString = ""
        Select Case id
            Case 0
                Return "Scanner"
            Case 1
                Return "Reviewer"
            Case 2
                Return "Company Manager"
            Case 4
                Return "SP Manager"
            Case 5
                Return "Server Manager"
        End Select

    End Function

    Private Sub setGrpList(ByRef U As CxPortal.UserData, ByRef G As List(Of CxPortal.Group))
        Dim numTeams As Integer = G.Count

        'dumb routine - must submit an array of groups as { group1, group2, group3 } - must be more intuitive way that
        'does not require known number but cannot submit LIST .NET object
        Select Case numTeams
            Case 1
                U.GroupList = {G(0)}

            Case 2
                U.GroupList = {G(0), G(1)}

            Case 3
                U.GroupList = {G(0), G(1), G(2)}

            Case 4
                U.GroupList = {G(0), G(1), G(2), G(3)}

            Case 5
                U.GroupList = {G(0), G(1), G(2), G(3), G(4)}

            Case 6
                U.GroupList = {G(0), G(1), G(2), G(3), G(4), G(5)}

        End Select
    End Sub

    Private Function getGroupList(teamNames$) As List(Of CxPortal.Group)
        getGroupList = New List(Of CxPortal.Group)

        Dim tName() As String = Split(teamNames, ",")

        Dim tCtr As Integer
        For tCtr = 0 To UBound(tName)
            Dim G As New CxPortal.Group
            G.Guid = getGUIDofTEAM(LTrim(tName(tCtr)))
            getGroupList.Add(G)
        Next

    End Function


    Private Sub addUser(userType$, userName$, rolename$, firstName$, lastName$, Team$, eMail$, jobtitle$, countrY$, phonE$, cellPhone$, langLCID$, auditUser$, activeUser$, expireDays$, Optional ByVal passworD$ = "")

        If userType = "" Or userName = "" Or rolename$ = "" Or firstName = "" Or lastName = "" Or Team = "" Or eMail = "" Then
            addLOG("CONSOLE:The following parameters are required:")
            addLOG("CONSOLE:usertype      - either APPLICATION,LDAP or SAML")
            addLOG("CONSOLE:username      - Username of user *without* prefix (eg LDAP\ or SAML\)")
            addLOG("CONSOLE:role          - ServerManager,CompanyManager,SPManager,Scanner,Reviewer")
            addLOG("CONSOLE:firstname     - First name")
            addLOG("CONSOLE:lastname      - Last name")
            addLOG("CONSOLE:team          - Fully qualified team name(s) eg CxServer\SP\Company\Team1,CxServer\SP\Company\Team2")
            addLOG("CONSOLE:email         - User email address")
            Exit Sub
        End If

        If LCase(userType) = "application" And passworD = "" Then
            addLOG("CONSOLE:You must provide a password for APPLICATION user types")
            Exit Sub
        End If

        If auditUser = "" Then addLOG("CONSOLE:DEFAULT audituser=false")
        If activeUser = "" Then addLOG("CONSOLE:DEFAULT activeuser=true")
        If langLCID = "" Then addLOG("CONSOLE:DEFAULT langlcid=1033 (English)")
        If expireDays = "" Then addLOG("CONSOLE:DEFAULT expiredays=365")

        Dim U As New CxPortal.UserData
        With U

            .UserName = userName
            .FirstName = firstName
            .LastName = lastName
            .Email = eMail

            If LCase(userType).Equals("ldap") Then
                addLOG("CONSOLE: Lookup user in LDAP")

                Dim ldapuserPair As String() = userName.Split(New Char() {"\"c, " "c}, StringSplitOptions.RemoveEmptyEntries)
                Dim user As CxPortal.CxDomainUser = CxWrap.CXGetUserFromUserDirectory(ldapuserPair(0), ldapuserPair(1))

                If user IsNot Nothing Then
                    .FirstName = user.FirstName
                    .LastName = user.LastName
                    .Email = user.Email
                End If
            End If

            Dim rolE As New CxPortal.Role
            rolE.ID = returnRoleID(rolename)
            .RoleData = rolE


            Dim GL As List(Of CxPortal.Group) = getGroupList(Team)
            .GroupList = GL.ToArray
            'Call setGrpList(U, GL)

            'auto-fill
            .DateCreated = Now
            .LastLoginDate = Now
            .LimitAccessByIPAddress = False

            'optionally blank
            If Len(countrY) <> 0 Then .country = countrY
            If Len(jobtitle) <> 0 Then .JobTitle = jobtitle
            If Len(phonE) <> 0 Then .Phone = phonE
            If Len(cellPhone) <> 0 Then .CellPhone = cellPhone

            Dim miscData$ = ""
            'optional set default
            If Len(langLCID) <> 0 Then miscData = langLCID Else miscData = "1033"
            If Len(miscData) Then .UserPreferedLanguageLCID = Val(miscData)

            If Len(auditUser) <> 0 Then miscData = UCase(auditUser) Else miscData = "FALSE"
            If Len(miscData) Then .AuditUser = CBool(miscData)

            If Len(activeUser) <> 0 Then miscData = UCase(activeUser) Else miscData = "TRUE"
            If Len(miscData) Then .IsActive = CBool(miscData)

            If Len(expireDays) <> 0 Then miscData = expireDays Else miscData = "365"
            If Len(miscData) Then .willExpireAfterDays = Val(miscData)

        End With

        With U
            .DateCreated = Now
            .LastLoginDate = Now
            .LimitAccessByIPAddress = False
            If LCase(userType) = "application" Then U.Password = passworD
        End With


        addLOG("CONSOLE:Adding user " + userType + " " + U.UserName + " " + U.Email + " " + U.FirstName + " " + U.LastName + " " + U.RoleData.ID.ToString + "(" + rolename + ") " + Team + " " + U.GroupList.Count.ToString + " " + U.UserPreferedLanguageLCID.ToString + " " + U.AuditUser.ToString + " " + U.IsActive.ToString + " " + U.willExpireAfterDays.ToString)
        Select Case UCase(userType)
            Case "APPLICATION"
                addLOG("CONSOLE:User added: " + CxWrap.CXaddUser(U))

            Case "LDAP"
                addLOG("CONSOLE:User added: " + CxWrap.CXaddUser(U, True))

            Case "SAML"
                addLOG("CONSOLE:User added: " + CxWrap.CXaddUser(U,, True))
                addLOG("CONSOLE:You will need to execute SQL on the DB to ensure 'SAML\" + U.UserName + "' is the username for this user")

        End Select


        Exit Sub

    End Sub




    Private Function CXdisableUserFromFile(ByVal fileN$, ByVal matchOn$, Optional ByVal enableUser As Boolean = False, Optional ByVal deleteUser As Boolean = False) As Integer
        CXdisableUserFromFile = 0

        Dim C As Collection
        C = CSVFiletoCOLL(fileN)
        If C.Count = 0 Then
            addLOG("CONSOLE:No entries found")
        Else
            addLOG("CONSOLE:Matching " + matchOn + " on " + Str(C.Count) + " entries")
        End If

        Dim numEntries As Integer = C.Count
        Dim currEntry As Integer = 0

        For Each U In C
            currEntry += 1
            Call CXdisableUser(matchOn, U, enableUser, deleteUser)
        Next


    End Function

    Private Function CXdisableUserFromColl(ByVal matchOn$, ByRef userCollection As Collection, Optional ByVal enableUser As Boolean = False, Optional ByVal deleteUser As Boolean = False) As Integer
        CXdisableUserFromColl = 0

        Dim numEntries As Integer = userCollection.Count
        Dim currEntry As Integer = 0

        For Each U In userCollection
            currEntry += 1
            Call CXdisableUser(matchOn, U, enableUser, deleteUser)
        Next

    End Function

    Public Function returnUser(ByVal ID As Long) As CxPortal.UserData
        returnUser = New CxPortal.UserData
        For Each U In allUsers.UserDataList
            If U.ID = ID Then returnUser = U
        Next
    End Function

    Private Function returnUserUsingMatch(ByVal matchOn$, ByVal valEquals$) As CxPortal.UserData
        returnUserUsingMatch = New CxPortal.UserData
        returnUserUsingMatch.UserName = "NOTFOUND"
        For Each U In allUsers.UserDataList
            Select Case LCase(matchOn)
                Case "mail"
                    If LCase(U.Email) = LCase(valEquals) Then returnUserUsingMatch = U

                Case "name"
                    If LCase(U.FirstName) + " " + LCase(U.LastName) = LCase(valEquals) Then returnUserUsingMatch = U

                Case "username"
                    If LCase(U.UserName) = LCase(valEquals) Then returnUserUsingMatch = U

            End Select
        Next

    End Function

    Private Function CXdisableUser(ByVal matchOn$, ByVal valEquals$, Optional ByVal enableUser As Boolean = False, Optional ByVal deleteUser As Boolean = False) As Boolean
        Dim nU As CxPortal.UserData

        nU = returnUserUsingMatch(matchOn, valEquals)

        If IsNothing(nU) = True Or nU.UserName = "" Or nU.UserName = "NOTFOUND" Then
            addLOG("CONSOLE:User not found, " + matchOn + "=" + valEquals)
            Return False
            Exit Function
        End If

        Dim succesS As Boolean = False

        If deleteUser = True Then
            succesS = CxWrap.CXdeleteUser(nU.ID)
        Else
            succesS = CxWrap.CXsetUserActivationState(nU.ID, enableUser)
        End If

        Dim actionStr$ = " disable"
        If enableUser = True Then actionStr = " enable"
        If deleteUser = True Then actionStr = " delete"


        If succesS = True Then
            addLOG("CONSOLE:User " + nU.UserName + ":" + nU.Email + " " + actionStr + "d")
            Return True
        Else
            addLOG("CONSOLE:ERROR: Could not " + actionStr + " User " + nU.UserName + ":" + nU.Email)
            Return False
        End If
    End Function













    Public Sub addLOG(ByVal a$, Optional ByVal suppressDT As Boolean = False, Optional ByVal forceLog As Boolean = False)
        On Error GoTo errorcatch
        'logginGenabled = True

        If Mid(a, 1, 8) = "CONSOLE:" Then
            Console.WriteLine(Mid(a, 9))
        End If

forFileOnly:
        If suppressDT = False Then a = CStr(Now.ToLocalTime) + ": " + a

        Dim fileN$ = "currutil_log.txt"

        If loggingEnabled = False And forceLog = False Then GoTo writelineOnly

        Dim FF As Integer = FreeFile()

        If Dir(fileN) = "" Then
            FileOpen(FF, fileN, OpenMode.Output, OpenAccess.Write, OpenShare.Shared)
        Else
            FileOpen(FF, fileN, OpenMode.Append, OpenAccess.Write, OpenShare.Shared)
        End If


        Print(FF, a + vbCrLf)
        FileClose(FF)

writelineOnly:

errorcatch:
    End Sub


End Module
