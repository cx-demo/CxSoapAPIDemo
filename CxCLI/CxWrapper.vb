Imports System.Net
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

Public Class CxWrapper
    Public sessionID$ = ""
    Public sdkID$ = ""
    Public CxSDKProxy As CxSDKns.CxSDKWebServiceSoapClient
    Public CxProxy As CxPortal.CxPortalWebServiceSoapClient
    Public webURL$
    Public Event reportCompleted(ByRef R As getReportArgs) ', ByVal rptDate As DateTime)


    Public Function ActivateSession() As String
        ActivateSession = "True"
        sessionID = ""


        Dim pW3DES As New Simple3Des("2&#263gdjSiUEYkadhEII276#*763298")

        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()
        Dim CxSDKProxy = New CxSDKns.CxSDKWebServiceSoapClient()

        Dim creD = New CxPortal.Credentials()
        creD.User = filePROP("config.txt", "CXUN") : creD.Pass = filePROP("config.txt", "CXPW")
        creD.Pass = pW3DES.Decode(creD.Pass)

        Dim cred2 = New CxSDKns.Credentials()
        cred2.User = creD.User
        cred2.Pass = creD.Pass

        pW3DES = Nothing

        Dim lResp As CxPortal.CxWSResponseLoginData
        Dim lResp2 As CxSDKns.CxWSResponseLoginData

        lResp = CxProxy.Login(creD, 1033)
        lResp2 = CxSDKProxy.Login(cred2, 1033)

        If lResp.IsSuccesfull = True Then sessionID = lResp.SessionId

        If lResp2.IsSuccesfull = True Then sdkID = lResp2.SessionId

        creD.Pass = ""
        cred2.Pass = ""

        If sessionID = "" Or sdkID = "" Then
            addLOG("Cannot obtain sessionID! API calls will fail.")
            ActivateSession = "ERROR: Portal- " + lResp.ErrorMessage + " / SDK - " + lResp2.ErrorMessage
        Else
            addLOG("Obtained Session IDs For PORTAL And SDK APIs")
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




    Public Function CXgetLicenseData() As CxPortal.CxWSResponseServerLicenseData
        CXgetLicenseData = New CxPortal.CxWSResponseServerLicenseData
        CXgetLicenseData = CxProxy.GetServerLicenseData(sessionID)
        Return CXgetLicenseData
    End Function

    Public Function CXcompareScans(ByVal scan1 As Long, ByVal scan2 As Long) As CxPortal.CxWSSingleResultCompareData()
        Dim resP As CxPortal.CxWSResponceScanCompareResults

        resP = CxProxy.GetCompareScanResults(sessionID, scan1, scan2)

        CXcompareScans = resP.Results

    End Function

    Public Function CXgetVulnDetails(ByVal scanID As Long, ByVal pathID As Long) As CxPortal.CxWSResultPath

        Dim resP As CxPortal.CxWSResponceResultPath

        resP = CxProxy.GetResultPath(sessionID, scanID, pathID)

        If resP.IsSuccesfull = True Then
            CXgetVulnDetails = resP.Path
        Else
            CXgetVulnDetails = Nothing
        End If

    End Function

    Public Function CXgetScanResults(ByVal scanID As Long, ByRef SR() As CxPortal.CxWSSingleResultData) As String
        Dim resP As CxPortal.CxWSResponceScanResults

        resP = CxProxy.GetResultsForScan(sessionID, scanID)

        If resP.IsSuccesfull = True Then
            CXgetScanResults = "True"
            SR = resP.Results
        Else
            CXgetScanResults = resP.ErrorMessage
            SR = Nothing
        End If

    End Function

    Public Function CXgetAllProjects() As CxPortal.ProjectDisplayData()

        Dim resP As CxPortal.CxWSResponseProjectsDisplayData

        resP = CxProxy.GetProjectsDisplayData(sessionID)

        '        If resP.IsSuccesfull = True Then
        Return resP.projectList
        '       Else
        '      Return CXgetAllProjects
        '     End If

    End Function


    Public Function CxGetLdapServers(ByRef LD As List(Of CxPortal.CxWSLdapServerConfiguration)) As String
        CxGetLdapServers = ""
        LD = New List(Of CxPortal.CxWSLdapServerConfiguration)


        Dim checkLDAP As CxPortal.CxWSResponseLDAPServersConfiguration
        'checkLDAP = CxProxy.GetConfiguredLdapServerNames(sessionID, False)
        checkLDAP = CxProxy.GetLdapServersConfigurations(sessionID)

        If checkLDAP.IsSuccesfull Then
            LD = checkLDAP.serverConfigs.ToList
            CxGetLdapServers = "True" + " " + LD.Count.ToString

        Else
            CxGetLdapServers = "ERROR: Could not retrieve LDAP Configs - " + checkLDAP.ErrorMessage
        End If

    End Function

    Public Function CXgetLDAPUsers(ByVal ldapServerName$, ByVal searchTXT$) As List(Of CxPortal.CxDomainUser) ' ByRef allLdapUsers As List(Of CxPortal.CxDomainUser)) As String

        CXgetLDAPUsers = New List(Of CxPortal.CxDomainUser)

        Dim K As Integer = 0
        CxProxy = New CxPortal.CxPortalWebServiceSoapClient

        Dim LD As New CxPortal.CxWSResponseDomainUserList
        LD = CxProxy.GetAllUsersFromUserDirectory(sessionID, ldapServerName$, searchTXT, 1)

        If LD.IsSuccesfull = True And LD.UserList.Count Then
            CXgetLDAPUsers = LD.UserList.ToList
            '            CXgetLDAPUsers = "True"
        End If

    End Function



    Public Sub CXgetRoles(ByRef allRoles As List(Of CxPortal.Role))
        '        Dim R As CxPortal.CxWSBasicRepsonse
        allRoles = New List(Of CxPortal.Role)
        Dim K As Integer
        For K = 0 To 5
            Dim R As New CxPortal.Role
            R.ID = Trim(Str(K))
            Select Case K
                Case 5
                    R.Name = "Server Manager"
                Case 4

                Case 3
                    R.Name = "Service Provider Manager"
                Case 2
                    R.Name = "Company Manager"
                Case 1
                    R.Name = "Reviewer"

                Case 0
                    R.Name = "Scanner"
            End Select
        Next

    End Sub



    Public Function CXgetScanResults(ByVal scanID As Long) As CxPortal.CxWSResponceScanResults
        CXgetScanResults = New CxPortal.CxWSResponceScanResults
        CXgetScanResults = CxProxy.GetResultsForScan(sessionID, scanID)
    End Function

    Public Function CXgetAvailResultStates() As Collection
        CXgetAvailResultStates = New Collection

        Dim initCxProxy = New CxPortal.CxPortalWebServiceSoapClient()

        Dim C As CxPortal.CxWSResponseResultStateList
        C = CxProxy.GetResultStateList(sessionID)

        For Each rS In C.ResultStateList
            CXgetAvailResultStates.Add(rS.ResultName)
        Next
    End Function

    Public Sub CXsetPostScanActions()

    End Sub

    Public Sub CXgetGroups(ByRef allGroups As List(Of CxSDKns.Group), Optional ByVal forceREFRESH As Boolean = False)
        Static alreadyGotGroups As Boolean = False

        If alreadyGotGroups = True And forceREFRESH = False Then
            addLOG("Already loaded projects")
            Exit Sub
        End If

        CxSDKProxy = New CxSDKns.CxSDKWebServiceSoapClient

        addLOG("Retrieving list Of Teams/Groups..")
        allGroups = New List(Of CxSDKns.Group)

        Dim resP As CxSDKns.CxWSResponseGroupList

        resP = CxSDKProxy.GetAssociatedGroupsList(sessionID)
        If resP.IsSuccesfull = False Then
            addLOG("Error: Could not load Team Data - " + resP.ErrorMessage)
            Exit Sub
        Else
            For Each G In resP.GroupList
                allGroups.Add(G)
            Next
        End If

        addLOG(Trim(Str(allGroups.Count)) + " groups loaded")

        alreadyGotGroups = True
        '        For Each G In allGroups
        '        TextBox1.Text += G.Team.GroupName + " - " + G.Team.ID + " - " + G.Team.Guid + vbCrLf
        '        Next

        '        addLOG(Trim(Str(allGroups.Count)) + " groups loaded")




    End Sub


    Public Function CXrunScan(ByRef scanArgs As CxSDKns.CliScanArgs) As CxSDKns.CxWSResponseRunID
        Dim resP As New CxSDKns.CxWSResponseRunID
        CxSDKProxy = New CxSDKns.CxSDKWebServiceSoapClient()

        resP = CxSDKProxy.Scan(sdkID, scanArgs)
        If resP.IsSuccesfull = True Then
            addLOG("SCAN Successful - RUN ID =" + Str(resP.RunId) + ", PROJ ID =" + Str(resP.ProjectID))
        Else
            addLOG("ERROR: Scan Failed - " + resP.ErrorMessage)
        End If

        Return resP
    End Function

    Public Function CXaddGroup(ByRef parentGUID$, ByRef fullPath$) As String
        CXaddGroup = "Unidentified Failure"

        Dim resP As CxPortal.CxWSBasicRepsonse

        Dim LD(0) As CxPortal.CxWSLdapGroupMapping
        '        Dim LDg As New CxPortal.CxWSLdapGroup

        '        With LDg
        '        .DN = ""
        '        .Name = ""
        '        End With
        '       .LdapGroup = LDg
        '       .LdapServerId = 0
        '       End With



        resP = CxProxy.CreateNewTeam(sessionID, parentGUID, stripToFilename(fullPath), LD)



        If resP.IsSuccesfull = True Then CXaddGroup = "True" Else CXaddGroup = resP.ErrorMessage

    End Function

    Public Function getScansInQueue() As CxPortal.CxWSResponseExtendedScanStatus()
        CxProxy = New CxPortal.CxPortalWebServiceSoapClient

        Dim cxScans As New CxPortal.CxWSResponseScanStatus
        Dim cxS As New CxPortal.CxWSResponseExtendedScanStatusArray

        cxS = CxProxy.GetScansStatuses(sessionID)

        Dim SS() As CxPortal.CxWSResponseExtendedScanStatus

        SS = cxS.statusArr

        Return SS
    End Function

    Public Function cancelScanID(runID As String) As String
        Dim resP As CxPortal.CxWSBasicRepsonse

        resP = CxProxy.CancelScan(sessionID, runID)

        If resP.IsSuccesfull = False Then
            cancelScanID = resP.ErrorMessage.ToString
        Else
            cancelScanID = resP.IsSuccesfull.ToString
        End If

    End Function

    Public Function CXaddUser(ByRef U As CxPortal.UserData, Optional ByVal useLDAP As Boolean = False, Optional ByVal useSAML As Boolean = False) As String
        CXaddUser = "Unidentified Failure"

        CxProxy = New CxPortal.CxPortalWebServiceSoapClient

        Dim resP As CxPortal.CxWSBasicRepsonse

        If useSAML = True Then
            U.Password = "" 'this doesn't matter, but being rejected if not set
            'U.UserName = "SAML\" + U.UserName
            resP = CxProxy.AddNewUser(sessionID, U, CxPortal.CxUserTypes.SAML)
            GoTo wasSAML
        End If

        If useLDAP = True Then
            U.Password = "" 'this doesn't matter, but being rejected if not set
            resP = CxProxy.AddNewUser(sessionID, U, CxPortal.CxUserTypes.LDAP)
        Else
            resP = CxProxy.AddNewUser(sessionID, U, CxPortal.CxUserTypes.Application)
        End If

wasSAML:

        addLOG("User Add: " + U.UserName + "-" + U.LastName + "," + U.FirstName + " Success: " + (CStr(resP.IsSuccesfull)) + " - " + resP.ErrorMessage)
        If resP.IsSuccesfull = True Then CXaddUser = "True" Else CXaddUser = resP.ErrorMessage

    End Function




    Public Function CXgetResultStates(ByRef RS As CxPortal.ResultState()) As String
        CXgetResultStates = "True"

        Static alreadyLoaded = False

        If alreadyLoaded = True Then Exit Function

        Dim resP As CxPortal.CxWSResponseResultStateList
        resP = CxProxy.GetResultStateList(sessionID)

        If resP.IsSuccesfull = True Then
            RS = resP.ResultStateList
        Else
            CXgetResultStates = resP.ErrorMessage
        End If


    End Function



    Public Function CXgetCategories(ByRef CQ As CxPortal.CxQueryCategory()) As String
        CXgetCategories = "True"

        Static alreadyLoaded = False

        If alreadyLoaded = True Then Exit Function

        Dim resP As CxPortal.CxWSResponseQueriesCategories
        resP = CxProxy.GetQueriesCategories(sessionID)

        If resP.IsSuccesfull = True Then
            CQ = resP.QueriesCategories
        Else
            CXgetCategories = resP.ErrorMessage
        End If

    End Function

    Public Function CXgetVulnCats(ByRef CG As CxPortal.CxWSQueryGroup()) As String
        CXgetVulnCats = "True"

        Static alreadyLoaded = False

        If alreadyLoaded = True Then Exit Function

        Dim resP As CxPortal.CxQueryCollectionResponse

        resP = CxProxy.GetQueryCollection(sessionID)

        If resP.IsSuccesfull = True Then
            CG = resP.QueryGroups
        Else
            CXgetVulnCats = resP.ErrorMessage
        End If

    End Function

    Public Function CXgetScanVulns(ByRef CG As CxPortal.CxWSQueryVulnerabilityData(), ByRef scanID As Long) As String
        CXgetScanVulns = "True"

        Static alreadyLoaded = False

        If alreadyLoaded = True Then Exit Function

        Dim resP As CxPortal.CxWSResponceQuerisForScan

        resP = CxProxy.GetQueriesForScan(sessionID, scanID)

        If resP.IsSuccesfull = True Then
            CG = resP.Queries
        Else
            CXgetScanVulns = resP.ErrorMessage
        End If

    End Function

    Public Function CXgetCustomFields(ByRef CF As CxPortal.CxWSCustomField()) As String
        CXgetCustomFields = "True"

        Dim cFields As CxPortal.CxWSResponseCustomFields
        Static alreadyLoaded = False

        If alreadyLoaded = True Then Exit Function

        cFields = CxProxy.GetCustomFields(sessionID)

        If cFields.IsSuccesfull = False Then
            CXgetCustomFields = cFields.ErrorMessage
        Else
            CF = cFields.fieldsArray()
        End If
        alreadyLoaded = True

    End Function

    Public Function CXgetCustomFieldID(ByVal cfName$) As Long
        CXgetCustomFieldID = 0

        Static cFields As CxPortal.CxWSResponseCustomFields
        Static alreadyLoaded = False

        If alreadyLoaded = False Then cFields = CxProxy.GetCustomFields(sessionID)

        For Each F In cFields.fieldsArray
            If LCase(F.Name) = LCase(cfName) Then
                Return F.Id
                Exit Function
            End If
        Next

    End Function

    Public Function CXgetUserRoleData(ByRef U As CxPortal.UserData, Optional ByVal setFP$ = "", Optional ByVal setDELETEopt$ = "", Optional ByVal setEDITresults$ = "") As CxPortal.CxWSRoleWithUserPrivileges
        CXgetUserRoleData = New CxPortal.CxWSRoleWithUserPrivileges

        CXgetUserRoleData.ID = U.RoleData.ID

        Dim CrudITEMS As New CxPortal.CxWSItemAndCRUD

        Dim CrudACTIONS As New CxPortal.CxWSEnableCRUDAction

        Dim enumLIST As New CxPortal.CxWSCrudEnum

        '        ? enumLIST.Create
        '        Create {0}
        '? enumLIST.Delete
        '        Delete {1}
        '? enumLIST.Investigate
        '        Investigate {5}
        '? enumLIST.Run
        '        Run {4}
        '? enumLIST.View
        '        View {3}
        '? enumLIST.Update
        '        Update {2}


        Select Case U.RoleData.ID
            Case 0 'scanner
                CXgetUserRoleData.Name = "Scanner"

            Case 1 'reviewer
                CXgetUserRoleData.Name = "Reviewer"

        End Select

        Dim K As Integer = 1


    End Function

    Public Sub CXgetUsers(ByRef allUsers As CxPortal.CxWSResponseUserData, Optional ByVal forceREFRESH As Boolean = False)
        Static alreadyGotUsers As Boolean = False

        If alreadyGotUsers = True And forceREFRESH = False Then
            addLOG("Already loaded users")
            Exit Sub
        End If

        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()
        Dim CxSDKProxy = New CxSDKns.CxSDKWebServiceSoapClient()

        allUsers = New CxPortal.CxWSResponseUserData

        addLOG("Retrieving list of users..")
        allUsers = CxProxy.GetAllUsers(sessionID)

        If allUsers.IsSuccesfull = False Then
            addLOG("ERROR: Could not pull Users - " + allUsers.ErrorMessage)
            Exit Sub
        End If

        addLOG(Trim(Str(allUsers.UserDataList.Count)) + " users loaded")
        alreadyGotUsers = True

    End Sub

    Public Sub CXgetXML(CR As getReportArgs)
        '        On Error Resume Next
        Dim fileExt$ = CR.fileExt
        Dim fileType$ = CR.rptType
        Dim rptNDX As Integer = 3

        If LCase(fileType) = "xml" Then fileExt = ".xml"
        If LCase(fileType) = "pdf" Then
            fileExt = ".pdf"
            rptNDX = 0
        End If
        If LCase(fileType) = "csv" Then
            fileExt = ".csv"
            rptNDX = 2
        End If

        If Len(fileExt) = 0 Then Exit Sub


        Dim SRsdk As New CxSDKns.CxWSCreateReportResponse

        Dim RR As New CxSDKns.CxWSReportRequest
        RR.ScanID = CR.reportID
        RR.Type = rptNDX

        SRsdk = CxSDKProxy.CreateScanReport(sessionID, RR)
        Dim reportID As Long = SRsdk.ID

        Dim checkRPT As New CxSDKns.CxWSReportStatusResponse

        Do Until checkRPT.IsReady = True Or checkRPT.IsFailed = True
            checkRPT = CxSDKProxy.GetScanReportStatus(sessionID, reportID)
            System.Threading.Thread.Sleep(250)

            If checkRPT.IsReady = True Then
                Dim SR As New CxSDKns.CxWSResponseScanResults
                SR = CxSDKProxy.GetScanReport(sessionID, reportID)

                If SR.IsSuccesfull = True Then
                    File.WriteAllBytes(CR.fileName, SR.ScanResults)
                    '                Dim D As CxSDKns.
                    RaiseEvent reportCompleted(CR)
                Else
                    Debug.Print("ERROR: " + SR.ErrorMessage)
                End If
            End If

        Loop

    End Sub

    Public Sub CXgetProjectsDisplayData(ByRef projectS As List(Of CxPortal.ProjectDisplayData), Optional ByVal forceREFRESH As Boolean = False)
        Static alreadyGotProjects As Boolean = False

        If alreadyGotProjects = True And forceREFRESH = False Then
            addLOG("Already loaded projects")
            Exit Sub
        End If

        projectS = New List(Of CxPortal.ProjectDisplayData)
        CxProxy = New CxPortal.CxPortalWebServiceSoapClient


        Dim resP As CxPortal.CxWSResponseProjectsDisplayData = CxProxy.GetProjectsDisplayData(sessionID) '
        If resP.IsSuccesfull = True Then
            projectS = resP.projectList.ToList
            Call addLOG("Total # of Projects: " + Str(projectS.Count))
        Else
            Call addLOG("ERROR: Could not load Projects - " + resP.ErrorMessage)
        End If

        alreadyGotProjects = resP.IsSuccesfull
    End Sub



    Public Sub CXgetProjectsWithScans(ByRef projectS As CxPortal.CxWSResponseProjectsScansList, Optional ByVal forceREFRESH As Boolean = False)
        Static alreadyGotProjects As Boolean = False

        If alreadyGotProjects = True And forceREFRESH = False Then
            addLOG("Already loaded projects")
            Exit Sub
        End If

        projectS = New CxPortal.CxWSResponseProjectsScansList
        CxProxy = New CxPortal.CxPortalWebServiceSoapClient

        projectS = CxProxy.GetProjectsWithScans(sessionID)
        If projectS.IsSuccesfull = True Then
            Call addLOG("Total # of Projects with Scans: " + Str(projectS.projects.Count))
        Else
            Call addLOG("ERROR: Could not load Projects with Scans - " + projectS.ErrorMessage)
        End If

        alreadyGotProjects = projectS.IsSuccesfull
    End Sub

    Public Function CXgetEngines() As CxPortal.CxWSResponseEngineServers
        CXgetEngines = New CxPortal.CxWSResponseEngineServers

        CXgetEngines = CxProxy.GetEngineServers(sessionID)

        If CXgetEngines.IsSuccesfull = False Then
            addLOG("Could not obtain Engine Information")
        End If

    End Function

    Public Sub CXgetScans(ByRef allScans As CxSDKns.CxWSResponseScansDisplayData, Optional ByVal forceREFRESH As Boolean = False, Optional ByRef errorMsg$ = "")
        Static alreadyGotScans As Boolean = False


        On Error GoTo errorMessage

        If alreadyGotScans = True And forceREFRESH = False Then
            addLOG("Already loaded scans")
            Exit Sub
        End If

        allScans = New CxSDKns.CxWSResponseScansDisplayData
        CxSDKProxy = New CxSDKns.CxSDKWebServiceSoapClient() ' = New CxSDKproxy.CxPortalWebServiceSoapClient


        allScans = CxSDKProxy.GetScansDisplayDataForAllProjects(sessionID)

        If allScans.IsSuccesfull = True Then
            Call addLOG("Total # of Scans: " + Str(allScans.ScanList.LongCount))
            alreadyGotScans = True
        Else
            Call addLOG("ERROR: Could not load scan data - " + allScans.ErrorMessage)
            errorMsg = allScans.ErrorMessage
        End If

        Exit Sub
errorMessage:
        errorMsg = "Error during CXgetScans: " & ErrorToString()

    End Sub

    Public Function CXgetFailedScans(ByRef allFailed As CxPortal.CxWSResponseFailedScansDisplayData) As String
        CxProxy = New CxPortal.CxPortalWebServiceSoapClient


        allFailed = New CxPortal.CxWSResponseFailedScansDisplayData
        allFailed = CxProxy.GetFailedScansDisplayData(sessionID)


        If allFailed.IsSuccesfull = False Then
            CXgetFailedScans = allFailed.ErrorMessage
        Else
            CXgetFailedScans = "TRUE"
        End If


    End Function
    Public Sub CXgetProjectScansDisplayData(ByRef allScans As CxPortal.CxWSResponseProjectScannedDisplayData, Optional ByVal forceREFRESH As Boolean = False, Optional ByRef errorMsg$ = "")
        Static alreadyGotScans As Boolean = False

        On Error GoTo errorMessage

        If alreadyGotScans = True And forceREFRESH = False Then
            addLOG("Already loaded scans")
            Exit Sub
        End If

        allScans = New CxPortal.CxWSResponseProjectScannedDisplayData
        CxProxy = New CxPortal.CxPortalWebServiceSoapClient


        allScans = CxProxy.GetProjectScannedDisplayData(sessionID)

        If allScans.IsSuccesfull = True Then
            Call addLOG("Total # of Scans: " + Str(allScans.ProjectScannedList.LongCount))
            alreadyGotScans = True
        Else
            Call addLOG("ERROR: Could not load scan data - " + allScans.ErrorMessage)
            errorMsg = allScans.ErrorMessage
        End If

        Exit Sub
errorMessage:
        errorMsg = "Error during CXgetProjectScans: " & ErrorToString()

    End Sub

    Public Function CXsetUserActivationState(ByVal userID As Integer, ByVal activateUser As Boolean) As Boolean
        CxProxy = New CxPortal.CxPortalWebServiceSoapClient()

        Return CxProxy.SetUserActivationState(sessionID, userID, activateUser).IsSuccesfull

    End Function

    Public Function CXeditUser(ByRef U As CxPortal.UserData) As String
        CXeditUser = "Unidentified Error"
        ' On Error GoTo errorcatch

        '        addLOG("Entering cxedituser " + U.ID.ToString + " " + U.Email)

        Dim R As CxPortal.CxWSBasicRepsonse
        CxProxy = New CxPortal.CxPortalWebServiceSoapClient
        R = CxProxy.UpdateUserData(sessionID, U)

        If R.IsSuccesfull = True Then
            addLOG("Submitted User " + U.UserName + ": " + U.LastName + "," + U.FirstName)
            CXeditUser = "True"
        Else
            addLOG("ERROR: CxServer rejected user " + U.UserName + " " + R.ErrorMessage)
            CXeditUser = R.ErrorMessage
        End If

errorcatch:

        '        CxProxy.UpdateUserGroups(sessionID, U.ID,)
    End Function

    Public Function CXeditUserGroups(ByRef U As CxPortal.UserData, ByRef unSubscribed() As CxPortal.Group) As String
        CXeditUserGroups = "Unidentified Failure"

        Dim R As CxPortal.CxWSBasicRepsonse
        CxProxy = New CxPortal.CxPortalWebServiceSoapClient
        R = CxProxy.UpdateUserGroups(sessionID, U.ID, unSubscribed, U.GroupList, U.RoleData)

        If R.IsSuccesfull = True Then
            addLOG("Updated Groups for " + U.UserName + ": " + U.LastName + "," + U.FirstName)
            CXeditUserGroups = "True"
        Else
            addLOG("ERROR: Could not update groups for " + U.UserName + " " + R.ErrorMessage)
            CXeditUserGroups = R.ErrorMessage
        End If
        '        CxProxy.UpdateUserGroups(sessionID, U.ID,)
    End Function

    Public Function CXdeleteUser(ByVal uID As Long) As Boolean
        Dim CxProxy = New CxPortal.CxPortalWebServiceSoapClient()

        Dim resP As CxPortal.CxWSBasicRepsonse

        resP = CxProxy.DeleteUser(sessionID, uID)

        If resP.IsSuccesfull = True Then
            addLOG("User " + Trim(Str(uID)) + " deleted")
            Return True
        Else
            addLOG("ERROR: Could not delete User " + Trim(Str(uID)))
            Return False
        End If
    End Function


    Public Function CXgetProjectConfig(ByVal projectID As Long) As CxPortal.ProjectConfiguration
        CXgetProjectConfig = New CxPortal.ProjectConfiguration
        Dim respProjConfig As CxPortal.CxWSResponseProjectConfig
        respProjConfig = CxProxy.GetProjectConfiguration(sessionID, projectID)

        If respProjConfig.IsSuccesfull = True Then
            CXgetProjectConfig = respProjConfig.ProjectConfig
            addLOG("Retrieved Project Config for " + Str(projectID))
        Else
            addLOG("ERROR: Could not load Proj Config " + respProjConfig.ErrorMessage)
        End If
    End Function

    Public Sub CXgetStateList(ByRef allStates As List(Of CxPortal.ResultState))
        Static alreadyLoaded As Boolean = False

        If alreadyLoaded = True Then
            addLOG("Already loaded result states")
            Exit Sub
        End If

        Dim R As New CxPortal.CxWSResponseResultStateList
        R = CxProxy.GetResultStateList(sessionID)

        allStates = New List(Of CxPortal.ResultState)

        For Each ST In R.ResultStateList
            allStates.Add(ST)
        Next
        alreadyLoaded = True
    End Sub


    Public Function lastPresetDetails() As CxPortal.CxPresetDetails
        lastPresetDetails = New CxPortal.CxPresetDetails

        Dim allPresets As New CxPortal.CxWSResponsePresetList
        Call CXgetPresetList(allPresets, True)

        Dim lastPreset As New CxPortal.Preset
        Dim lastPresetID As Long = 0

        For Each P In allPresets.PresetList
            If P.ID > lastPresetID Then
                lastPresetID = P.ID
                lastPreset = P
            End If
        Next

        If lastPresetID > 0 Then
            Call getPresetDetails(lastPreset.ID, lastPresetDetails)
        End If
    End Function


    Public Function importPreset(ByRef PresetDetails As CxPortal.CxPresetDetails) As String
        Dim lastPresetObj As New CxPortal.CxPresetDetails
        lastPresetObj = lastPresetDetails()

        With PresetDetails
            .id = lastPresetObj.id + 1
            If .id < 100000 Then .id = 100000
            .isUserAllowToDelete = lastPresetObj.isUserAllowToDelete
            .isUserAllowToUpdate = lastPresetObj.isUserAllowToUpdate
            .isPublic = lastPresetObj.isPublic
            .owningteam = lastPresetObj.owningteam
            .IsDuplicate = lastPresetObj.IsDuplicate
            .owner = lastPresetObj.owner
        End With

        Dim createNewPreset = New CxPortal.CxWSResponsePresetDetails

        createNewPreset = CxProxy.CreateNewPreset(sessionID, PresetDetails)

        If createNewPreset.IsSuccesfull = True Then
            Return "True"
        Else
            Return createNewPreset.ErrorMessage
        End If

    End Function



    Public Function getPresetDetails(presetID As Long, ByRef PresetDetails As CxPortal.CxPresetDetails) As String
        PresetDetails = New CxPortal.CxPresetDetails
        Dim pQueries As New CxPortal.CxWSResponsePresetDetails

        pQueries = CxProxy.GetPresetDetails(sessionID, presetID)

        If pQueries.IsSuccesfull = True Then
            PresetDetails = pQueries.preset
            getPresetDetails = "True"
        Else
            getPresetDetails = pQueries.ErrorMessage
        End If

    End Function


    Public Sub CXgetPresetList(ByRef allPresets As CxPortal.CxWSResponsePresetList, Optional ByVal forceREFRESH As Boolean = False)
        Static alreadyLoaded As Boolean = False

        If alreadyLoaded = True And forceREFRESH = False Then
            addLOG("Already loaded presets")
            Exit Sub
        End If

        CxProxy = New CxPortal.CxPortalWebServiceSoapClient

        '  Dim allPresets As CxPortal.
        allPresets = New CxPortal.CxWSResponsePresetList
        allPresets = CxProxy.GetPresetList(sessionID)


        If allPresets.IsSuccesfull = True Then
            addLOG("Retrieved " + Str(allPresets.PresetList.Count) + " Presets")
        Else
            addLOG("ERROR: Could not load presets - " + allPresets.ErrorMessage)
        End If
        alreadyLoaded = True



    End Sub
    Private Sub addLOG(ByVal a$)
        On Error Resume Next
        Call Module1.addLOG(a)


    End Sub

    Public Function CXgetCWE(ByVal cID As Integer) As CxPortal.CxWSResponseShortQueryDescription

        Return CxProxy.GetQueryShortDescription(sessionID, cID)

    End Function


End Class

Public Class CxDerivedData
    'derived in batch
    Public vulnActions As List(Of vulnAction)
    Public daysToResolution As Integer 'days
    Public daysToAcknowledgement As Integer
    Public numDaysSinceZeroDay As Integer
    '    Public numScansSinceZeroDay As Integer

    Public firstScanID As Long
    Public lastScanID As Long
    Public firstDTAppeared As DateTime
    Public lastDTAppeared As DateTime

    Public userActivityCount As Integer

    Public fixedScanID As Long

    Public currStatus$
    Public currState$
    Public currSev$
    Public latestComment$

    Public numScansObserved As Integer
End Class


Public Class batchOfScans
    Public uniqueSimID As Collection
    Public numResolved As Collection
    Public filesIncluded As Collection
    Public projScans As New List(Of projectScanBuilder)
    Public projInfo As CxPortal.CxWSResponseBasicProjectData

    Public beginLOC As Long
    Public lastLOC As Long

    Public Event lifeCycleIdentified(ByVal similarityID As Long)

    Public Function getAllOfSim(ByRef similarityID As Long, ByRef vList As List(Of CxResult)) As List(Of CxResult)
        getAllOfSim = New List(Of CxResult)
        For Each V In vList
            If V.Flow.similarityID = similarityID Then
                getAllOfSim.Add(V)
            End If
        Next
    End Function

    Public Sub getLifecycle(lcArgs As lifeCycleArgs)
        Dim sevLvl$ = ""
        Dim similarityID As Long
        Dim addLifeCycleToV As CxResult
        ' sevlvl, addlifecycletov similarityid
        With lcArgs
            sevLvl$ = .sevLvl
            similarityID = .similarityID
            addLifeCycleToV = .V
        End With

        If addLifeCycleToV.alreadyDerived = True Then
            Exit Sub
        End If

        Dim newLifecycle = New CxDerivedData

        Dim foundINbatch As Boolean = False
        Dim foundINscan As Boolean = False

        Dim confirmedFIXED As Boolean = False


        Dim newV As New CxResult

        newLifecycle.vulnActions = New List(Of vulnAction)

        Dim numAppearances As Integer = 0
        Dim daysSinceFirstScan As Integer = 0

        Dim numScan As Integer = 0


        For Each P In projScans
            If numScan = 0 Then Me.beginLOC = P.ScanData.lineSoFcodE
            If numScan = projScans.Count - 1 Then Me.lastLOC = P.ScanData.lineSoFcodE
            numScan += 1
            foundINscan = False
            For Each V In P.getResults(,,, addLifeCycleToV.CxNamE, sevLvl)
                With newLifecycle
                    If V.Flow.similarityID <> similarityID Then GoTo nextVuln
                    numAppearances += 1
                    .numScansObserved = numAppearances
                    foundINscan = True
                    newV = V


                    ' mostRECENT = CStr(P.ScanData.scanFinished)
                    'here is VULN of matching similarity
                    If foundINbatch = False Then
                        'first time finding it
                        foundINbatch = True
                        .firstDTAppeared = P.ScanData.scanFinished
                        .firstScanID = P.ScanData.scanID
                        .vulnActions.Add(New vulnAction("ZeroDay", .firstDTAppeared, "Status changed to New", "System", 0))
                    End If

                    .lastDTAppeared = P.ScanData.scanFinished
                    .lastScanID = P.ScanData.scanID


                    If confirmedFIXED = True Then
                        confirmedFIXED = False
                        newLifecycle.vulnActions.Add(New vulnAction("Vuln Reappeared", P.ScanData.scanFinished, "Vulnerability Returned", "System", daysSinceFirstScan))
                        newLifecycle.fixedScanID = 0
                        addLifeCycleToV.statuS = "Recurring"
                        addLifeCycleToV.stateNDX = newV.stateNDX
                    End If



dontCompareToLast:

                End With
nextVuln:
            Next
            If foundINscan = False And newLifecycle.firstScanID <> 0 And newLifecycle.fixedScanID = 0 Then
                confirmedFIXED = True
                newLifecycle.fixedScanID = P.ScanData.scanID
                newLifecycle.vulnActions.Add(New vulnAction("Resolved", P.ScanData.scanFinished, "Status changed to Fixed", "System", DateDiff("d", newLifecycle.firstDTAppeared, P.ScanData.scanFinished)))
                newLifecycle.daysToResolution = DateDiff("d", newLifecycle.firstDTAppeared, newLifecycle.lastDTAppeared)
                addLifeCycleToV.statuS = "Fixed"
                addLifeCycleToV.stateNDX = 99
                'here, add it to most recent scan    lastScanID as "FIXED"
            End If
        Next

        If confirmedFIXED = False Then
            addLifeCycleToV.statuS = newV.statuS
            addLifeCycleToV.stateNDX = newV.stateNDX
        End If
        addLifeCycleToV.severityNDX = newV.severityNDX
        addLifeCycleToV.assignedUser = newV.assignedUser


        addLifeCycleToV.deriveD = newLifecycle

        ' here parse comments and add as actions
        Call commentsLifecycle(addLifeCycleToV, newV)

        With addLifeCycleToV
            If .statuS <> "Fixed" Then
                .deriveD.daysToResolution = DateDiff("d", newLifecycle.firstDTAppeared, Now)
                If .deriveD.daysToAcknowledgement = 0 Then .deriveD.daysToAcknowledgement = DateDiff("d", newLifecycle.firstDTAppeared, Now)
            End If
        End With


        addLifeCycleToV.alreadyDerived = True
        RaiseEvent lifeCycleIdentified(similarityID)
    End Sub



    Private Sub commentsLifecycle(ByRef V As CxResult, ByRef mostRecentV As CxResult)
        Dim a$ = ""
        a$ = mostRecentV.commentS

        If Len(a) = 0 Then Exit Sub
        Dim commentS As Object = Split(a, vbCrLf)

        Dim commentLoop As Integer = 0
        Dim typeUserAction$ = ""
        Dim cD$ = ""
        Dim c$ = ""
        Dim d$ = ""
        Dim useR$ = ""

        Dim srcH As Integer = 0

        ' With mostRecentV
        ' V.severityNDX = .severityNDX
        ' V.Severity = .Severity
        ' V.stateNDX = .stateNDX
        ' V.assignedUser = .assignedUser
        ' End With
        Dim ageAtTimeOfAction = 0

        With V.deriveD

            For commentLoop = 0 To UBound(commentS)
                '          admin adminm CxUtil, [Sunday, February 19, 2017 3:37:40 PM]: Changed status To Confirmed&#xD;&#xA;admin adminm CxUtil, [Sunday, February 19, 2017 2: 50:22 PM]: Changed status to Confirmed 
                '     admin adminm CxUtil, [Monday, May 15, 2017 5:42:33 PM]: Some miscellaneous comment&#xD;&#xA;admin adminm CxUtil, [Monday, May 15, 2017 5:41:45 PM]: Assigned to mhorty&#xD;&#xA;admin adminm CxUtil, [Monday, May 15, 2017 5:41:38 PM]: Changed severity to Medium&#xD;&#xA;admin adminm CxUtil, [Monday, May 15, 2017 5:41:27 PM]: Changed status to Urgent&#xD;&#xA;admin adminm CxUtil, [Sunday, February 19, 2017 3:37:40 PM]: Changed status to Confirmed&#xD;&#xA;admin adminm CxUtil, [Sunday, February 19, 2017 2:50:22 PM]: Changed status to Confirmed"
                c = Replace(commentS(commentLoop), vbCrLf, "")
                cD = Mid(c, InStr(c, "[") + 1)
                cD = Mid(cD, 1, InStr(cD, "]") - 1)
                d$ = ""

                useR = stripLastWord(Mid(c, 1, InStr(c, ",") - 1))

                srcH = InStr(c, "Changed status to ")
                If srcH Then
                    typeUserAction = "State Change"
                    d$ = "State changed to " + Mid(c, srcH + 18)
                End If
                srcH = InStr(c, "Assigned to ")
                If srcH Then
                    typeUserAction = "User Assignment"
                    d$ = "Assigned to " + Mid(c, srcH + 12)
                End If
                srcH = InStr(c, "Changed severity to ")
                If srcH Then
                    typeUserAction = "Severity Change"
                    d$ = "Severity changed to " + Mid(c, srcH + 20)
                End If

                If d$ = "" Then
                    ' must be a comment
                    typeUserAction = "Made Comment"
                    d = Mid(c, InStr(c, "]") + 3)
                    .latestComment = d
                End If

                ageAtTimeOfAction = DateDiff("d", .firstDTAppeared, CDate(cD))

                If commentLoop = UBound(commentS) Then
                    .vulnActions.Add(New vulnAction("First Acknowledgement", CDate(cD), typeUserAction, useR, ageAtTimeOfAction))
                    .daysToAcknowledgement = DateDiff("d", .firstDTAppeared, CDate(cD))
                End If
                .vulnActions.Add(New vulnAction(typeUserAction, CDate(cD), d, useR, ageAtTimeOfAction))
                .userActivityCount += 1

            Next
            .vulnActions.Sort(Function(x, y) x.actionDate.CompareTo(y.actionDate))

        End With
        a$ = ""

    End Sub

    Private Function returnNDXofScanID(ByRef scanID As Long, ByRef allScans As CxSDKns.CxWSResponseScansDisplayData) As Long
        Dim K As Long = 0
        returnNDXofScanID = 0
        For K = 0 To allScans.ScanList.LongCount - 1
            If scanID = allScans.ScanList(K).ScanID Then
                returnNDXofScanID = K + 1
            End If
        Next
    End Function
    Public Sub New(ByRef proJ As CxPortal.CxWSResponseBasicProjectData, ByRef allScans As CxSDKns.CxWSResponseScansDisplayData, Optional ByRef recentOnly As Boolean = False, Optional ByRef maxDays As Integer = 0, Optional ByRef maxScans As Integer = 0)
        'Dim projScans As New List(Of projectScanBuilder)
        Dim scanNDX As Long = 0
        Dim numScans As Integer = 0

        uniqueSimID = New Collection

        Me.projInfo = proJ


        For K = proJ.scans.Count - 1 To 0 Step -1
            If recentOnly = True And K > 0 Then GoTo skipThisScan
            scanNDX = returnNDXofScanID(proJ.scans(K).ID, allScans)
            If maxDays > 0 And DateDiff("d", CXconvertDT(allScans.ScanList(scanNDX - 1).FinishedDateTime), Date.Now) > maxDays Then GoTo skipThisScan
            If Dir(addSlash(CurDir) + "ScanData\" + Trim(Str(proJ.ID)) + "_" + Trim(Str(allScans.ScanList(scanNDX - 1).ScanID)) + ".xml") = "" Then GoTo skipThisScan

            If scanNDX = 0 Then GoTo skipThisScan

            scanNDX -= 1
            numScans += 1
            If maxScans <> 0 And numScans > maxScans Then Exit Sub

            Dim XB As New projectScanBuilder
            XB.setToXML(addSlash(CurDir) + "ScanData\" + Trim(Str(proJ.ID)) + "_" + Trim(Str(allScans.ScanList(scanNDX).ScanID)) + ".xml")

            XB.ScanData.scanFinished = CXconvertDT(allScans.ScanList(scanNDX).FinishedDateTime)
            XB.ScanData.scanQueued = CXconvertDT(allScans.ScanList(scanNDX).QueuedDateTime)
            XB.ScanData.riskLevel = allScans.ScanList(scanNDX).RiskLevelScore
            XB.ScanData.engineName = allScans.ScanList(scanNDX).ServerName


            Me.projScans.Add(XB)
            'Dim PP As List(Of CxResult)
skipThisScan:
        Next
    End Sub

End Class

Public Class projectScanBuilder
    Public ScanData As projectScan

    Public ResultTypes As Collection
    Public ResultCategories As Collection

    Public useOrigSeverity As Boolean
    Public honorFPsettings As Boolean

    Public Sub New()
        useOrigSeverity = False
        honorFPsettings = True

    End Sub


    Public Function numScanResultsBySev(ByVal sevNDX As Integer) As Long
        numScanResultsBySev = 0
        If sevNDX = 0 Then Exit Function

        For Each Q In ScanData.QueryResults
            For Each R In Q.Results
                If R.FalsePositive = True And honorFPsettings = True Then GoTo nextVuln
                If useOrigSeverity = True Then
                    If sevNDX - 1 = Q.origSeverityNDX Then numScanResultsBySev += 1
                Else
                    If sevNDX - 1 = R.severityNDX Then numScanResultsBySev += 1
                End If
nextVuln:
            Next

        Next
    End Function

    Public Function numScanResults() As Long
        numScanResults = 0
        Static alreadyCalc As Boolean = False
        If alreadyCalc = True Then Exit Function

        numScanResults = 0
        For Each Q In ScanData.QueryResults
            For Each R In Q.Results
                If R.FalsePositive = True And honorFPsettings = True Then GoTo nextVULN
                numScanResults += 1
nextVULN:
            Next
        Next
        alreadyCalc = True

    End Function

    Public Function getResults(Optional ByVal ofSevNDX As Integer = 0, Optional ByVal ofStateNDX As Integer = 0, Optional ByVal ofStatus$ = "", Optional ByVal ofVulnType$ = "", Optional ByVal ofSevStrings$ = "all") As List(Of CxResult)
        getResults = New List(Of CxResult)

        'status = new/recurring/resolved
        'severit = high/med/low as defined by Cx (if useOrigSeverity=true) or User post-processing (if =false)

        Dim newResult As CxResult
        Dim rSevNDX As Integer

        Dim skipVuln As Boolean
        Dim idNum As Integer = 0

        For Each Q In ScanData.QueryResults
            If Len(ofVulnType) And LCase(ofVulnType) <> LCase(Q.namE) Then GoTo nextQuery
            idNum = 0
            For Each R In Q.Results
                If Len(ofStatus) And ofStatus <> R.statuS Then GoTo nextVuln
                If ofStateNDX And ofStateNDX - 1 <> R.stateNDX Then GoTo nextVuln

                If honorFPsettings = True And R.FalsePositive = True Then GoTo nextVuln

                If useOrigSeverity = True Then rSevNDX = R.CxOrigSeverityNDX Else rSevNDX = R.severityNDX
                If ofSevNDX And ofSevNDX - 1 <> rSevNDX Then GoTo nextVuln

                If ofSevStrings <> "all" Then
                    Select Case rSevNDX
                        Case 3
                            If InStr(LCase(ofSevStrings), "high") = 0 Then skipVuln = True
                        Case 2
                            If InStr(LCase(ofSevStrings), "med") = 0 Then skipVuln = True
                        Case 1
                            If InStr(LCase(ofSevStrings), "low") = 0 Then skipVuln = True
                        Case 0
                            If InStr(LCase(ofSevStrings), "info") = 0 Then skipVuln = True
                    End Select
                    If skipVuln = True Then
                        skipVuln = False
                        GoTo nextVuln
                    End If
                End If

                idNum += 1
                newResult = New CxResult
                newResult = R
                newResult.Flow.iD = idNum
                getResults.Add(newResult)
nextVuln:
            Next
nextQuery:
        Next

    End Function

    Public Function numResults(Optional ByVal ofSevNDX As Integer = 0, Optional ByVal ofStateNDX As Integer = 0, Optional ByVal ofStatus$ = "", Optional ByVal ofVulnType$ = "") As Long
        numResults = 0

        Dim gR = getResults(ofSevNDX, ofStateNDX, ofStatus, ofVulnType)
        numResults = gR.Count

        gR = Nothing
    End Function


    Public Sub setToXML(ByVal fileN$)
        If Dir(fileN) = "" Then Exit Sub

        ScanData = New projectScan

        Dim xDoc As XDocument = XDocument.Load(fileN)

        Dim K As Integer
        K = 1

        ResultTypes = New Collection
        ResultCategories = New Collection

        For Each X As XElement In xDoc.Elements("CxXMLResults")

            With ScanData
                .teamFullPath = X.Attribute("TeamFullPathOnReportDate").Value
                .lineSoFcodE = Val(X.Attribute("LinesOfCodeScanned").Value)
                .initiatorName = X.Attribute("InitiatorName").Value
                .owneR = X.Attribute("Owner").Value
                .scanID = Val(X.Attribute("ScanId").Value)
                .projectID = Val(X.Attribute("ProjectId").Value)
                .projectName = X.Attribute("ProjectName").Value
                .scanStart = CDate(X.Attribute("ScanStart").Value)
                .rptCreatedTime = CDate(X.Attribute("ReportCreationTime").Value)
                .preseT = X.Attribute("Preset").Value
                .scanTime = X.Attribute("ScanTime").Value
                .numFiles = Val(X.Attribute("FilesScanned").Value)
                .cxVersion = X.Attribute("CheckmarxVersion").Value
                .scanComments = X.Attribute("ScanComments")
                If X.Attribute("ScanType").Value = "Full" Then .fullScan = True Else .fullScan = False
                .sourceOrigin = X.Attribute("SourceOrigin").Value
                If X.Attribute("Visibility").Value = "Public" Then .privateScan = False Else .privateScan = True
                .QueryResults = New List(Of CxQuery)
            End With

            For Each QR As XElement In X.Elements("Query")
                Dim Q As New CxQuery
                With Q
                    .namE = QR.Attribute("name").Value
                    .cwE = Val(QR.Attribute("cweId").Value)
                    .languagE = QR.Attribute("Language").Value
                    .queryID = Val(QR.Attribute("id").Value)
                    .queryGRP = QR.Attribute("group")
                    .queryPath = QR.Attribute("QueryPath").Value
                    .origSeverityNDX = Val(QR.Attribute("SeverityIndex").Value)
                    .origSeveritY = QR.Attribute("Severity").Value

                    If grpNDX(ResultTypes, .namE) = 0 Then ResultTypes.Add(.namE)
                    'check for null on category
                End With
                Q.Results = New List(Of CxResult)

                For Each P As XElement In QR.Elements("Result")
                    Dim R As New CxResult
                    With R
                        .assignedUser = P.Attribute("AssignToUser").Value
                        .commentS = P.Attribute("Remark").Value
                        .deepLink = P.Attribute("DeepLink").Value
                        .NodeId = P.Attribute("NodeId").Value
                        .Severity = P.Attribute("Severity").Value
                        .severityNDX = Val(P.Attribute("SeverityIndex").Value)
                        .startFileN = P.Attribute("FileName").Value
                        '''''''''pull this from 1st node don't need it here'''.startLOC = Val(P.Attribute("Line").Value)
                        .statuS = P.Attribute("Status").Value
                        .stateNDX = Val(P.Attribute("state").Value)

                        If P.Attribute("FalsePositive").Value = "False" Then .FalsePositive = False Else .FalsePositive = True

                        .CxCWE = Q.cwE
                        .CxNamE = Q.namE
                        .CxQueryGRP = Q.queryGRP
                        .CxorigSeveritY = Q.origSeveritY
                        .CxLanguagE = Q.languagE
                        .CxOrigSeverityNDX = Q.origSeverityNDX
                        .CxQueryPath = Q.queryPath
                        .CxQID = Q.queryID
                    End With
                    R.Flow = New dFlow


                    For Each PATH As XElement In P.Elements("Path") ' QR.Elements("Result").Elements("Path")

                        Dim F As New dFlow
                        With F
                            .pathID = Val(PATH.Attribute("PathId").Value)
                            .resultID = CLng(PATH.Attribute("ResultId").Value)
                            .similarityID = CLng(PATH.Attribute("SimilarityId").Value)
                            .iD = 0
                            ' .similarityID = .similarityNum + Trim(Str(.pathID))
                        End With
                        F.Nodes = New List(Of dfNode)

                        For Each N As XElement In PATH.Elements("PathNode")
                            Dim pNode As New dfNode
                            With pNode
                                .fileN = N.Element("FileName").Value
                                .loC = Val(N.Element("Line").Value)
                                .namE = N.Element("Name").Value
                                .nodeID = Val(N.Element("NodeId").Value)
                                pNode.codeSnippet = New dfSnippet

                                For Each SNIP As XElement In N.Elements("Snippet").Elements("Line")
                                    pNode.codeSnippet.loC = Val(SNIP.Element("Number").Value)
                                    pNode.codeSnippet.codE = SNIP.Element("Code").Value
                                Next

                            End With
                            F.Nodes.Add(pNode)
                        Next
                        R.Flow = F
                    Next

                    Q.Results.Add(R)
                Next

                ScanData.QueryResults.Add(Q)
            Next
        Next
        K = 1

    End Sub
End Class

Public Class teamProps
    Private teamPath$
    Private teamName$
    Public teamGUID$
    Public childGUIDs As Collection
    Private numChildren As Long
    Private numProj As Long
    Private numScan As Long
    Private numProjChildren As Long
    Private numScanChildren As Long
    Public allUsers As List(Of CxPortal.UserData)
    Public inheritedUsers As List(Of CxPortal.UserData)
    Public allProj As List(Of CxPortal.ProjectDisplayData)
    Public allScan As List(Of CxSDKns.ScanDisplayData)
    Public childreN As List(Of teamProps)

    Public Sub New(ByRef G As CxSDKns.Group, ByRef S As List(Of CxSDKns.ScanDisplayData), ByRef P As List(Of CxPortal.ProjectDisplayData), ByRef U As List(Of CxPortal.UserData))
        ' define all users, projects and scans of this team
        allUsers = New List(Of CxPortal.UserData)
        allUsers = U
        allProj = New List(Of CxPortal.ProjectDisplayData)
        allProj = P
        allScan = New List(Of CxSDKns.ScanDisplayData)
        allScan = S
        inheritedUsers = New List(Of CxPortal.UserData)
        teamPath = G.GroupName
        teamName = stripToFilename(G.GroupName)
        teamGUID = G.ID

        childGUIDs = New Collection
    End Sub


    Public ReadOnly Property Name As String
        Get
            Return teamName
        End Get
    End Property
    Public ReadOnly Property CxType As String
        Get
            Dim lvL As Integer
            lvL = countChars(teamPath, "\")
            Dim a$ = ""
            Select Case lvL
                Case 0
                    a = "0 [CxRoot]"
                Case 1
                    a = "1 [SP]"
                Case 2
                    a = "2 [Company]"
                Case 3
                    a = "3 [Team]"
            End Select

            If lvL > 3 Then a = Trim(Str(lvL)) + " [Nested]"
            Return a
        End Get

    End Property


    Public ReadOnly Property Users As Integer
        Get
            Return allUsers.Count
        End Get
    End Property

    Public ReadOnly Property Members As Integer
        Get
            Return allUsers.Count + inheritedUsers.Count
        End Get
    End Property

    Public ReadOnly Property Projs_P As Integer
        Get
            Return allProj.Count
        End Get
    End Property
    Public ReadOnly Property Scans_S As Integer
        Get
            Return allScan.Count
        End Get
    End Property
    Public ReadOnly Property Nests As Integer
        Get
            Return childGUIDs.Count
        End Get
    End Property
    Public Property Nests_P As Integer
        Set(ByVal numChilds As Integer)
            numProjChildren = numChilds
        End Set
        Get
            Return numProjChildren
        End Get
    End Property
    Public Property Nests_S As Integer
        Set(ByVal numChilds As Integer)
            numScanChildren = numChilds
        End Set
        Get
            Return numScanChildren
        End Get
    End Property
    Public ReadOnly Property LastScan As String
        Get
            If Me.allScan.Count Then
                Return CXconvertDT(Me.allScan(0).FinishedDateTime).ToShortDateString
            Else
                Return "N/A"
            End If
        End Get

    End Property

    Public ReadOnly Property FullPath As String
        Get
            Return teamPath
        End Get
    End Property

End Class

Public Class userDisplayRows
    Private usrName$
    Private eMailaddy$
    Private fullN$
    Private ISactivE As Boolean
    Public grpName$
    Private numOfScans As Integer
    Private dateCreateD As Date
    Private dateExpireD As Date
    Private dateLastLogin As Date
    Public userID As Long

    Public Sub New(ByRef U As CxPortal.UserData, ByVal groupName$, ByVal numScansInGroup As Integer)

        With U
            usrName = .UserName
            eMailaddy = .Email
            fullN = .FirstName + " " + .LastName
            ISactivE = .IsActive
            dateCreateD = .DateCreated
            dateExpireD = DateAdd(DateInterval.Day, CDbl(.willExpireAfterDays), Today.Date)
            dateLastLogin = .LastLoginDate
            userID = .ID
        End With
        grpName = groupName

    End Sub

    Public Sub addScan()
        numOfScans += 1
    End Sub

    Public ReadOnly Property Username As String
        Get
            Return usrName
        End Get
    End Property
    Public ReadOnly Property EMail As String
        Get
            Return eMailaddy
        End Get
    End Property
    Public ReadOnly Property FullName As String
        Get
            Return fullN
        End Get
    End Property
    Public ReadOnly Property Active As Boolean
        Get
            Return ISactivE
        End Get
    End Property
    Public ReadOnly Property Business As String
        Get
            Dim a$ = grpName
            If a = "CxServer" Then Return "CxServer"
            a = Replace(a, "CxServer", "")
            a = Replace(a, stripToFilename(a), "")
            Return a
        End Get
    End Property
    Public ReadOnly Property Teams As String
        Get
            Return stripToFilename(grpName)
        End Get
    End Property
    Public ReadOnly Property Scans As Integer
        Get
            Return numOfScans
        End Get
    End Property
    Public ReadOnly Property Created As Date
        Get
            Return dateCreateD
        End Get
    End Property
    Public ReadOnly Property Expired As Date
        Get
            Return dateExpireD
        End Get
    End Property
    Public ReadOnly Property LastLogin As Date
        Get
            Return dateLastLogin
        End Get
    End Property


End Class

Public Class CxResultAPI
    Public Flow As dFlow
    Public NodeId$
    Public startFileN$
    Public FalsePositive As Boolean
    Public assignedUser$
    Public Severity$
    Public stateNDX As Integer
    Public severityNDX As Integer
    Public statuS$ ''As Integer
    Public commentS$
    Public deepLink$

    Public CxCWE As Long
    Public CxNamE$
    Public CxQueryGRP$
    Public CxorigSeveritY$
    Public CxLanguagE$
    Public CxOrigSeverityNDX As Integer
    Public CxQueryPath$

    Public scanID As Long

End Class



Public Class projDisplayProps
    Private projName$
    'Private teamName$
    Private projID As Long
    Private numScans As Integer
    Private lastLOC As Long
    Private lastRiskLvl As Integer
    Private lastHighRes As Long
    Private lastMedRes As Long
    Private lastLowRes As Long
    Private groupGUID$
    Public teamName$

    Public Sub New(ByRef P As CxPortal.CxWSResponseBasicProjectData, ByRef S As CxSDKns.ScanDisplayData, ByVal groupName$, ByVal gGuid$)
        projName = P.Name
        numScans = P.scans.Count
        '        Dim lastScanID As Long =
        projID = P.ID

        teamName = groupName
        groupGUID = gGuid

        With S
            lastHighRes = .HighSeverityResults
            lastMedRes = .MediumSeverityResults
            lastLowRes = .LowSeverityResults
            lastLOC = .LOC
            lastRiskLvl = .RiskLevelScore
            '            teamName = "Have to get this"
        End With
    End Sub

    Public ReadOnly Property ID As Long
        Get
            Return projID
        End Get
    End Property
    Public ReadOnly Property Name As String
        Get
            Return projName
        End Get
    End Property
    Public ReadOnly Property tlScans As Integer
        Get
            Return numScans
        End Get
    End Property
    Public ReadOnly Property LOC As Long
        Get
            Return lastLOC
        End Get
    End Property
    Public ReadOnly Property Score As Integer
        Get
            Return lastRiskLvl
        End Get
    End Property
    Public ReadOnly Property HighSev As Long
        Get
            Return lastHighRes
        End Get
    End Property
    Public ReadOnly Property MedSev As Long
        Get
            Return lastMedRes
        End Get
    End Property
    Public ReadOnly Property LowSev As Long
        Get
            Return lastLowRes
        End Get
    End Property
    Public ReadOnly Property Team As String
        Get
            Dim a$ = Replace(teamName, "CxServer", "")
            If a = "" Then a = "[CxRoot]"
            Return a
        End Get
    End Property


End Class

Public Class xlsImportArgs
    Public numItems As Integer
    Public big3D(,) As String
    Public fieldS As Collection
    Public fileN As String
End Class

Public Class rightClickArgs
    Public objectS As Collection
    Public currView As String
    Public action2Take As String
End Class

Public Class getReportArgs
    Public reportID As Long
    Public projectID As Long
    Public rptType$
    Public rptDate As DateTime
    Public fileName$
    Public fileExt$
    Public workingFolder$
    Public targetZIP$ 'this is highly inefficient but objects temp
    Public forDataImport As Boolean
End Class

Public Class bulkProjScanArgs
    '    Call CxWrap.CXgetXML(allProjIDs, recentFilesOnly, fType, maxAge)
    Public allProjIDs As Collection
    Public recentFilesOnly As Boolean
    Public fileType$
    Public maxAge As Integer
    Public zipFilename$
    Public forDataImport As Boolean
End Class

Public Class backgroundUserArgs
    Public U As CxPortal.UserData
    Public changeActiveState As Boolean
    Public addORedit$
    Public isLDAP As Boolean
    Public unSubscribed() As CxPortal.Group
End Class

Public Class backgroundUserActivationArgs
    Public matchOn$
    Public actioN$
    Public fileN$
    Public usrColl As Collection
End Class

Public Class backgroundScanArgs
    Public projName$
    Public projPreset$
    Public projTeam$
    Public folderPath$
    'maybe add more later
End Class

Public Class lifeCycleArgs
    Public V As CxResult
    Public similarityID As Long
    Public sevLvl$
End Class
Public Class reportingArgs
    Public rptName$ = ""
    Public booL1 As Boolean = False
    Public booL2 As Boolean = False
    Public s1$ = ""
    Public s2$ = ""
    Public s3$ = ""
    Public numeriC As Integer = 0
    Public numeriC2 As Integer = 0
    Public numeriC3 As Integer = 0
    Public someColl As Collection
End Class


Public Class projectScan
    Public QueryResults As List(Of CxQuery)
    Public lineSoFcodE As Long
    Public initiatorName$
    Public owneR$
    Public scanID As Long
    Public projectID As Long
    Public projectName$
    Public teamFullPath$
    Public scanStart As DateTime
    Public rptCreatedTime As DateTime
    Public preseT$
    Public scanTime$
    Public numFiles As Long
    Public cxVersion$
    Public scanComments$
    Public fullScan As Boolean
    Public sourceOrigin$
    Public privateScan As Boolean

    'derived from allScans()
    Public scanFinished As DateTime
    Public scanQueued As DateTime
    Public riskLevel As Integer
    Public engineName$
End Class

Public Class CxQuery
    Public Results As List(Of CxResult)
    Public queryID As Long
    'Public categorY$
    Public cwE As Long
    Public namE$
    Public queryGRP$
    Public origSeveritY$
    Public languagE$
    Public origSeverityNDX As Integer
    Public queryPath$
End Class

Public Class CxResult
    Public Flow As dFlow
    Public NodeId$
    Public startFileN$
    Public FalsePositive As Boolean
    Public assignedUser$
    Public Severity$
    Public stateNDX As Integer
    Public severityNDX As Integer
    Public statuS$ ''As Integer
    Public commentS$
    Public deepLink$

    Public CxCWE As Long
    Public CxNamE$
    Public CxQueryGRP$
    Public CxorigSeveritY$
    Public CxLanguagE$
    Public CxOrigSeverityNDX As Integer
    Public CxQueryPath$
    Public CxQID As Long

    Public deriveD As CxDerivedData
    Public alreadyDerived As Boolean
End Class

Public Class vulnAction
    Public actionType$
    Public actionDate As DateTime
    Public actionDesc$
    Public actionUser$
    Public ageAtAction As Integer
    'action types --
    'first acknowledgement
    'change status
    'mark as Not Exploitable
    'change severity
    'comment
    'resolved
    Public Sub New(typeOfAction$, dateOfAction As DateTime, actionDescription$, actionBy$, agE As Integer)
        Me.actionType = typeOfAction
        Me.actionDate = dateOfAction
        Me.actionDesc = actionDescription
        Me.actionUser = actionBy
        Me.ageAtAction = agE
    End Sub
End Class

Public Class dFlow
    Public Nodes As List(Of dfNode)
    Public resultID As Long
    Public pathID As Long 'path is unique within a scan
    Public similarityID As Long
    Public iD As Integer
End Class

Public Class dfNode

    Public fileN$
    Public loC As Long
    Public nodeID As Long
    Public codeSnippet As dfSnippet
    Public namE$

End Class
Public Class dfSnippet
    Public loC As Long
    Public codE$
End Class

Public Class CLIArgs
    Public matchOn$
    Public editCmd$ 'currently addgroups, subtractgroups, role
    Public newVal$
    Public uData$


End Class


'for sorting
Public Class AdvancedList(Of T)
    Inherits BindingList(Of T)
    Implements IBindingListView
    Protected Overrides ReadOnly Property IsSortedCore() As Boolean
        Get
            Return sorts IsNot Nothing
        End Get
    End Property
    Protected Overrides Sub RemoveSortCore()
        sorts = Nothing
    End Sub
    Protected Overrides ReadOnly Property SupportsSortingCore() As Boolean
        Get
            Return True
        End Get
    End Property
    Protected Overrides ReadOnly Property SortDirectionCore() As ListSortDirection
        Get
            Return If(sorts Is Nothing, ListSortDirection.Ascending, sorts.PrimaryDirection)
        End Get
    End Property
    Protected Overrides ReadOnly Property SortPropertyCore() As PropertyDescriptor
        Get
            Return If(sorts Is Nothing, Nothing, sorts.PrimaryProperty)
        End Get
    End Property
    Protected Overrides Sub ApplySortCore(ByVal prop As PropertyDescriptor, ByVal direction As ListSortDirection)
        Dim arr As ListSortDescription() = {New ListSortDescription(prop, direction)}
        ApplySort(New ListSortDescriptionCollection(arr))
    End Sub
    Private sorts As PropertyComparerCollection(Of T)
    Public Sub ApplySort(ByVal sortCollection As ListSortDescriptionCollection) Implements IBindingListView.ApplySort
        Dim oldRaise As Boolean = RaiseListChangedEvents
        RaiseListChangedEvents = False
        Try
            Dim tmp As New PropertyComparerCollection(Of T)(sortCollection)
            Dim items As New List(Of T)(Me)
            items.Sort(tmp)
            Dim index As Integer = 0
            For Each item As T In items
                SetItem(index, item)
                index += 1
            Next
            sorts = tmp
        Finally
            RaiseListChangedEvents = oldRaise
            ResetBindings()
        End Try
    End Sub
    Private Property IBindingListView_Filter() As String Implements IBindingListView.Filter
        Get
            Throw New NotImplementedException()
        End Get
        Set(ByVal value As String)
            Throw New NotImplementedException()
        End Set
    End Property
    Private Sub IBindingListView_RemoveFilter() Implements IBindingListView.RemoveFilter
        Throw New NotImplementedException()
    End Sub
    Private ReadOnly Property IBindingListView_SortDescriptions() As ListSortDescriptionCollection Implements IBindingListView.SortDescriptions
        Get
            Return sorts.Sorts
        End Get
    End Property
    Private ReadOnly Property IBindingListView_SupportsAdvancedSorting() As Boolean Implements IBindingListView.SupportsAdvancedSorting
        Get
            Return True
        End Get
    End Property
    Private ReadOnly Property IBindingListView_SupportsFiltering() As Boolean Implements IBindingListView.SupportsFiltering
        Get
            Return False
        End Get
    End Property
End Class

Public Class PropertyComparerCollection(Of T)
    Implements IComparer(Of T)
    Private ReadOnly m_sorts As ListSortDescriptionCollection
    Private ReadOnly comparers As PropertyComparer(Of T)()
    Public ReadOnly Property Sorts() As ListSortDescriptionCollection
        Get
            Return m_sorts
        End Get
    End Property
    Public Sub New(ByVal sorts As ListSortDescriptionCollection)
        If sorts Is Nothing Then
            Throw New ArgumentNullException("sorts")
        End If
        Me.m_sorts = sorts
        Dim list As New List(Of PropertyComparer(Of T))()
        For Each item As ListSortDescription In sorts
            list.Add(New PropertyComparer(Of T)(item.PropertyDescriptor, item.SortDirection = ListSortDirection.Descending))
        Next
        comparers = list.ToArray()
    End Sub
    Public ReadOnly Property PrimaryProperty() As PropertyDescriptor
        Get
            Return If(comparers.Length = 0, Nothing, comparers(0).[Property])
        End Get
    End Property
    Public ReadOnly Property PrimaryDirection() As ListSortDirection
        Get
            Return If(comparers.Length = 0, ListSortDirection.Ascending, If(comparers(0).Descending, ListSortDirection.Descending, ListSortDirection.Ascending))
        End Get
    End Property

    Private Function IComparer_Compare(ByVal x As T, ByVal y As T) As Integer Implements IComparer(Of T).Compare
        Dim result As Integer = 0
        For i As Integer = 0 To comparers.Length - 1
            result = comparers(i).Compare(x, y)
            If result <> 0 Then
                Exit For
            End If
        Next
        Return result
    End Function

End Class

Public Class PropertyComparer(Of T)
    Implements IComparer(Of T)
    Private ReadOnly m_descending As Boolean
    Public ReadOnly Property Descending() As Boolean
        Get
            Return m_descending
        End Get
    End Property
    Private ReadOnly m_property As PropertyDescriptor
    Public ReadOnly Property [Property]() As PropertyDescriptor
        Get
            Return m_property
        End Get
    End Property
    Public Sub New(ByVal [property] As PropertyDescriptor, ByVal descending As Boolean)
        If [property] Is Nothing Then
            Throw New ArgumentNullException("property")
        End If
        Me.m_descending = descending
        Me.m_property = [property]
    End Sub
    Public Function Compare(ByVal x As T, ByVal y As T) As Integer Implements IComparer(Of T).Compare
        ' todo; some null cases
        Dim value As Integer = Comparer.[Default].Compare(m_property.GetValue(x), m_property.GetValue(y))
        Return If(m_descending, -value, value)
    End Function
End Class

Public Class defaultSettings
    Public lciD As Long 'Language - when logging in and when adding users
    Public userExpire As String 'EOY = End of Year, EOL = License Expiration Date, or number (ie 365)
    Public pivotOrPlain As String 'Default Report format when pivot is an option
    Public scannerOrReviewer As String 'Default user type when adding
    Public defaultGroup As String 'Automatically add users to this group (no need to select)
    Public lastDayOfLicense As Date
End Class

Public Class dashInfo
    Public tlNumProjects As Long
    Public tlNumProjectsWithScans As Long
    Public tlNumUsers As Long
    Public tlNumActiveUsers As Long
    Public tlNumScans As Long
    Public licenseData As CxPortal.CxWSResponseServerLicenseData
    Public resultState(10) As Long
    Public resultSeverity(10) As Long
    Public expireSoon As Long
    Public totalLOC As Double
    Public tlNumSPs As Long
    Public tlNumCompanies As Long
    Public tlNumTeams As Long
End Class

Public Class vulnInfoAPI
    Public ProjectName As String
    Public VulnID As Long
    Public PathID As Long
    Public QueryID As Long
    Public ScanID As Long
    Public ProjectID As Long
    Public SimilarityID As String
    Public TeamName As String
    Public Initiator As String
    Public Origin As String
    Public LOC As Long
    Public EngineName As String
    Public State As String
    Public Severity As String
    Public Comment As String
    Public Status As String
    Public AssignedUser As String
    Public Vuln_CWE As Long
    Public Vuln_LangCategory As String
    Public Vuln_Description As String
    Public Language As String
    Public Orig_Severity As String
    Public RiskLevel As Integer
    Public QueuedDateTime As DateTime
    Public FinishedDateTime As DateTime
    Public ScanTime As String
    Public customFields As Collection
End Class
