Imports System.Security.Cryptography


Module modCommon
    Public defaultS As defaultSettings


    Public Function sevString(ByRef sevNum As Integer) As String
        sevString = ""
        Select Case sevNum
            Case 0
                sevString = "Info"
            Case 1
                sevString = "Low"
            Case 2
                sevString = "Medium"
            Case 3
                sevString = "High"
        End Select
    End Function


    Public Function grpNDX(ByRef C As Collection, ByRef a$, Optional ByVal caseSensitive As Boolean = True) As Integer
        Dim K As Integer
        grpNDX = 0

        If caseSensitive = False Then GoTo dontEvalCase

        For Each S In C
            K += 1
            If a = S Then
                grpNDX = K
                Exit Function
            End If
        Next
        Exit Function

dontEvalCase:
        For Each S In C
            K += 1
            If LCase(a) = LCase(S) Then
                grpNDX = K
                Exit Function
            End If
        Next


    End Function

    Public Function safeFilename(ByVal a) As String
        safeFilename = Replace(a, "\", "")
        safeFilename = Replace(safeFilename, "..", "")
    End Function

    Public Function CXconvertDT(ByRef CD As CxSDKns.CxDateTime) As DateTime
        On Error GoTo errorcatch

        Dim a$ = ""
        With CD
            a$ += Trim(Str(.Month)) + "/" + Trim(Str(.Day)) + "/" + Trim(Str(.Year))
            'theres the date
            a += " "
            a += Trim(Str(.Hour)) + ":" + Trim(Str(.Minute)) + ":" + Trim(Str(.Second))
        End With
        Return CDate(a)

errorcatch:
        'saw this when trying to convert dates from Yuval

        a$ = CD.ToString
        Return CDate(a)
    End Function

    Public Function CXconvertDTportal(ByRef CD As CxPortal.CxDateTime) As DateTime
        On Error GoTo errorcatch

        Dim a$ = ""
        With CD
            a$ += Trim(Str(.Month)) + "/" + Trim(Str(.Day)) + "/" + Trim(Str(.Year))
            'theres the date
            a += " "
            a += Trim(Str(.Hour)) + ":" + Trim(Str(.Minute)) + ":" + Trim(Str(.Second))
        End With
        Return CDate(a)

errorcatch:
        'saw this when trying to convert dates from Yuval

        a$ = CD.ToString
        Return CDate(a)
    End Function

    Public Function listNDX(ByRef C As List(Of String), ByRef a$) As Integer
        Dim K As Integer = 0
        listNDX = 0

        For K = 0 To C.Count - 1
            If C(K).ToString = a Then
                listNDX = K + 1
                Exit Function
            End If
        Next

    End Function


    Public Function arrNDX(ByRef A$(), ByRef matcH$) As Integer
        'returns 0 if not found, otherwise NDX + 1
        Dim K As Long
        arrNDX = 0
        For K = 0 To UBound(A)
            If Trim(Str(A(K))) = matcH Then
                arrNDX = K + 1
                Exit Function
            End If
        Next
    End Function

    Public Function removeExtraSpaces(a) As String
        removeExtraSpaces = ""
        If Len(a) = 0 Then Exit Function
        Dim lastSpace As Boolean = False

        Dim K As Integer = 0
        For K = 0 To Len(a) - 1
            If lastSpace = False Then
                removeExtraSpaces += Mid(a, K + 1, 1)
            Else
                If Mid(a, K + 1, 1) <> " " Then removeExtraSpaces += Mid(a, K + 1, 1)
            End If
            If Mid(a, K + 1, 1) = " " Then
                lastSpace = True
            End If
        Next

    End Function

    Public Function countChars(a$, chr2Count$) As Integer
        countChars = 0

        Dim K As Integer = 0
        For K = 0 To Len(a) - 1
            If Mid(a, K + 1, 1) = chr2Count Then countChars += 1
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

    Public Function addSlash(ByVal a$) As String
        addSlash = a
        If Len(a) = 0 Then Exit Function

        If Mid(a, Len(a), 1) <> "\" Then addSlash += "\"
    End Function

    Public Function getParentGroup(ByVal g$) As String
        Dim a$ = StrReverse(g)
        a = Mid(a, InStr(a, "\") + 1)
        Return StrReverse(a)
    End Function

    Public Function stripLastWord(ByVal g$) As String
        Dim a$ = StrReverse(g)
        a = Mid(a, InStr(a, " ") + 1)
        Return StrReverse(a)
    End Function

    Public Function getPDFappend(ByVal fileN$) As String
        'takes Cx filename used in calls to pull apart date section.. everything from -DD.M.YYYY-H.MM.SS.pdf- ---- cant assume '-' isnt used in filename
        'go to second - of reverse
        Dim a$ = StrReverse(fileN)
        Dim numDash As Integer = 0

        fileN = ""
        getPDFappend = ""

        If InStr(a, "-") = 0 Then Exit Function
        Dim K As Integer

        For K = 1 To Len(a)
            fileN += Mid(a, K, 1)
            If Mid(fileN, K, 1) = "-" Then numDash += 1
            If numDash = 2 Then Exit For
        Next

        getPDFappend = StrReverse(fileN)
    End Function

    Public Function assembleCollFromCLI(clI$) As Collection
        Dim C As New Collection
        ' takes windows dos-style dir output and makes sense of it for collection storage
        Dim tempStr$ = clI
        Dim K As Integer
        Do Until InStr(tempStr, "  ") = 0
            K = InStr(tempStr, "  ")
            If Len(Mid(tempStr, 1, K - 1)) Then C.Add(Mid(tempStr, 1, K - 1))
            tempStr = Replace(tempStr, Mid(tempStr, 1, K - 1) + "  ", "")
            'Debug.Print(tempStr)
        Loop
        tempStr = LTrim(tempStr)
        C.Add(Mid(tempStr, 1, InStr(tempStr, " ") - 1))
        tempStr = Replace(tempStr, Mid(tempStr, 1, InStr(tempStr, " ") - 1), "")
        C.Add(LTrim(tempStr))
        Return C

    End Function

    Public Function CSVtoCOLL(ByRef csV$) As Collection
        CSVtoCOLL = New Collection

        Dim splitCHR$ = ","
        If InStr(csV, splitCHR) = 0 Then splitCHR = ";"


        Dim longS = Split(csV, splitCHR)

        Dim K As Integer
        For K = 0 To UBound(longS)
            CSVtoCOLL.Add(longS(K))
        Next

    End Function


    Public Function CSVFiletoCOLL(ByRef csV$) As Collection
        CSVFiletoCOLL = New Collection
        If Dir(csV) = "" Then Exit Function

        'use file
        Dim FF As Integer
        FF = FreeFile()

        FileOpen(FF, csV, OpenMode.Input)

        Do Until EOF(FF) = True
            CSVFiletoCOLL.Add(LineInput(FF))
        Loop
        FileClose(FF)

    End Function

    Public Function argPROP(proP$, Optional ByVal preserveCase As Boolean = False) As String
        Dim a$ = ""
        argPROP = ""

        If Len(proP) = 0 Then Exit Function

        proP = LCase(proP)
        For Each arg In My.Application.CommandLineArgs
            If preserveCase = False Then a = LCase(arg) Else a = arg
            If InStr(a, "=") = 0 Then GoTo nextArg
            'If InStr(arg, "hello") Then addLOG("CONSOLE:ARGPROP SEES " + arg)

            If proP = Mid(a, 1, InStr(a, "=") - 1) Then
                argPROP = Replace(a, proP + "=", "")
            End If
nextArg:
        Next

    End Function

    Public Function filePROP(fileN$, proP$, Optional ByVal preserveCase As Boolean = False) As String
        filePROP = ""
        If Dir(fileN) = "" Then Exit Function

        If Len(proP) = 0 Then Exit Function

        Dim FF As Integer = FreeFile()

        FileOpen(FF, fileN, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)

        Dim a$ = ""

        a = LineInput(FF)
        Do Until a = "" Or EOF(FF) = True
            If InStr(a, "=") = 0 Then GoTo nextLine

            If UCase(proP) = Mid(a, 1, InStr(a, "=") - 1) Then
                filePROP = Replace(a, proP + "=", "")
            End If
nextLine:
            a = LineInput(FF)
        Loop

        If Len(a) = 0 Then GoTo closeHere

        If UCase(proP) = Mid(a, 1, InStr(a, "=") - 1) Then
            filePROP = Replace(a, proP + "=", "")
        End If

closeHere:

        FileClose(FF)
    End Function

    Public Function xlsDataType(dType$) As String
        xlsDataType = "nonefound"
        Select Case dType
            Case "bigint", "int", "numeric", "float"
                xlsDataType = "Numeric"
            Case "datetime", "datetime2"
                xlsDataType = "DateTime"
            Case "date"
                xlsDataType = "Date"
            Case "time"
                xlsDataType = "Time"
            Case "bit"
                xlsDataType = "Boolean"
            Case "ntext", "nvarchar", "nchar", "varchar", "image", "uniqueidentifier", "real"
                xlsDataType = "String"
        End Select
        If xlsDataType = "nonefound" Then
            Debug.Print("No Def: " + dType)
            xlsDataType = "String"
        End If
    End Function

    Public Function xlsColName(colNum As Integer) As String
        Dim d As Integer
        Dim m As Integer
        Dim name As String
        d = colNum
        name = ""
        Do While (d > 0)
            m = (d - 1) Mod 26
            name = Chr(65 + m) + name
            d = Int((d - m) / 26)
        Loop
        xlsColName = name
    End Function


    Public NotInheritable Class Simple3Des
        Private TripleDes As TripleDESCryptoServiceProvider

        Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()

            Dim sha1 As New SHA1CryptoServiceProvider

            ' Hash the key.
            Dim keyBytes() As Byte =
                System.Text.Encoding.Unicode.GetBytes(key)
            Dim hash() As Byte = sha1.ComputeHash(keyBytes)

            ' Truncate or pad the hash.
            ReDim Preserve hash(length - 1)
            Return hash
        End Function

        Sub New(ByVal key As String)
            ' Initialize the crypto provider.
            TripleDes = New TripleDESCryptoServiceProvider
            TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
            TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
        End Sub

        Private Function EncryptData(ByVal plaintext As String) As String

            ' Convert the plaintext string to a byte array. 
            Dim plaintextBytes() As Byte =
                System.Text.Encoding.Unicode.GetBytes(plaintext)

            ' Create the stream. 
            Dim ms As New System.IO.MemoryStream
            ' Create the encoder to write to the stream. 
            Dim encStream As New CryptoStream(ms,
                TripleDes.CreateEncryptor(),
                System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
            encStream.FlushFinalBlock()

            ' Convert the encrypted stream to a printable string. 


            Return Convert.ToBase64String(ms.ToArray)

        End Function

        Private Function DecryptData(ByVal encryptedtext As String) As String

            ' Convert the encrypted text string to a byte array. 
            Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

            ' Create the stream. 
            Dim ms As New System.IO.MemoryStream
            ' Create the decoder to write to the stream. 
            Dim decStream As New CryptoStream(ms,
                TripleDes.CreateDecryptor(),
                System.Security.Cryptography.CryptoStreamMode.Write)

            ' Use the crypto stream to write the byte array to the stream.
            decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
            decStream.FlushFinalBlock()

            ' Convert the plaintext stream to a string. 
            Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
        End Function

        Public Function Decode(cipher As String) As String
            Try
                Return DecryptData(cipher)
            Catch ex As CryptographicException
                Throw New Exception(ex.Message)
            End Try

        End Function

        Public Function Encode(txt As String) As String
            Try
                Return EncryptData(txt)
            Catch ex As CryptographicException
                Throw New Exception(ex.Message)
            End Try
        End Function

    End Class

End Module

