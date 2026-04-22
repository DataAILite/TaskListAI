Imports System.Data.SqlClient
Imports System.Data
Imports System.Math
Imports System.IO
Imports System.Web.UI.Page
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Imports System.Reflection
Imports MySql.Data.MySqlClient
Imports System.Xml
Public Module mFuncs

    Public Function CopyArray(ByVal objArray As Object) As Object
        Try
            Dim objRet As Object = Nothing
            If Not objArray Is Nothing Then
                Dim typeObj As Type = objArray.GetType

                'Object must be an array
                If typeObj.HasElementType Then
                    Dim typeElement As Type = typeObj.GetElementType
                    Dim i As Integer
                    Dim arr As Array = CType(objArray, Array)
                    Dim n As Integer = arr.Length
                    'Create a copy of the array (no values in it yet)
                    Dim arCopy As Array = Array.CreateInstance(typeElement, n)
                    Dim obj As Object
                    'Put copied values in copied array
                    For i = 0 To n - 1
                        obj = arr.GetValue(i)
                        If Not obj Is Nothing Then
                            obj = CopyObject(obj)
                            arCopy.SetValue(obj, i)
                        End If
                    Next
                    objRet = arCopy
                Else
                    objRet = CopyObject(objArray)
                End If
            End If
            Return objRet
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function CopyObject(ByVal objSource As Object) As Object
        Try
            Dim objRet As Object = Nothing
            If Not objSource Is Nothing Then
                Dim typeObj As Type = objSource.GetType
                Dim i As Integer

                'Only creates new object and copies fields/properties if objSource is a class
                'and not string 
                If typeObj.IsClass And typeObj.Name <> "String" Then
                    Dim Fields As FieldInfo() = typeObj.GetFields(BindingFlags.Public Or BindingFlags.Instance)
                    Dim Field As FieldInfo
                    Dim Props As PropertyInfo() = typeObj.GetProperties(BindingFlags.Public Or BindingFlags.Instance)
                    Dim Prop As PropertyInfo
                    Dim objConstructor As ConstructorInfo = typeObj.GetConstructor(BindingFlags.Public Or BindingFlags.Instance, Nothing, CallingConventions.Standard, Type.EmptyTypes, Nothing)

                    If Not objConstructor Is Nothing Then
                        objRet = objConstructor.Invoke(Nothing)
                        Dim obj As Object

                        If Not objRet Is Nothing Then
                            'Copy fields
                            If Not Fields Is Nothing Then
                                For i = 0 To Fields.Length - 1
                                    Field = Fields(i)
                                    Try
                                        obj = Field.GetValue(objSource)
                                        'If field is not a value type, make a copy of it too
                                        If Not obj Is Nothing AndAlso
                                           obj.GetType.IsClass AndAlso
                                           obj.GetType.Name <> "String" Then
                                            If Not obj.GetType.HasElementType Then
                                                obj = CopyObject(obj)
                                            Else
                                                obj = CopyArray(obj)
                                            End If
                                        End If
                                        Field.SetValue(objRet, obj)
                                    Catch ex As Exception
                                        'no processing
                                    End Try
                                Next
                            End If
                            'Copy properties
                            If Not Props Is Nothing Then
                                For i = 0 To Props.Length - 1
                                    Prop = Props(i)
                                    Dim Params As ParameterInfo()

                                    Try
                                        'Only copy read/write properties
                                        If Prop.CanRead And Prop.CanWrite Then
                                            Params = Prop.GetIndexParameters
                                            If Params.Length = 0 Then
                                                obj = Prop.GetValue(objSource, Nothing)
                                                Prop.SetValue(objRet, obj, Nothing)
                                            End If
                                        End If
                                    Catch ex As Exception
                                        'no processing
                                    End Try
                                Next
                            End If
                        Else
                            objRet = objSource
                        End If
                    Else
                        objRet = objSource
                    End If
                Else
                    objRet = objSource
                End If
            End If

            Return objRet
        Catch ex As Exception
            'Err.ProcessError(ex)
            Return Nothing
        End Try
    End Function
    Public Function Pieces(ByVal sStr As String, ByVal Delim As String) As Integer
        'This function returns the number of pieces in sStr delimited by
        'Delim.

        Dim i As Integer = 0
        Dim b As Integer = sStr.IndexOf(Delim)
        While b > -1
            i = i + 1
            b = sStr.IndexOf(Delim, b + Delim.Length)
        End While

        Return i + 1
    End Function

    Public Function Pieces(ByVal sStr As String, ByVal Delim As Char) As Integer
        'This function returns the number of pieces in sStr delimited by
        'Delim.
        Dim i As Integer = 0
        Dim b As Integer = sStr.IndexOf(Delim)

        While b > -1
            i = i + 1
            b = sStr.IndexOf(Delim, b + 1)
        End While

        Return i + 1
    End Function
    Public Function Piece(ByVal sStr As String, ByVal Delim As String,
                     Optional ByVal iStart As Integer = 0,
                     Optional ByVal iEnd As Integer = 0) As String
        'This function acts like the $Piece function in Mumps
        'Usage: sStr=Piece(sAnotherString,sDelimeter,iStartPiece,iEndPiece)

        Dim iPrev As Integer = 1
        Dim b As Integer
        Dim iPieces As Integer = Pieces(sStr, Delim)
        Dim sRetStr As String = ""
        Dim i As Integer

        If iStart < 1 Then iStart = 1
        If iEnd < iStart Then iEnd = iStart

        If iPieces = 1 Then
            If iStart = 1 Then sRetStr = sStr
        Else
            For i = 1 To iPieces
                b = sStr.IndexOf(Delim, iPrev - 1) + 1
                If b = 0 Then b = sStr.Length + 1
                If i = iStart Then
                    sRetStr = Mid(sStr, iPrev, b - iPrev)
                ElseIf i > iStart Then
                    sRetStr = sRetStr & Delim & Mid(sStr, iPrev, b - iPrev)
                End If
                If i = iEnd Then Exit For
                iPrev = b + Delim.Length
            Next i
        End If
        Return sRetStr
    End Function
    Public Function SetPiece(ByVal sStr As String, ByVal sPiece As String, ByVal sDelim As String, ByVal nPiece As Integer) As String
        Dim sRet As String
        Dim sS As String
        Dim nPieces As Integer


        sRet = sStr
        If nPiece > 0 Then
            nPieces = Pieces(sStr, sDelim)
            If nPiece > nPieces Then
                sS = StringOf(sDelim, nPiece - nPieces)
                sRet = sRet & sS
                nPieces = nPiece
            ElseIf nPieces = 1 Then
                sRet = sPiece
            End If

            If nPieces > 1 Then
                If nPiece = 1 Then
                    sRet = sPiece & sDelim & Piece(sStr, sDelim, 2, nPieces)
                ElseIf nPiece = nPieces Then
                    sRet = Piece(sStr, sDelim, 1, nPiece - 1) & sDelim & sPiece
                Else
                    sRet = Piece(sStr, sDelim, 1, nPiece - 1) & sDelim & sPiece & sDelim & Piece(sStr, sDelim, nPiece + 1, nPieces)
                End If
            End If
        End If
        Return sRet
    End Function

    Public Function StringOf(ByVal sS As String, ByVal n As Integer) As String
        Dim sb As New StringBuilder
        Dim i As Integer

        For i = 1 To n
            sb.Append(sS)
        Next

        Return sb.ToString
    End Function


    Public Function cleanTextFromRepeatedCommas(ByVal strText As String) As String
        Dim i, l, j, k As Integer
        j = 0
        k = 0
        Dim letter, newstr As String
        newstr = ""
        strText = Trim(strText)
        newstr = strText
        l = Len(strText)
        If l > 0 Then
            For i = 1 To l 'CHECK the beginning commas
                letter = Mid(strText, i, 1)
                If j = 0 Then
                    If (letter = ",") Then
                        newstr = Mid(strText, i + 1, l - i)
                    Else
                        j = i ' not "," found in i position
                        strText = Trim(newstr)
                        Exit For
                    End If
                End If
            Next
        End If
        l = Len(strText)
        If l > 0 Then
            For i = 0 To l - 1 'CHECK the ending commas
                letter = Mid(strText, l - i, 1)
                If k = 0 Then
                    If (letter = "," Or letter = " ") Then
                        newstr = Mid(strText, 1, l - i - 1)
                    Else
                        k = i 'not "," found in i position
                        strText = Trim(newstr)
                        l = Len(strText)
                        Exit For
                    End If
                End If
            Next
        End If
        j = 0
        k = 0
        'TODO CHECK the middle commas
        newstr = ""
        Replace(strText, " ", "")
        'CHECK the middle commas
        Replace(strText, ",,", ",", l)
        Return cleanText(strText)
    End Function

    Public Function cleanText(ByVal strText As String) As String
        'TODO more!!
        If strText Is Nothing OrElse strText = "" Then
            Return ""
        End If
        strText = strText.Replace("<%", "***")
        strText = strText.Replace("%>", "***")
        strText = strText.Replace("</", "***")
        strText = strText.Replace("/>", "***")
        strText = strText.Replace("'", "***")
        strText = strText.Replace("<>", "!=")
        strText = strText.Replace("<", "***less than***")
        strText = strText.Replace(">", "***more than***")
        Dim i, l As Integer
        Dim letter As String
        l = Len(strText)
        If l > 0 Then
            For i = 1 To l
                letter = Mid(strText, i, 1)
                If (letter = ">") Or (letter = "<") Or (letter = "%") Or (letter = "'") Or (letter = Chr(39)) Or (letter = Chr(10)) Or (letter = Chr(13)) Then
                    If (i = 1) Then
                        strText = " " & Right(strText, l - 1)
                    Else
                        If (i = l) Then
                            strText = Left(strText, l - 1)
                        Else
                            strText = Left(strText, i - 1) & " " & Right(strText, l - i)
                        End If
                    End If
                End If
            Next
        End If
        cleanText = Trim(strText)
    End Function
    Public Function cleanTextLight(ByVal strText As String) As String
        'TODO more!!
        If strText Is Nothing OrElse strText = "" Then
            Return ""
        End If
        strText = strText.Replace("<%", "***")
        strText = strText.Replace("%>", "***")
        strText = strText.Replace("</", "***")
        strText = strText.Replace("/>", "***")
        strText = strText.Replace("'", "***")
        strText = strText.Replace("<>", "!=")
        strText = strText.Replace("<", "***less than***")
        strText = strText.Replace(">", "***more than***")
        Dim i, l As Integer
        Dim letter As String
        l = Len(strText)
        If l > 0 Then
            For i = 1 To l
                letter = Mid(strText, i, 1)
                If (letter = ">") Or (letter = "<") Or (letter = "%") Or (letter = "'") Or (letter = Chr(39)) Then 'Or (letter = Chr(10)) Or (letter = Chr(13)) Then
                    If (i = 1) Then
                        strText = " " & Right(strText, l - 1)
                    Else
                        If (i = l) Then
                            strText = Left(strText, l - 1)
                        Else
                            strText = Left(strText, i - 1) & " " & Right(strText, l - i)
                        End If
                    End If
                End If
            Next
        End If
        Return strText.Trim
    End Function
    Public Function DeleteFiles(path As String, FileSpecs As String) As String
        Try
            Dim FileList As String() = Directory.GetFiles(path, FileSpecs)
            For Each f As String In FileList
                File.Delete(f)
            Next
        Catch ex As Exception
            Return ex.Message
        End Try
        Return String.Empty
    End Function
    Public Function cleanFile(ByVal fpath As String, Optional ByVal clean As Boolean = False) As String
        'TODO!! XXE
        'show external url
        Try
            Dim filetext As String
            'read file to filetext
            filetext = File.ReadAllText(fpath)
            'look for possible urls
            Dim urladdress As String = String.Empty
            Dim k As Integer
            k = filetext.IndexOf("http://")
            If k > 0 AndAlso filetext.Substring(k, 29) <> "http://schemas.microsoft.com/" Then
                If clean = False Then
                    Return "XXE " & filetext.Substring(k, 29) & "... in the file " & fpath & " , upload is prohibited for security reason!!"
                End If

            End If
            k = filetext.IndexOf("https://")
            If k > 0 Then
                If clean = False Then
                    Return "XXE " & filetext.Substring(k, 29) & "... in the file " & fpath & " , upload is prohibited for security reason!!"
                End If

            End If
            k = filetext.IndexOf("file:///")
            If k > 0 Then
                If clean = False Then
                    Return "XXE " & filetext.Substring(k, 29) & "... in the file " & fpath & " , upload is prohibited for security reason!!"
                End If
            End If
            'clean and save filetext in fpath
            filetext = filetext.Replace("http://", "...").Replace("https://", "...").Replace("file:///", "...")
            File.Delete(fpath)
            File.WriteAllText(fpath, filetext)
            Return fpath
        Catch ex As Exception
            Return ex.Message
        End Try
        Return ""
    End Function
    Public Function cleanSQL(ByVal strText As String) As String
        'TODO!!
        'Dim i, l As Integer
        'Dim letter As String
        'l = Len(strText)
        ''If l > 0 Then
        ''    For i = 1 To l
        ''        letter = Mid(strText, i, 1)
        ''        If (letter = ">") Or (letter = "<") Or (letter = "%") Or (letter = "'") Or (letter = Chr(39)) Then
        ''            If (i = 1) Then
        ''                strText = " " & Right(strText, l - 1)
        ''            Else
        ''                If (i = l) Then
        ''                    strText = Left(strText, l - 1)
        ''                Else
        ''                    strText = Left(strText, i - 1) & " " & Right(strText, l - i)
        ''                End If
        ''            End If
        ''        End If
        ''    Next
        ''End If
        cleanSQL = Trim(strText)
    End Function

    Public Function FormatAsHTML(ByVal comments As String) As String
        Dim textHTML As String
        Dim i As Integer
        textHTML = "<b><u>"
        For i = 1 To Len(comments)
            If Mid(comments, i, 1) = "|" Then
                textHTML = textHTML & "<br/><b><u>"
            ElseIf Mid(comments, i, 2) = vbCrLf Then
                textHTML &= "<br/>"
            Else
                If Mid(comments, i, 2) = ": " Then
                    textHTML = textHTML & ": </b></u>"
                ElseIf Mid(comments, i, 2) = "):" Then
                    textHTML = textHTML & ")</b></u>"
                Else
                    textHTML = textHTML & Mid(comments, i, 1)
                End If
            End If
        Next
        'textHTML = MakeLinks(textHTML)  'not allowed by security of .NET
        If textHTML.IndexOf("File attached: </b></u> ") >= 0 Then
            textHTML = FileAttachLink(textHTML)
        End If
        Return textHTML
    End Function
    Public Function MakeLinks(ByVal comments As String) As String
        Dim textHTML, tmp, addr As String
        Dim i, j As Integer
        textHTML = ""
        '<a href="Default.aspx">Log Off</a>
        For i = 1 To Len(comments)
            If Mid(comments, i, 7) = "http://" Then
                textHTML = textHTML & "<br><b><a href="""
                tmp = comments.Substring(i + 6)
                j = tmp.Length
                If tmp.IndexOf(" ") > 0 OrElse tmp.IndexOf("|") OrElse tmp.IndexOf("|") OrElse tmp.IndexOf(",") OrElse tmp.IndexOf("!") OrElse tmp.IndexOf(";") Then
                    If tmp.IndexOf(" ") > 0 Then j = Min(j, tmp.IndexOf(" "))
                    If tmp.IndexOf("|") > 0 Then j = Min(j, tmp.IndexOf("|"))
                    If tmp.IndexOf(",") > 0 Then j = Min(j, tmp.IndexOf(","))
                    If tmp.IndexOf("!") > 0 Then j = Min(j, tmp.IndexOf("!"))
                    If tmp.IndexOf(";") > 0 Then j = Min(j, tmp.IndexOf(";"))
                    addr = "http://" & tmp.Substring(0, j).Trim
                    i = i + 6 + j
                Else
                    addr = tmp
                    i = i + 6
                End If
                textHTML = textHTML & addr & """>" & addr & "</a> </b></u>"
                Continue For
            ElseIf Mid(comments, i, 8) = "https://" Then
                textHTML = textHTML & "<br><b><a href="""
                tmp = comments.Substring(i + 7)
                j = tmp.Length
                If tmp.IndexOf(" ") > 0 OrElse tmp.IndexOf("|") OrElse tmp.IndexOf("|") OrElse tmp.IndexOf(",") OrElse tmp.IndexOf("!") OrElse tmp.IndexOf(";") Then
                    If tmp.IndexOf(" ") > 0 Then j = Min(j, tmp.IndexOf(" "))
                    If tmp.IndexOf("|") > 0 Then j = Min(j, tmp.IndexOf("|"))
                    If tmp.IndexOf(",") > 0 Then j = Min(j, tmp.IndexOf(","))
                    If tmp.IndexOf("!") > 0 Then j = Min(j, tmp.IndexOf("!"))
                    If tmp.IndexOf(";") > 0 Then j = Min(j, tmp.IndexOf(";"))
                    addr = "https://" & tmp.Substring(0, j).Trim
                    i = i + 7 + j
                Else
                    addr = tmp
                    i = i + 7
                End If
                textHTML = textHTML & addr & """>" & addr & "</a> </b></u>"
                Continue For
            Else
                textHTML = textHTML & Mid(comments, i, 1)
            End If
        Next
        Return textHTML
    End Function
    Public Function FileAttachLink(ByVal comments As String) As String
        Dim textHTML, tmp, addr As String
        Dim i, j As Integer
        textHTML = "<b>"
        '<a href="Default.aspx">Log Off</a>
        For i = 1 To Len(comments)
            If Mid(comments, i, 24) = "File attached: </b></u> " Then
                textHTML = textHTML & "<br><b>File attached: <a href="""
                tmp = comments.Substring(i + 23)
                If tmp.IndexOf(" ") > 0 Then
                    j = tmp.IndexOf(" ")
                    addr = tmp.Substring(0, j).Trim
                    i = i + 23 + j
                Else
                    addr = tmp
                    i = i + 23
                End If
                textHTML = textHTML & addr & """>" & addr & "</a> </b></u>"
                Continue For
            Else
                textHTML = textHTML & Mid(comments, i, 1)
            End If
        Next
        Return textHTML
    End Function
    Public Function FormatAsHTMLsimple(ByVal comments As String) As String
        Dim textHTML As String
        Dim i As Integer
        textHTML = ""
        For i = 1 To Len(comments)
            If Mid(comments, i, 1) = "|" Then
                textHTML = textHTML & "<br>"
            Else
                textHTML = textHTML & Mid(comments, i, 1)
            End If

        Next
        FormatAsHTMLsimple = textHTML
    End Function

    Function SendHTMLEmail(ByVal attach As String, ByVal subj As String, ByVal textbody As String, ByVal towhom As String, ByVal fromwho As String) As String
        Dim ret As String = String.Empty
        Dim at As Attachment = Nothing

        If towhom = "" OrElse fromwho = "" Then
            ret = "Email has not been sent. No email addresses are found."
            Return ret
        End If
        Dim iMsg As MailMessage
        iMsg = New MailMessage
        iMsg.IsBodyHtml = False
        Dim addrs(1) As MailAddress
        Try
            'MailAddress validates email address format starting from .Net 4.0
            addrs(0) = New MailAddress(towhom)
            addrs(1) = New MailAddress(fromwho)
        Catch ex As Exception
            ret = ex.Message
            Return ret
        End Try
        iMsg.From = New MailAddress(fromwho)
        iMsg.To.Add(addrs(0))
        iMsg.Bcc.Add(addrs(1))
        iMsg.Subject = subj
        iMsg.Body = textbody
        If attach <> String.Empty Then
            For i As Integer = 1 To Pieces(attach, ",")
                at = New Attachment(Piece(attach, ",", i))
                iMsg.Attachments.Add(at)
            Next
        End If
        '----------------------------------------------------
        Try
            ' set up smtp client
            Dim smtpClnt As New SmtpClient()
            smtpClnt.Host = ConfigurationManager.AppSettings("SmtpCred").ToString
            Dim smtpemail As String = ConfigurationManager.AppSettings("smtpemail").ToString
            Dim smtpemailpass = ConfigurationManager.AppSettings("smtpemailpass").ToString
            smtpClnt.Credentials = New Net.NetworkCredential(smtpemail, smtpemailpass)
            smtpClnt.Port = 587
            smtpClnt.EnableSsl = True

            ' send the message  
            smtpClnt.Send(iMsg)
            iMsg = Nothing
            'inform user that message has sent   
            ret = "Email has been sent to: " & towhom & " with subject: " & subj
        Catch ex As Exception
            ret = "ERROR!! " & ex.Message
            'ret = SendHTMLEmailOrg(attach, subj, textbody, towhom, fromwho)
        End Try
        Return ret
    End Function
    Function SendHTMLEmailOrg(ByVal attach As String, ByVal subj As String, ByVal textbody As String, ByVal towhom As String, ByVal fromwho As String) As String
        'for cloud1
        Dim ret As String = String.Empty
        Dim at As Attachment = Nothing

        If towhom = "" OrElse fromwho = "" Then
            ret = "Email has not been sent. No email addresses are found."
            Return ret
        End If
        Dim iMsg As MailMessage
        iMsg = New MailMessage
        iMsg.IsBodyHtml = False
        Dim addrs(1) As MailAddress
        addrs(0) = New MailAddress(towhom)
        addrs(1) = New MailAddress(fromwho)
        iMsg.From = New MailAddress(fromwho)
        iMsg.To.Add(addrs(0))
        iMsg.Bcc.Add(addrs(1))
        iMsg.Subject = subj
        iMsg.Body = textbody
        If attach <> String.Empty Then
            For i As Integer = 1 To Pieces(attach, ",")
                at = New Attachment(Piece(attach, ",", i))
                iMsg.Attachments.Add(at)
            Next
        End If
        '----------------------------------------------------
        Try
            ' set up smtp client And credentials  
            Dim SmtpCred As String = ConfigurationManager.AppSettings("SmtpCred").ToString
            'Dim smtpClnt As New SmtpClient()
            Dim smtpClnt As New SmtpClient(SmtpCred)
            smtpClnt.UseDefaultCredentials = True
            smtpClnt.Port = 25
            ' send the message  
            smtpClnt.Send(iMsg)
            'inform user that message has sent   
            ret = "Email has been sent To: " & towhom & " with subject: " & subj
        Catch ex As Exception
            ret = ex.Message
        End Try
        Return ret
    End Function

    Public Function CleanDBNulls(ByRef dv3t As DataTable, Optional ByRef err As String = "") As DataTable
        'Dim coltype As String = ""
        Dim i, j As Integer
        Try
            For i = 0 To dv3t.Rows.Count - 1
                For j = 0 To dv3t.Columns.Count - 1
                    'coltype = dv3t.Columns(j).DataType.FullName
                    If IsDBNull(dv3t.Rows(i)(j)) Then
                        'dv3t.Rows(i)(j) = ""
                        If ColumnTypeIsNumeric(dv3t.Columns(j)) Then
                            dv3t.Rows(i)(j) = 0
                        ElseIf ColumnTypeIsDateTime(dv3t.Columns(j)) Then
                            dv3t.Rows(i)(j) = DateTime.MinValue
                        ElseIf ColumnTypeIsString(dv3t.Columns(j)) Then
                            dv3t.Rows(i)(j) = ""
                        Else
                            dv3t.Rows(i)(j) = Nothing
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            'err = coltype & " " & ex.Message
            err = err & " " & ex.Message
        End Try
        Return dv3t
    End Function

    Public Function ColumnTypeIsNumeric(ByVal col As DataColumn) As Boolean
        If (col.DataType.FullName = "System.Single" OrElse col.DataType.FullName = "System.Double" OrElse col.DataType.FullName = "System.Decimal" OrElse col.DataType.FullName = "System.Byte" OrElse col.DataType.FullName = "System.Int16" OrElse col.DataType.FullName = "System.Int32" OrElse col.DataType.FullName = "System.Int64") Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function ColumnTypeIsDateTime(ByVal col As DataColumn) As Boolean
        If (col.DataType.FullName = "System.DateTime" OrElse col.DataType.FullName = "System.Date" OrElse col.DataType.FullName = "System.Time") Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function ColumnTypeIsString(col As DataColumn) As Boolean
        If col.DataType.FullName = "System.String" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetTextFromFile(ByVal dir As String, ByVal filename As String, ByVal ext As String) As String
        Dim restext As String = String.Empty
        Dim ret As String = String.Empty
        Try
            Dim mytextStreamReader As StreamReader = New StreamReader(dir & filename & ext)
            restext = mytextStreamReader.ReadToEnd
            Return restext
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function CreateRoutine(ByVal RoutineName As String, ByVal RoutineText As String, ByVal connstr As String, ByVal connprv As String) As String
        Try
            Dim StoredProcName As String = String.Empty
            Dim ParamName(1) As String
            Dim ParamType(1) As String
            Dim ParamValue(1) As String
            Dim ret As String = String.Empty
            If RoutineName <> String.Empty Then
                ParamName(0) = "RoutineName"
                ParamType(0) = "String"
                ParamValue(0) = RoutineName
                ParamName(1) = "RoutineText"
                ParamType(1) = "String"
                ParamValue(1) = RoutineText
                StoredProcName = "OUR.BUILDROUTINE"
                'run storproc
                ret = RunSP(StoredProcName, 2, ParamName, ParamType, ParamValue, connstr, connprv)
            End If
            Return ret
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function CreateClass(ByVal ClassText As String, ByVal connstr As String, ByVal connprv As String) As String
        Dim ret As String = String.Empty
        Try
            Dim StoredProcName As String = String.Empty
            Dim ParamName(0) As String
            Dim ParamType(0) As String
            Dim ParamValue(0) As String
            If ClassText <> String.Empty Then
                'make params
                ReDim ParamName(0)
                ReDim ParamType(0)
                ReDim ParamValue(0)
                ParamName(0) = "ClassText"
                ParamName(0) = "String"
                ParamValue(0) = ClassText
                StoredProcName = "OUR.BUILDCLASSFROMSTRING"  'StorProc in OUR.INIT class 
                'run storproc
                ret = RunSP(StoredProcName, 1, ParamName, ParamType, ParamValue, connstr, connprv)
            End If
            Return ret
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function CorrectRecordOrder(ByVal dt As DataTable, ByVal orderfld As String) As DataTable
        Dim dtt As DataTable = dt
        Dim i As Integer
        'reorder
        For i = 0 To dtt.Rows.Count - 1
            dtt.Rows(i)(orderfld) = i + 1
        Next
        Return dtt
    End Function
    Public Function UpRowInDataTable(ByVal dt As DataTable, ByVal orderfld As String, ByVal uprowid As String) As DataTable
        Dim exm As String = String.Empty
        Dim ord1 As Integer = 0
        Dim ord2 As Integer = 0
        CorrectRecordOrder(dt, orderfld)
        Try
            ord1 = dt.Rows(uprowid - 1)(orderfld)
            ord2 = dt.Rows(uprowid)(orderfld)
            dt.Rows(uprowid - 1)(orderfld) = ord2
            dt.Rows(uprowid)(orderfld) = ord1
        Catch ex As Exception
            exm = ex.Message
        End Try
        Return dt
    End Function
    Public Function DownRowInDataTable(ByVal dt As DataTable, ByVal orderfld As String, ByVal uprowid As String) As DataTable
        Dim exm As String = String.Empty
        Dim ord1 As Integer = 0
        Dim ord2 As Integer = 0
        CorrectRecordOrder(dt, orderfld)
        Try
            ord1 = dt.Rows(uprowid + 1)(orderfld)
            ord2 = dt.Rows(uprowid)(orderfld)
            dt.Rows(uprowid + 1)(orderfld) = ord2
            dt.Rows(uprowid)(orderfld) = ord1
        Catch ex As Exception
            exm = ex.Message
        End Try
        Return dt
    End Function
    Public Function UpdateRecordOrderInDB(ByVal tbl As String, ByVal fldord As String, ByVal fldindx As String, ByVal fldindxtyp As String, ByVal dt As DataTable, Optional ByVal fld As String = "", Optional ByVal fldvalue As String = "", Optional ByVal fldtyp As String = "") As String
        Try
            Dim i As Integer
            Dim sqly As String = String.Empty
            Dim ret As String = String.Empty
            For i = 0 To dt.Rows.Count - 1
                sqly = "UPDATE " & tbl & " SET " & fldord & "=" & dt.Rows(i)(fldord).ToString & " WHERE " & fldindx & "="
                If fldindxtyp = "int" Then
                    sqly = sqly & dt.Rows(i)(fldindx).ToString
                Else
                    sqly = sqly & "'" & dt.Rows(i)(fldindx).ToString & "'"
                End If

                If fld <> "" AndAlso fldvalue <> "" AndAlso fldtyp = "int" Then
                    sqly = sqly & " AND " & fld & "=" & fldvalue
                ElseIf fld <> "" AndAlso fldvalue <> "" AndAlso fldtyp <> "int" Then
                    sqly = sqly & " AND " & fld & "='" & fldvalue & "'"
                End If
                ret = ExequteSQLquery(sqly)
                If ret <> "Query executed fine." Then
                    Return ret
                End If
            Next
        Catch ex As Exception
            Return ex.Message
        End Try
        Return ""
    End Function
    'Public Function GetTaskListSetting(ByVal unitname As String, ByVal logon As String, ByVal prop1 As String, Optional ByRef fldvalue As String = "") As String
    '    Dim ret As String = String.Empty
    '    Dim sqli As String = String.Empty
    '    If fldvalue.Trim <> "" Then
    '        sqli = "Select * FROM ourtasklistsetting WHERE UnitName=" & unitname.ToString & " AND [User]='" & logon & "' AND Prop1='" & prop1 & "' AND FldText='" & fldvalue & "'"
    '    Else
    '        sqli = "Select * FROM ourtasklistsetting WHERE UnitName=" & unitname.ToString & " AND [User]='" & logon & "' AND Prop1='" & prop1 & "'"
    '    End If
    '    Dim dvi As DataView = mRecords(sqli)
    '    If Not dvi Is Nothing AndAlso Not dvi.Table Is Nothing AndAlso dvi.Table.Rows.Count > 0 Then
    '        ret = dvi.Table.Rows(0)("FldColor")
    '        fldvalue = dvi.Table.Rows(0)("FldText")
    '    End If
    '    Return ret
    'End Function
    Public Function GetDefaultColors(ByVal unit As String, ByVal unitname As String, ByVal logon As String) As String
        Dim ret As String = String.Empty
        Dim er As String = String.Empty
        Dim sql As String = String.Empty
        Dim dv As DataView = Nothing
        Try
            If Not HasRecords("Select * FROM ourtasklistsetting WHERE UnitName='" & unitname & "' AND [User]='" & logon & "'") Then
                If Not HasRecords("Select * FROM ourtasklistsetting WHERE UnitName='" & unitname & "'") Then
                    'set default TaskList setting
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('header1','default','Task List',1,'#52573e','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('header2','default','Task',2,'#52573e','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('version','default','current',1,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('version','default','next',2,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('version','default','old',3,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('version','default','all',3,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('version','default','undefined',4,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)

                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','in progress',1,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','meeting',2,'#e9c46a','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','documentation',1,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','knowledge',2,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','how to',3,'#888677','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','problem',4,'#ff0000','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','done',4,'#6e6e6e','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)

                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','bug',1,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','eventually',2,'#e9c46a','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','planning',1,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','known bug',2,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','dismissed',3,'#888677','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','asap',4,'#ff0000','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','deleted',4,'#6e6e6e','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)

                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','redesign',1,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','testing',2,'#e9c46a','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','soon',1,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)
                    sql = "INSERT INTO ourtasklistsetting (Prop1, Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('status','default','other',2,'#ee8011','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                    er = ExequteSQLquery(sql)

                    ret = ret & "Task List default setting assigned."
                Else
                    'copy records
                    Dim dr As DataTable = mRecords("Select * FROM ourtasklistsetting WHERE UnitName='" & unitname & "' AND Prop3='default'").Table
                    For i = 0 To dr.Rows.Count - 1
                        sql = "INSERT INTO ourtasklistsetting (Prop1,Prop3,FldText,FldOrder,FldColor,Unit,UnitName,[User]) VALUES ('" & dr.Rows(i)("Prop1").ToString & "','','" & dr.Rows(i)("FldText").ToString & "'," & dr.Rows(i)("FldOrder").ToString & ",'" & dr.Rows(i)("FldColor").ToString & "','" & unit.ToString & "','" & unitname.ToString & "','" & logon & "')"
                        er = ExequteSQLquery(sql)
                    Next
                    ret = ret & "Task List default user setting assigned."
                End If
            End If
            dv = mRecords("Select * FROM ourtasklistsetting WHERE UnitName='" & unitname & "' AND [User]='" & logon & "'", er)
            If dv Is Nothing OrElse dv.Table Is Nothing Then
                ret = ret & "Task List setting failed."
            End If
        Catch ex As Exception
            ret = ex.Message
        End Try
        Return ret
    End Function
    Public Function ColorOfTask(ByVal taskid As Integer, ByVal logon As String, ByRef tooltp As String) As String
        Try
            Dim dv As DataView = mRecords("SELECT * FROM OURHelpDesk WHERE ID=" & taskid.ToString)
            If dv Is Nothing OrElse dv.Table Is Nothing OrElse dv.Table.Rows.Count = 0 Then
                Return ""
            End If
            Dim stas As String = dv.Table.Rows(0)("Status")
            tooltp = tooltp & ": " & dv.Table.Rows(0)("Ticket") & ", status = " & stas
            If stas.ToString.Trim = "" Then
                Return ""
            End If
            Dim unit = dv.Table.Rows(0)("Prop1")
            dv = mRecords("SELECT * FROM ourtasklistsetting WHERE Prop1='status' AND UnitName='" & unit & "' AND [User] ='" & logon & "' AND FldText ='" & stas & "'")
            If dv Is Nothing OrElse dv.Table Is Nothing OrElse dv.Table.Rows.Count = 0 Then
                Return ""
            End If
            Return dv.Table.Rows(0)("FldColor")
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function MakeDTColumnsNamesCLScompliant(ByVal dt As DataTable, Optional ByVal myconprv As String = "", Optional ByVal er As String = "") As DataTable
        Dim ret As String = String.Empty
        If myconprv = "" Then
            myconprv = System.Configuration.ConfigurationManager.ConnectionStrings.Item("mySQLconnection").ProviderName.ToString
        End If
        If dt Is Nothing Then
            ret = "ERROR!! No data"
            Return Nothing
        End If
        Try
            For i = 0 To dt.Columns.Count - 1
                dt.Columns(i).Caption = cleanText(FixReservedWords(dt.Columns(i).Caption, myconprv))
                dt.Columns(i).Caption = Regex.Replace(dt.Columns(i).Caption, "[^a-zA-Z0-9_]", "")  'CLS-compliant
                If myconprv.StartsWith("InterSystems.Data.") Then
                    dt.Columns(i).Caption = dt.Columns(i).Caption.Replace("_", "")
                End If
                dt.Columns(i).ColumnName = dt.Columns(i).Caption
            Next
            dt.AcceptChanges()
        Catch ex As Exception
            ret = "ERROR!! " & ex.Message
        End Try
        Return dt
    End Function
End Module
