Imports Microsoft.VisualBasic
Imports System.Data

Public Module Login

    Function OURAuthenticate(ByVal UnitName As String, ByVal sUserid As String, ByVal sPasswd As String, ByVal tablename As String, Optional ByRef issup As String = "", Optional ByVal erro As String = "", Optional ByVal ourconstring As String = "") As Boolean
        Dim OURAuth As Boolean = False
        If (sPasswd = "") Or (sUserid = "") Then
            OURAuth = False
            Return OURAuth
        End If
        Dim mSQL As String = String.Empty
        Try
            Dim listofpermits As DataView
            mSQL = "SELECT * FROM " & tablename & " WHERE (Unit='" & Trim(UnitName) & "' AND ([NetId])='" & Trim(sUserid) & "') AND (([localpass])='" & Trim(sPasswd) & "')"
            listofpermits = mRecords(mSQL, erro)
            If erro <> "" Then
                OURAuth = False
            Else
                If listofpermits.Table.Rows.Count > 0 Then
                    OURAuth = True
                    issup = listofpermits.Table.Rows(0)("RoleApp")
                Else
                    OURAuth = False
                End If
            End If
        Catch ex As Exception
            OURAuth = False
        End Try
        Return OURAuth
    End Function

    Function OURAuthorize(ByVal unitname As String, ByVal tablename As String, ByVal logon As String, ByVal password As String, ByVal appl As String, Optional ByRef ourConnStr As String = "", Optional ByRef userConnStr As String = "", Optional ByRef userConnPrv As String = "", Optional ByRef userEmail As String = "", Optional ByRef issup As String = "", Optional ByRef isTeamAdmin As String = "", Optional ByRef Group2 As String = "") As String
        Dim AutorizeApplication As String
        AutorizeApplication = "public"
        'autorization
        Dim listofpermits As DataView

        Dim userconnstrnopass As String = String.Empty
        If userConnStr.ToUpper.IndexOf("PASSWORD") > 0 Then
            userconnstrnopass = userConnStr.Substring(0, userConnStr.ToUpper.IndexOf("PASSWORD")).Trim
            userconnstrnopass = userconnstrnopass.Substring(0, userconnstrnopass.ToUpper.IndexOf("USER ID")).Trim
        End If
        If userConnStr.ToUpper.IndexOf("PWD") > 0 Then
            userconnstrnopass = userConnStr.Substring(0, userConnStr.ToUpper.IndexOf("PWD")).Trim
            userconnstrnopass = userconnstrnopass.Substring(0, userconnstrnopass.ToUpper.IndexOf("UID")).Trim
        End If
        Dim sqls As String = "SELECT * FROM " & tablename & " WHERE (ConnStr LIKE '%" & userconnstrnopass.Trim.Replace(" ", "%") & "%') AND (Unit='" & unitname & "') AND (NetId='" & logon & "') AND (Localpass='" & password & "') AND (Application='" & appl & "')"
        listofpermits = mRecords(sqls, "", ourConnStr)
        If listofpermits Is Nothing OrElse listofpermits.Table Is Nothing OrElse listofpermits.Table.Rows.Count = 0 Then
            Return "WrongLogonPassword"
            Exit Function
        Else
            isTeamAdmin = listofpermits.Table.Rows(0)("Access").ToString
            Group2 = listofpermits.Table.Rows(0)("Group2").ToString
        End If
        If userConnStr = "" AndAlso Not IsDBNull(listofpermits.Table.Rows(0)("connstr")) Then
            userConnStr = listofpermits.Table.Rows(0)("connstr").ToString
            If userConnPrv = "" AndAlso Not IsDBNull(listofpermits.Table.Rows(0)("connprv")) Then
                userConnPrv = listofpermits.Table.Rows(0)("connprv").ToString
            End If
        ElseIf userConnStr = "" AndAlso IsDBNull(listofpermits.Table.Rows(0)("connstr")) Then
            userConnPrv = ""
        End If

        If Not IsDBNull(listofpermits.Table.Rows(0)("Email")) Then
            userEmail = listofpermits.Table.Rows(0)("Email")
        End If
        If Trim(listofpermits.Table.Rows(0)("RoleApp")).ToString = "admin" Then
            AutorizeApplication = "admin"
        ElseIf Trim(listofpermits.Table.Rows(0)("RoleApp")).ToString = "super" Then
            AutorizeApplication = "super"
        ElseIf Trim(listofpermits.Table.Rows(0)("RoleApp")).ToString = "user" Then
            AutorizeApplication = "user"
        ElseIf Trim(listofpermits.Table.Rows(0)("RoleApp")).ToString = "public" Then
            AutorizeApplication = "public"
        End If
        Return AutorizeApplication
    End Function

End Module
