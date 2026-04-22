Imports Microsoft.VisualBasic
Imports System.Data

Public Module UnitLogin
    Function UnitAuthenticate(ByVal UnitName As String, ByVal sUserid As String, ByVal sPasswd As String, Optional ByVal erro As String = "") As Boolean
        'Unit authentication
        Dim UnitAuth As Boolean = False
        'example how make authentication with oracle db using LDAP
        If (sPasswd = "") Or (sUserid = "") Then
            Return False
            Exit Function
        End If
        Err.Number = 0
        Dim oProvider = GetObject("LDAP:")
        If Err.Number <> 0 Then
            MsgBox("GetObject = " & Err.Number)
            UnitAuth = False
        Else
            'or add other method of Unit login authentication here

            If Err.Number = 0 AndAlso erro = "" Then
                UnitAuth = True
            Else
                UnitAuth = False
            End If
        End If
        Return UnitAuth
    End Function

    Function UnitAuthorize(ByVal ps As Boolean, ByVal unit As String, ByVal tablename As String, ByVal rolefieldname As String, ByVal emailfieldname As String, ByVal logonfieldname As String, ByVal passwordfieldname As String, ByVal logon As String, Optional ByRef password As String = "", Optional ByRef userEmail As String = "") As String
        'Unit autorization
        'ps=True if check password
        'EXAMPLE how make custom autorization using unit autorization table - tablename
        Dim Autorize As String = ""
        'autorization
        Dim listofpermits As DataView
        'user autorization
        Dim sqlst As String = String.Empty
        If ps = True Then
            sqlst = "SELECT * FROM " & tablename & " WHERE (" & logonfieldname & "='" & logon & "') And (" & passwordfieldname & " ='" & password & "')"
        Else
            sqlst = "SELECT * FROM " & tablename & " WHERE (" & logonfieldname & "='" & logon & "')"
        End If
        listofpermits = mRecords(sqlst)
        If listofpermits Is Nothing OrElse listofpermits.Table Is Nothing OrElse listofpermits.Table.Rows.Count = 0 Then
            Return "WrongLogonPassword"
            Exit Function
        End If
        If listofpermits.Table.Rows.Count = 1 Then
            If Not IsDBNull(listofpermits.Table.Rows(0)("emailfieldname")) Then
                userEmail = listofpermits.Table.Rows(0)("emailfieldname")
            End If
            If Trim(listofpermits.Table.Rows(0)(rolefieldname)) = "admin" Then
                Autorize = "admin"
            ElseIf Trim(listofpermits.Table.Rows(0)(rolefieldname)) = "super" Then
                Autorize = "super"
            ElseIf Trim(listofpermits.Table.Rows(0)(rolefieldname)) = "user" Then
                Autorize = "user"
            Else
                Autorize = "public"
            End If
        Else
            Autorize = ""
        End If
        Return Autorize
    End Function
End Module

