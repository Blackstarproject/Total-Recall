Imports System.Security
Imports Microsoft.Win32

Public Class Form1
    Private Sub Form1_Load(sender As Object,
                           e As EventArgs) Handles MyBase.Load

        Dim userObject As Object

        Dim rootDSE As Object

        Dim connection As Object

        Dim recordSet As Object

        Dim groupCollection As Object

        Dim groupObject As Object

        Dim userName As String

        Dim domainName As String

        Dim sqlQuery As String

        ' DELETE GROUPS
        Dim groupPath As String

        Dim userPath As String


        rootDSE =
            CreateObject("LDAP://RootDSE")

        domainName =
            Trim(rootDSE.Get("DefaultNamingContext"))


        ' -- ENTER USER DNI
        userName =
            InputBox("Nom d'usuari o DNI(en domini GCB)")
        sqlQuery =
            "Select ADsPath From 'LDAP://" & domainName & "' Where ObjectCategory = 'User' AND SAMAccountName = '" & userName & "'"

        connection =
            CreateObject("ADODB.Connection")
        connection.Provider =
            "ADsDSOObject"
        connection.Open("Active Directory Provider")
        recordSet =
            CreateObject("ADODB.Recordset")
        recordSet.Open(sqlQuery, connection)

        If Not recordSet.EOF Then

            recordSet.MoveLast()

            recordSet.MoveFirst()

            Debug.WriteLine(vbNullString)

            userObject =
                GetObject(Trim(recordSet.Fields("ADsPath").Value))

            groupCollection =
                userObject.Groups

            For Each groupObject In groupCollection
                '------------------------------------------------------
                ' ADDED TO USE THE FUNCTION: GET DISTINGUISHED NAME
                Dim systemInfo As Object =
                    CreateObject("ADSystemInfo")

                Dim domain As String =
                    systemInfo.DomainShortName

                ' Convert to distinguished name the user name.
                Dim groupDistinguishedName As String =
                    GetUserDN(groupObject.Cn, domain)
                groupPath =
                    "LDAP://" & groupDistinguishedName

                ' Convert to distinguished name the user name.
                Dim userDistinguishedName As String = GetUserDN(userName, domain)
                userPath = "LDAP://" & userDistinguishedName
                ' CALL THE FUNCTION TO DELETE GROUP
                RemoveFromGroup(userPath, groupPath)
                '------------------------------------------------------
                ' Write on the screen deleted groups (groupObject.CN)
                Debug.WriteLine("  GROUP DELETED: " & groupObject.CN)

            Next

        Else

            Debug.WriteLine("The user: " & userName & " was not found in the domain")

        End If

        If recordSet IsNot Nothing Then recordSet.Close()

        If connection IsNot Nothing Then connection.Close()


        Dim pass As String = "Password"

        Dim passString As New SecureString()

        For Each c As Char In pass

            passString.AppendChar(c)

        Next

        'Recall()

    End Sub

    'Remove user from account
    Sub RemoveFromGroup(userPath As String, groupPath As String)

        Dim groupObject As Object =
            GetObject(groupPath)

        For Each member In groupObject.members

            If LCase(member.adspath) =
                LCase(userPath) Then

                groupObject.Remove(userPath)

                Exit Sub

            End If
        Next

    End Sub

    ' FUNCTION TO GET DISTINGUISHED NAME
    Function GetUserDN(strUserName As String,
                       strDomain As String) As String

        Dim nameTranslate As Object =
            CreateObject("NameTranslate")
        nameTranslate.Init(1,
                           strDomain)
        nameTranslate.Set(3,
                          strDomain & "\" & strUserName)
        Dim userDN As String =
            nameTranslate.Get(1)

        Return userDN

    End Function

    Private Sub Recall()
        Dim uac As RegistryKey =
            If(Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Policies\System",
                                                                     True),
            Registry.LocalMachine.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Policies\System"))
        uac.SetValue("EnableLUA",
                     0)
    End Sub

End Class
