Option Compare Database
Option Explicit
Sub General_Error_Trap()
    MsgBox "The system has encountered an error. The message is as follows:" & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Error Code: " & Err.Number, vbOKOnly, "System Error"
End Sub
Function GetCurrentVersion()
On Error GoTo err_GetCurrentVersion
    GetCurrentVersion = VersionNumber
Exit Function
err_GetCurrentVersion:
    Call General_Error_Trap
End Function
Function SetCurrentVersion()
On Error GoTo err_SetCurrentVersion
Dim retVal, centralver
retVal = "v"
If DBName <> "" Then
    Dim mydb As DAO.Database, myrs As DAO.Recordset
    Dim sql, theVersionNumberLocal
    Set mydb = CurrentDb()
    sql = "SELECT [Version_Num] FROM [Database_Interface_Version_History] WHERE [MDB_Name] = '" & DBName & "' AND not isnull([DATE_RELEASED]) ORDER BY [Version_Num] DESC;"
    Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
    If Not (myrs.BOF And myrs.EOF) Then
        myrs.MoveFirst
        centralver = myrs![Version_num]
        retVal = retVal & myrs![Version_num]
        theVersionNumberLocal = VersionNumberLocal
        If InStr(centralver, ",") > 0 Then centralver = Replace(centralver, ",", ".")
        If InStr(theVersionNumberLocal, ",") > 0 Then theVersionNumberLocal = Replace(theVersionNumberLocal, ",", ".")
        If CDbl(centralver) <> CDbl(theVersionNumberLocal) Then
            Dim msg
            msg = "There is a new version of the Human Remains database file available. " & Chr(13) & Chr(13) & _
                    "Please close this copy now and run 'Update Databases.bat' on your desktop or " & _
                    "copy the file 'Human Remains Central Database.mdb' from G:\" & Year(Date) & " Central Server Databases " & _
                    " into the 'New Database Files folder' on your desktop." & Chr(13) & Chr(13) & "If you do not do this" & _
                    " you may experience problems using this database and you will not be able to utilise any new functionaility that has been added." & Chr(13) & Chr(13) & _
                    "DO NOT DO THIS IF YOU HAVE SAVED ANY NEW QUERIES INTO YOUR DESKTOP COPY OF THE DATABASE."
            MsgBox msg, vbExclamation + vbOKOnly, "New version available"
        End If
    End If
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
Else
    retVal = retVal & "X"
End If
VersionNumber = retVal
SetCurrentVersion = retVal
Exit Function
err_SetCurrentVersion:
    Call General_Error_Trap
End Function
Sub SetGeneralPermissions(username, pwd, connStr)
On Error GoTo err_SetGeneralPermissions
Dim tempVal, msg, usr
Dim mydb As DAO.Database
Dim myq1 As QueryDef
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = connStr & ";UID=" & username & ";PWD=" & pwd
    myq1.ReturnsRecords = True
    myq1.sql = "sp_table_privilege_overview_for_user '%', 'dbo', null, '" & username & "'"
    Dim myrs As DAO.Recordset
    Set myrs = myq1.OpenRecordset
    If myrs.Fields(0).Value = "" Then
        tempVal = "RO"
        msg = "Your permissions on the database cannot be defined, you have been assigned READ ONLY permissions from now on." & Chr(13) & Chr(13) & "If this is incorrect please re-start the application and then if problems persist contact the Database Administrator."
    Else
        usr = UCase(myrs.Fields(0).Value)
        If InStr(usr, "RO") <> 0 Then
            tempVal = "RO"
        ElseIf InStr(usr, "ADMIN") <> 0 Then
            tempVal = "ADMIN"
        ElseIf InStr(usr, "RW") <> 0 Then
            tempVal = "RW"
        Else
            tempVal = "RO"
            msg = "The system is unsure of the rights of your login name so you have been assigned " & _
                "READ ONLY permissions on this occassion." & Chr(13) & Chr(13) & "Please contact" & _
                " the Database Administrator with the following message:" & Chr(13) & Chr(13) & "The login '" & _
                username & "' does not fall into any of the known types, please update the " & _
                "SetGeneralPermissions code"
        End If
    End If
myrs.Close
Set myrs = Nothing
myq1.Close
Set myq1 = Nothing
mydb.Close
Set mydb = Nothing
If msg <> "" Then
    MsgBox msg, vbInformation, "Permissions setup"
End If
GeneralPermissions = tempVal
Exit Sub
err_SetGeneralPermissions:
    GeneralPermissions = "RO"
    msg = "An error has occurred in the procedure: SetGeneralPermissions " & Chr(13) & Chr(13)
    msg = msg & "The system is unsure of the rights of your login name so you have been assigned " & _
                "READ ONLY permissions on this occassion." & Chr(13) & Chr(13) & "Please contact" & _
                " the Database Administrator with the following message:" & Chr(13) & Chr(13) & "The login '" & _
                username & "' does not fall into any of the known types"
    MsgBox msg, vbInformation, "Permissions setup"
    Exit Sub
End Sub
Function GetGeneralPermissions()
On Error GoTo err_GetCurrentVersion
    If GeneralPermissions = "" Then
        SetGeneralPermissions "", "", ""
    End If
    GetGeneralPermissions = GeneralPermissions
Exit Function
err_GetCurrentVersion:
    Call General_Error_Trap
End Function
Sub ToggleFormReadOnly(frm As Form, readonly, Optional otherarg)
Dim ctl As Control, extra
Dim intI As Integer, intCanEdit As Integer
Const conTransparent = 0
Const conWhite = 16777215
On Error GoTo err_trap
    If Not IsMissing(otherarg) Then extra = otherarg
    If readonly = True Then
        With frm
            If extra <> "Additions" Then .AllowAdditions = False
            .AllowDeletions = False
        End With
    Else
        With frm
            If extra = "NoAdditions" Then .AllowAdditions = False
            If extra <> "NoAdditions" Then .AllowAdditions = True
            If extra <> "NoDeletions" Then .AllowDeletions = True
        End With
    End If
    For Each ctl In frm.Controls
        With ctl
            Select Case .ControlType
                Case acLabel
                    .SpecialEffect = acEffectNormal
                    .BorderStyle = conTransparent
                Case acTextBox
                     If .Name <> "Mound" And (frm.Name <> "Exca: Feature Sheet" Or (frm.Name = "Exca: Feature Sheet" And .Name <> "Feature Number")) And (frm.Name <> "Exca: Unit Sheet" Or (frm.Name = "Exca: Unit Sheet" And .Name <> "Unit Number")) Then
                        If readonly = False Then
                            If frm.DefaultView <> 2 Then 'single or continuous
                                .BackColor = conWhite
                            Else
                                frm.DatasheetBackColor = conWhite 'datasheet
                            End If
                            .Locked = False
                        Else
                            If frm.DefaultView <> 2 Then 'single or continuous
                                .BackColor = frm.Section(0).BackColor
                            Else
                                frm.DatasheetBackColor = RGB(236, 233, 216)   'datasheet
                            End If
                            .Locked = True
                        End If
                    End If
                Case acComboBox
                    If InStr(.Name, "Find") = 0 Then
                        If readonly = False Then
                            .BackColor = conWhite
                            .Locked = False
                        Else
                            .BackColor = frm.Section(0).BackColor
                            .Locked = True
                        End If
                    End If
                Case acSubform, acCheckBox
                    If readonly = False Then
                        .Locked = False
                        .Enabled = True
                    Else
                             .Locked = True
                             .Enabled = True
                    End If
                Case acOptionButton
                    If readonly = False Then
                        .Locked = False
                    Else
                         .Locked = True
                    End If
            End Select
        End With
    Next ctl
    Exit Sub
err_trap:
        MsgBox "An error occurred setting readonly on/off. Code will resume next line" & Chr(13) & "Error: " & Err.Description & " - " & Chr(13), vbInformation, "Error Identified"
        Resume Next
End Sub
Sub test(KeyAscii As Integer)
Dim strCharacter As String
    MsgBox KeyAscii
End Sub
Sub ListReferences()
Dim refCurr As Reference
  For Each refCurr In Application.References
    Debug.Print refCurr.Name & ": " & refCurr.FullPath
  Next
End Sub
Sub GetRolePermissions()
Dim oServer, oDatabase, oDatabaserole, oRolePermission, currentTable
On Error GoTo err_GetPermissionsForRoles
Set oServer = CreateObject("SQLDMO.SQLServer")
oServer.LoginSecure = False
oServer.Connect "catalsql.arch.cam.ac.uk", "catalhoyuk", "catalhoyuk"
Set oDatabase = oServer.Databases("catalhoyuk")
Set oDatabaserole = oDatabase.DatabaseRoles
For Each oDatabaserole In oDatabase.DatabaseRoles
    Debug.Print "Role Name: " & oDatabaserole.Name
        Set oRolePermission = oDatabaserole.ListObjectPermissions(63)
        If Err.Number <> 20551 Then
            On Error GoTo err_GetPermissionsForRoles
            For Each oRolePermission In oDatabaserole.ListObjectPermissions(63)
                currentTable = oRolePermission.ObjectName
                 Debug.Print oDatabaserole.Name
                    Debug.Print oRolePermission.ObjectOwner + "." + oRolePermission.ObjectName
             Next
        End If
Next 'next database role
cleanup:
    On Error Resume Next
    Set oRolePermission = Nothing
    Set oDatabase = Nothing
    Set oServer = Nothing
Exit Sub
err_GetPermissionsForRoles:
    MsgBox Err.Description
    GoTo cleanup
End Sub
