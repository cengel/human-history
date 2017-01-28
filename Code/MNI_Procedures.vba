Option Compare Database
Option Explicit
Sub CalcMNI()
Dim mydb As DAO.Database, burials As DAO.Recordset, sql
Dim CurrentFeatureNo, CurrentAdultCount, CurrentJuvenileCount, CurrentNeonateCount, CurrentMNI
Dim MNIQueryBuilder As DAO.Recordset, TableName, FieldName, Criteria, WhereClause, AgeCategory
Dim CalcIndividuals As DAO.Recordset, sqlInsert, sqlDelete, AbleToCalc, AdditionalWhere
Set mydb = CurrentDb
Set burials = mydb.OpenRecordset("Select [Feature Number] from [Exca: Features] WHERE [Feature Type] = 'burial' ORDER By [Feature Number];", dbOpenSnapshot)
If Not burials.EOF And Not burials.BOF Then
    sqlDelete = "DELETE FROM [HR_Feature_MNI];"
    DoCmd.RunSQL sqlDelete
    burials.MoveFirst
    Do Until burials.EOF
        AbleToCalc = False
        CurrentFeatureNo = burials![Feature Number]
        Forms![FRM_MNI]![txtFeature] = "Calculating for Feature " & CurrentFeatureNo
        DoCmd.RepaintObject acForm, "FRM_MNI"
        Set MNIQueryBuilder = mydb.OpenRecordset("HR_Feature_MNI_QueryBuilder", dbOpenDynaset)
        If Not MNIQueryBuilder.EOF And Not MNIQueryBuilder.BOF Then
            MNIQueryBuilder.MoveFirst
            CurrentAdultCount = 0
            CurrentJuvenileCount = 0
            CurrentNeonateCount = 0
            Dim AdultfldWithMost, JuvfldWithMost, NeofldWithMost
            AdultfldWithMost = ""
            JuvfldWithMost = ""
            NeofldWithMost = ""
            Do Until MNIQueryBuilder.EOF
                TableName = "[" & MNIQueryBuilder![TableName] & "]"
                FieldName = "[" & MNIQueryBuilder![FieldName] & "]"
                Criteria = MNIQueryBuilder![Criteria]
                WhereClause = FieldName & " " & Criteria
                AgeCategory = MNIQueryBuilder![AgeCategory]
                Forms![FRM_MNI]![txtMsg] = "Checking " & TableName & " - " & FieldName & " for feature number " & CurrentFeatureNo
                DoCmd.RepaintObject acForm, "FRM_MNI"
                If TableName = "[HR_ageing and sexing]" Then
                    sql = "SELECT Count([dbo_Exca: Units in Features].In_feature) AS CountOfIn_feature " & _
                        "FROM ([Exca: Unit Sheet with Relationships] INNER JOIN [dbo_Exca: Units in Features] ON " & _
                        "[Exca: Unit Sheet with Relationships].[Unit Number] = [dbo_Exca: Units in Features].Unit) " & _
                        "INNER JOIN " & TableName & " ON [Exca: Unit Sheet with Relationships].[Unit Number] = " & TableName & ".[Unit Number] " & _
                        "GROUP BY [dbo_Exca: Units in Features].In_feature, " & TableName & "." & FieldName & _
                        " HAVING ((([dbo_Exca: Units in Features].In_feature)=" & CurrentFeatureNo & ") AND ((" & TableName & "." & WhereClause & ")));"
                Else
                    If InStr(TableName, "Measure") > 0 And AgeCategory <> "adult" Then
                        If AgeCategory = "neonate" Then
                            AdditionalWhere = "[HR_ageing and sexing].[age category]=0 OR [HR_ageing and sexing].[age category]=9"
                        ElseIf AgeCategory = "juvenile" Then
                            AdditionalWhere = "[HR_ageing and sexing].[age category]=1 OR [HR_ageing and sexing].[age category]=2"
                        End If
                        sql = "SELECT Count([dbo_Exca: Units in Features].In_feature) AS CountOfIn_feature  " & _
                        "FROM (([Exca: Unit Sheet with Relationships] INNER JOIN [dbo_Exca: Units in Features] ON " & _
                        "[Exca: Unit Sheet with Relationships].[Unit Number] = [dbo_Exca: Units in Features].Unit) " & _
                        "INNER JOIN " & TableName & "  ON [Exca: Unit Sheet with Relationships].[Unit Number] = " & TableName & ".UnitNumber) " & _
                        "INNER JOIN [HR_ageing and sexing] ON (" & TableName & ".[Individual number] = [HR_ageing and sexing].[Individual number]) " & _
                        " AND (" & TableName & ".UnitNumber = [HR_ageing and sexing].[unit number]) " & _
                        "GROUP BY [dbo_Exca: Units in Features].In_feature, " & TableName & "." & FieldName & ", [HR_ageing and sexing].[age category] " & _
                        "HAVING ((([dbo_Exca: Units in Features].In_feature)=" & CurrentFeatureNo & ") AND (" & AdditionalWhere & ") AND ((" & TableName & "." & WhereClause & ")));"
                    Else
                        sql = "SELECT Count([dbo_Exca: Units in Features].In_feature) AS CountOfIn_feature " & _
                        "FROM ([Exca: Unit Sheet with Relationships] INNER JOIN [dbo_Exca: Units in Features] ON " & _
                        "[Exca: Unit Sheet with Relationships].[Unit Number] = [dbo_Exca: Units in Features].Unit) " & _
                        "INNER JOIN " & TableName & " ON [Exca: Unit Sheet with Relationships].[Unit Number] = " & TableName & ".UnitNumber " & _
                        "GROUP BY [dbo_Exca: Units in Features].In_feature, " & TableName & "." & FieldName & _
                        " HAVING ((([dbo_Exca: Units in Features].In_feature)=" & CurrentFeatureNo & ") AND ((" & TableName & "." & WhereClause & ")));"
                    End If
                End If
                Set CalcIndividuals = mydb.OpenRecordset(sql, dbOpenSnapshot)
                If Not CalcIndividuals.EOF And Not CalcIndividuals.BOF Then
                    AbleToCalc = True
                    CalcIndividuals.MoveFirst
                    Do Until CalcIndividuals.EOF
                        If AgeCategory = "juvenile" Then
                            If CurrentJuvenileCount < CalcIndividuals![CountOfIn_Feature] Then
                                CurrentJuvenileCount = CalcIndividuals![CountOfIn_Feature]
                                JuvfldWithMost = FieldName
                            End If
                        ElseIf AgeCategory = "neonate" Then
                            If CurrentNeonateCount < CalcIndividuals![CountOfIn_Feature] Then
                                CurrentNeonateCount = CalcIndividuals![CountOfIn_Feature]
                                NeofldWithMost = FieldName
                            End If
                        Else 'adult
                            If CurrentAdultCount < CalcIndividuals![CountOfIn_Feature] Then
                                CurrentAdultCount = CalcIndividuals![CountOfIn_Feature]
                                AdultfldWithMost = FieldName
                            End If
                        End If
                    CalcIndividuals.MoveNext
                    Loop
                Else
                End If
                CalcIndividuals.Close
                Set CalcIndividuals = Nothing
            MNIQueryBuilder.MoveNext
            Loop
            If AbleToCalc = True Then
                sqlInsert = "INSERT INTO [HR_Feature_MNI] ([FeatureNumber], [MNI], [Adults], [Juveniles], [Neonates], [Notes], [LastGenerated]) VALUES (" & CurrentFeatureNo & ", " & CurrentAdultCount + CurrentJuvenileCount + CurrentNeonateCount & "," & CurrentAdultCount & "," & CurrentJuvenileCount & ", " & CurrentNeonateCount & ", 'Adult highest count from: " & AdultfldWithMost & ", Juv highest count from: " & JuvfldWithMost & ", Neo highest count from: " & NeofldWithMost & "','" & Now() & "');"
            Else
                sqlInsert = "INSERT INTO [HR_Feature_MNI] ([FeatureNumber], [Notes], [LastGenerated]) VALUES (" & CurrentFeatureNo & ", 'Not enough data available yet to undertake calculation', '" & Now() & "');"
            End If
            DoCmd.RunSQL sqlInsert
            MNIQueryBuilder.Close
            Set MNIQueryBuilder = Nothing
        Else
            MsgBox "The MNI Query Builder table which is used to calculate the MNI is empty. The MNI calculation cannot be done without this information. Contact the DBA", vbExclamation, "Cannot proceed"
            MNIQueryBuilder.Close
            Set MNIQueryBuilder = Nothing
            burials.Close
            Set burials = Nothing
            Exit Do
        End If
        burials.MoveNext
    Loop
Else
    MsgBox "No burial features have been found in the excavation database. No MNI can be calculated.", vbInformation, "No Burials found"
End If
burials.Close
Set burials = Nothing
End Sub
Sub CalcSpaceMNI()
Dim mydb As DAO.Database, spaces As DAO.Recordset, sql
Dim CurrentSpaceNo, CurrentAdultCount, CurrentJuvenileCount, CurrentNeonateCount, CurrentMNI
Dim MNIQueryBuilder As DAO.Recordset, TableName, FieldName, Criteria, WhereClause, AgeCategory
Dim CalcIndividuals As DAO.Recordset, sqlInsert, sqlDelete, AbleToCalc, AdditionalWhere
Set mydb = CurrentDb
Set spaces = mydb.OpenRecordset("SELECT DISTINCT [dbo_Exca: Space Sheet].[Space number] FROM [Exca: Unit Sheet] INNER JOIN ([dbo_Exca: Units in Spaces] INNER JOIN [dbo_Exca: Space Sheet] ON [dbo_Exca: Units in Spaces].In_space = [dbo_Exca: Space Sheet].[Space number]) ON [Exca: Unit Sheet].[Unit Number] = [dbo_Exca: Units in Spaces].Unit WHERE ((([Exca: Unit Sheet].Category)='skeleton'));")
If Not spaces.EOF And Not spaces.BOF Then
    sqlDelete = "DELETE FROM [HR_Space_MNI];"
    DoCmd.RunSQL sqlDelete
    spaces.MoveFirst
    Do Until spaces.EOF
        AbleToCalc = False
        CurrentSpaceNo = spaces![Space Number]
        Forms![FRM_MNI]![txtFeature] = "Calculating for Space " & CurrentSpaceNo
        DoCmd.RepaintObject acForm, "FRM_MNI"
        Set MNIQueryBuilder = mydb.OpenRecordset("HR_Feature_MNI_QueryBuilder", dbOpenDynaset)
        If Not MNIQueryBuilder.EOF And Not MNIQueryBuilder.BOF Then
            MNIQueryBuilder.MoveFirst
            CurrentAdultCount = 0
            CurrentJuvenileCount = 0
            CurrentNeonateCount = 0
            Dim AdultfldWithMost, JuvfldWithMost, NeofldWithMost
            AdultfldWithMost = ""
            JuvfldWithMost = ""
            NeofldWithMost = ""
            Do Until MNIQueryBuilder.EOF
                TableName = "[" & MNIQueryBuilder![TableName] & "]"
                FieldName = "[" & MNIQueryBuilder![FieldName] & "]"
                Criteria = MNIQueryBuilder![Criteria]
                WhereClause = FieldName & " " & Criteria
                AgeCategory = MNIQueryBuilder![AgeCategory]
                Forms![FRM_MNI]![txtMsg] = "Checking " & TableName & " - " & FieldName & " for space number " & CurrentSpaceNo
                DoCmd.RepaintObject acForm, "FRM_MNI"
                If TableName = "[HR_ageing and sexing]" Then
                    sql = "SELECT Count([dbo_Exca: Units in Spaces].In_space) AS CountOfIn_space " & _
                        "FROM ([Exca: Unit Sheet with Relationships] INNER JOIN [dbo_Exca: Units in Spaces] ON " & _
                        "[Exca: Unit Sheet with Relationships].[Unit Number] = [dbo_Exca: Units in Spaces].Unit) " & _
                        "INNER JOIN " & TableName & " ON [Exca: Unit Sheet with Relationships].[Unit Number] = " & TableName & ".[Unit Number] " & _
                        "GROUP BY [dbo_Exca: Units in Spaces].In_space, " & TableName & "." & FieldName & _
                        " HAVING ((([dbo_Exca: Units in Spaces].In_space)=" & CurrentSpaceNo & ") AND ((" & TableName & "." & WhereClause & ")));"
                Else
                    If InStr(TableName, "Measure") > 0 And AgeCategory <> "adult" Then
                        If AgeCategory = "neonate" Then
                            AdditionalWhere = "[HR_ageing and sexing].[age category]=0 OR [HR_ageing and sexing].[age category]=9"
                        ElseIf AgeCategory = "juvenile" Then
                            AdditionalWhere = "[HR_ageing and sexing].[age category]=1 OR [HR_ageing and sexing].[age category]=2"
                        End If
                        sql = "SELECT Count([dbo_Exca: Units in Spaces].In_Space) AS CountOfIn_space  " & _
                        "FROM (([Exca: Unit Sheet with Relationships] INNER JOIN [dbo_Exca: Units in Spaces] ON " & _
                        "[Exca: Unit Sheet with Relationships].[Unit Number] = [dbo_Exca: Units in Spaces].Unit) " & _
                        "INNER JOIN " & TableName & "  ON [Exca: Unit Sheet with Relationships].[Unit Number] = " & TableName & ".UnitNumber) " & _
                        "INNER JOIN [HR_ageing and sexing] ON (" & TableName & ".[Individual number] = [HR_ageing and sexing].[Individual number]) " & _
                        " AND (" & TableName & ".UnitNumber = [HR_ageing and sexing].[unit number]) " & _
                        "GROUP BY [dbo_Exca: Units in Spaces].In_space, " & TableName & "." & FieldName & ", [HR_ageing and sexing].[age category] " & _
                        "HAVING ((([dbo_Exca: Units in Spaces].In_space)=" & CurrentSpaceNo & ") AND (" & AdditionalWhere & ") AND ((" & TableName & "." & WhereClause & ")));"
                    Else
                        sql = "SELECT Count([dbo_Exca: Units in Spaces].In_space) AS CountOfIn_space " & _
                        "FROM ([Exca: Unit Sheet with Relationships] INNER JOIN [dbo_Exca: Units in Spaces] ON " & _
                        "[Exca: Unit Sheet with Relationships].[Unit Number] = [dbo_Exca: Units in Spaces].Unit) " & _
                        "INNER JOIN " & TableName & " ON [Exca: Unit Sheet with Relationships].[Unit Number] = " & TableName & ".UnitNumber " & _
                        "GROUP BY [dbo_Exca: Units in Spaces].In_space, " & TableName & "." & FieldName & _
                        " HAVING ((([dbo_Exca: Units in Spaces].In_space)=" & CurrentSpaceNo & ") AND ((" & TableName & "." & WhereClause & ")));"
                    End If
                End If
                Set CalcIndividuals = mydb.OpenRecordset(sql, dbOpenSnapshot)
                If Not CalcIndividuals.EOF And Not CalcIndividuals.BOF Then
                    AbleToCalc = True
                    CalcIndividuals.MoveFirst
                    Do Until CalcIndividuals.EOF
                        If AgeCategory = "juvenile" Then
                            If CurrentJuvenileCount < CalcIndividuals![CountOfIn_Space] Then
                                CurrentJuvenileCount = CalcIndividuals![CountOfIn_Space]
                                JuvfldWithMost = FieldName
                            End If
                        ElseIf AgeCategory = "neonate" Then
                            If CurrentNeonateCount < CalcIndividuals![CountOfIn_Space] Then
                                CurrentNeonateCount = CalcIndividuals![CountOfIn_Space]
                                NeofldWithMost = FieldName
                            End If
                        Else 'adult
                            If CurrentAdultCount < CalcIndividuals![CountOfIn_Space] Then
                                CurrentAdultCount = CalcIndividuals![CountOfIn_Space]
                                AdultfldWithMost = FieldName
                            End If
                        End If
                    CalcIndividuals.MoveNext
                    Loop
                Else
                End If
                CalcIndividuals.Close
                Set CalcIndividuals = Nothing
            MNIQueryBuilder.MoveNext
            Loop
            If AbleToCalc = True Then
                sqlInsert = "INSERT INTO [HR_Space_MNI] ([SpaceNumber], [MNI], [Adults], [Juveniles], [Neonates], [Notes], [LastGenerated]) VALUES (" & CurrentSpaceNo & ", " & CurrentAdultCount + CurrentJuvenileCount + CurrentNeonateCount & "," & CurrentAdultCount & "," & CurrentJuvenileCount & ", " & CurrentNeonateCount & ",'Adult highest count from: " & AdultfldWithMost & ", Juv highest count from: " & JuvfldWithMost & ", Neo highest count from: " & NeofldWithMost & "', '" & Now() & "');"
            Else
                sqlInsert = "INSERT INTO [HR_Space_MNI] ([SpaceNumber], [Notes], [LastGenerated]) VALUES (" & CurrentSpaceNo & ", 'Not enough data available yet to undertake calculation', '" & Now() & "');"
            End If
            DoCmd.RunSQL sqlInsert
            MNIQueryBuilder.Close
            Set MNIQueryBuilder = Nothing
        Else
            MsgBox "The MNI Query Builder table which is used to calculate the MNI is empty. The MNI calculation cannot be done without this information. Contact the DBA", vbExclamation, "Cannot proceed"
            MNIQueryBuilder.Close
            Set MNIQueryBuilder = Nothing
            spaces.Close
            Set spaces = Nothing
            Exit Do
        End If
        spaces.MoveNext
    Loop
Else
    MsgBox "No spaces containing skeleton units have been found in the excavation database. No MNI can be calculated.", vbInformation, "No spaces found"
End If
spaces.Close
Set spaces = Nothing
End Sub
Sub CalcBuildingMNI()
Dim mydb As DAO.Database, buildings As DAO.Recordset, sql
Dim CurrentBuildingNo, CurrentAdultCount, CurrentJuvenileCount, CurrentNeonateCount, CurrentMNI
Dim MNIQueryBuilder As DAO.Recordset, TableName, FieldName, Criteria, WhereClause, AgeCategory
Dim CalcIndividuals As DAO.Recordset, sqlInsert, sqlDelete, AbleToCalc, AdditionalWhere
Set mydb = CurrentDb
Set buildings = mydb.OpenRecordset("SELECT DISTINCT [dbo_Exca: Units in Buildings].In_building As [Building number] FROM [Exca: Unit Sheet] INNER JOIN [dbo_Exca: Units in Buildings] ON [Exca: Unit Sheet].[Unit Number] = [dbo_Exca: Units in Buildings].Unit WHERE ((([Exca: Unit Sheet].Category)='skeleton'));")
If Not buildings.EOF And Not buildings.BOF Then
    sqlDelete = "DELETE FROM [HR_Building_MNI];"
    DoCmd.RunSQL sqlDelete
    buildings.MoveFirst
    Do Until buildings.EOF
        AbleToCalc = False
        CurrentBuildingNo = buildings![Building Number]
        Forms![FRM_MNI]![txtFeature] = "Calculating for Building " & CurrentBuildingNo
        DoCmd.RepaintObject acForm, "FRM_MNI"
        Set MNIQueryBuilder = mydb.OpenRecordset("HR_Feature_MNI_QueryBuilder", dbOpenDynaset)
        If Not MNIQueryBuilder.EOF And Not MNIQueryBuilder.BOF Then
            MNIQueryBuilder.MoveFirst
            CurrentAdultCount = 0
            CurrentJuvenileCount = 0
            CurrentNeonateCount = 0
            Dim AdultfldWithMost, JuvfldWithMost, NeofldWithMost
            AdultfldWithMost = ""
            JuvfldWithMost = ""
            NeofldWithMost = ""
            Do Until MNIQueryBuilder.EOF
                TableName = "[" & MNIQueryBuilder![TableName] & "]"
                FieldName = "[" & MNIQueryBuilder![FieldName] & "]"
                Criteria = MNIQueryBuilder![Criteria]
                WhereClause = FieldName & " " & Criteria
                AgeCategory = MNIQueryBuilder![AgeCategory]
                Forms![FRM_MNI]![txtMsg] = "Checking " & TableName & " - " & FieldName & " for Building number " & CurrentBuildingNo
                DoCmd.RepaintObject acForm, "FRM_MNI"
                If TableName = "[HR_ageing and sexing]" Then
                    sql = "SELECT Count([dbo_Exca: Units in buildings].In_Building) AS CountOfIn_Building " & _
                        "FROM ([Exca: Unit Sheet with Relationships] INNER JOIN [dbo_Exca: Units in buildings] ON " & _
                        "[Exca: Unit Sheet with Relationships].[Unit Number] = [dbo_Exca: Units in buildings].Unit) " & _
                        "INNER JOIN " & TableName & " ON [Exca: Unit Sheet with Relationships].[Unit Number] = " & TableName & ".[Unit Number] " & _
                        "GROUP BY [dbo_Exca: Units in buildings].In_Building, " & TableName & "." & FieldName & _
                        " HAVING ((([dbo_Exca: Units in buildings].In_Building)=" & CurrentBuildingNo & ") AND ((" & TableName & "." & WhereClause & ")));"
                Else
                    If InStr(TableName, "Measure") > 0 And AgeCategory <> "adult" Then
                        If AgeCategory = "neonate" Then
                            AdditionalWhere = "[HR_ageing and sexing].[age category]=0 OR [HR_ageing and sexing].[age category]=9"
                        ElseIf AgeCategory = "juvenile" Then
                            AdditionalWhere = "[HR_ageing and sexing].[age category]=1 OR [HR_ageing and sexing].[age category]=2"
                        End If
                        sql = "SELECT Count([dbo_Exca: Units in buildings].In_Building) AS CountOfIn_Building  " & _
                        "FROM (([Exca: Unit Sheet with Relationships] INNER JOIN [dbo_Exca: Units in buildings] ON " & _
                        "[Exca: Unit Sheet with Relationships].[Unit Number] = [dbo_Exca: Units in buildings].Unit) " & _
                        "INNER JOIN " & TableName & "  ON [Exca: Unit Sheet with Relationships].[Unit Number] = " & TableName & ".UnitNumber) " & _
                        "INNER JOIN [HR_ageing and sexing] ON (" & TableName & ".[Individual number] = [HR_ageing and sexing].[Individual number]) " & _
                        " AND (" & TableName & ".UnitNumber = [HR_ageing and sexing].[unit number]) " & _
                        "GROUP BY [dbo_Exca: Units in buildings].In_Building, " & TableName & "." & FieldName & ", [HR_ageing and sexing].[age category] " & _
                        "HAVING ((([dbo_Exca: Units in buildings].In_Building)=" & CurrentBuildingNo & ") AND (" & AdditionalWhere & ") AND ((" & TableName & "." & WhereClause & ")));"
                    Else
                        sql = "SELECT Count([dbo_Exca: Units in buildings].In_Building) AS CountOfIn_Building " & _
                        "FROM ([Exca: Unit Sheet with Relationships] INNER JOIN [dbo_Exca: Units in buildings] ON " & _
                        "[Exca: Unit Sheet with Relationships].[Unit Number] = [dbo_Exca: Units in buildings].Unit) " & _
                        "INNER JOIN " & TableName & " ON [Exca: Unit Sheet with Relationships].[Unit Number] = " & TableName & ".UnitNumber " & _
                        "GROUP BY [dbo_Exca: Units in buildings].In_Building, " & TableName & "." & FieldName & _
                        " HAVING ((([dbo_Exca: Units in buildings].In_Building)=" & CurrentBuildingNo & ") AND ((" & TableName & "." & WhereClause & ")));"
                    End If
                End If
                Set CalcIndividuals = mydb.OpenRecordset(sql, dbOpenSnapshot)
                If Not CalcIndividuals.EOF And Not CalcIndividuals.BOF Then
                    AbleToCalc = True
                    CalcIndividuals.MoveFirst
                    Do Until CalcIndividuals.EOF
                        If AgeCategory = "juvenile" Then
                            If CurrentJuvenileCount < CalcIndividuals![CountOfIn_Building] Then
                                CurrentJuvenileCount = CalcIndividuals![CountOfIn_Building]
                                JuvfldWithMost = FieldName
                            End If
                        ElseIf AgeCategory = "neonate" Then
                            If CurrentNeonateCount < CalcIndividuals![CountOfIn_Building] Then
                                CurrentNeonateCount = CalcIndividuals![CountOfIn_Building]
                                NeofldWithMost = FieldName
                            End If
                        Else 'adult
                            If CurrentAdultCount < CalcIndividuals![CountOfIn_Building] Then
                                CurrentAdultCount = CalcIndividuals![CountOfIn_Building]
                                AdultfldWithMost = FieldName
                            End If
                        End If
                    CalcIndividuals.MoveNext
                    Loop
                Else
                End If
                CalcIndividuals.Close
                Set CalcIndividuals = Nothing
            MNIQueryBuilder.MoveNext
            Loop
            If AbleToCalc = True Then
                sqlInsert = "INSERT INTO [HR_Building_MNI] ([BuildingNumber], [MNI], [Adults], [Juveniles], [Neonates], [Notes], [LastGenerated]) VALUES (" & CurrentBuildingNo & ", " & CurrentAdultCount + CurrentJuvenileCount + CurrentNeonateCount & "," & CurrentAdultCount & "," & CurrentJuvenileCount & ", " & CurrentNeonateCount & ", 'Adult highest count from: " & AdultfldWithMost & ", Juv highest count from: " & JuvfldWithMost & ", Neo highest count from: " & NeofldWithMost & "','" & Now() & "');"
            Else
                sqlInsert = "INSERT INTO [HR_Building_MNI] ([BuildingNumber], [Notes], [LastGenerated]) VALUES (" & CurrentBuildingNo & ", 'Not enough data available yet to undertake calculation', '" & Now() & "');"
            End If
            DoCmd.RunSQL sqlInsert
            MNIQueryBuilder.Close
            Set MNIQueryBuilder = Nothing
        Else
            MsgBox "The MNI Query Builder table which is used to calculate the MNI is empty. The MNI calculation cannot be done without this information. Contact the DBA", vbExclamation, "Cannot proceed"
            MNIQueryBuilder.Close
            Set MNIQueryBuilder = Nothing
            buildings.Close
            Set buildings = Nothing
            Exit Do
        End If
        buildings.MoveNext
    Loop
Else
    MsgBox "No buildings containing skeleton units have been found in the excavation database. No MNI can be calculated.", vbInformation, "No buildings found"
End If
buildings.Close
Set buildings = Nothing
End Sub
