VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ManualClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private PrimaryFieldName As String
Private SecondaryFieldName As String
Private SecondaryFieldId As Long
Private PrimaryFieldValue As String
Private SecondaryFieldValue As String
Private ButtonValue As String
Private TargetID As Long

Private ErrorString As String

Private Main As Object
Private Csv As Object
'Private Main As ccWeb3.MainClass
'
'========================================================================
'   v3.3 Compatibility
'       To make an Add-on that works the same in v3.3 and v3.4, use this adapter instead of the execute above
'========================================================================
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    If Not (MainObject Is Nothing) Then
        If Not MainObject.isadmin() Then
            Execute = "" _
                & "<div style=""background-color:white; padding: 20px;"">This tool requires administrator rights.</div>" _
                & ""
        Else
            Set Csv = CsvObject
            Call Init(MainObject)
            Execute = GetContent(OptionString)
        End If
    End If
End Function
'
'
'
Public Sub Init(MainObject As Object)
    Set Main = MainObject
    Exit Sub
ErrorTrap:
    Call HandleClassTrapError("Init")
    'Call HandleError("LoginClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
End Sub
'
'
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    '
    If Not (Main Is Nothing) Then
        If Not Main.isadmin() Then
            GetContent = "<div style=""padding:20px"">You must be an administrator to use the User DeDup Tools</div>"
        Else
            Call Main.SetStreamTimeout(2000)
            'Call LoadForm
            'Call ProcessForm
            GetContent = GetForm
        End If
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetContent")
    'Call HandleError("BenefitClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
End Function

Private Function GetForm() As String
    On Error GoTo ErrorTrap
    '
    Dim PrimaryFieldListId As Long
    Dim KeyValue As String
    Dim LcaseTableName As String
    Dim caption As String
    Dim Description As String
    Dim AdminUI As Object
    Dim ButtonList As String
    Dim ManualLink As String
    Dim RQS As String
    Dim Checkbox As String
    Dim Tables() As String
    Dim Ptr As Long
    Dim MemberID As Long
    'Dim TableList As String
    Dim TargetMemberID As Long
    Dim RequestName As String
    Dim SQL As String
    Dim CS As Long
    Dim TableName As String
    Dim TableCnt As Long
    Dim MemberCID As Long
    Dim Stream As String
    Dim InnerStream As String
    Dim CSPointer As Long
    Dim Criteria As String
    Dim locTarget As Boolean
    Dim Style As String
    Dim RowCount As Long
    Dim Results As String
    Dim FormID As Long
    Dim Hint As String
    Dim Commonx As New CommonClass
    Dim Csv As Object
    Dim CSConn As Variant
    '
    Set Csv = CreateObject("ccCsrvr3.ContentServerClass")
    '
    caption = "User DeDup Tool"
    Description = "" _
        & "<p>Use this tool to search for user duplicates, isolate the record which best describes the user, and merge the duplicates into the target record.</p>" _
        & "<p>The merge target is highlighted in red. This is the user record that will be preserved. All others will be removed, and all references to these records will be transfered to the target.</p>" _
        & "<p>The selection of the merge target is based on: 1) a record with a username and password, then 2) a record marked as administrator, then 3) the record with the most recent visit, and finally, 4) the oldest record.</p>"
    '
    CSConn = Csv.OpenConnection(Main.ApplicationName)
    If CSConn.ApplicationStatus = 2 Then
    'If CSConn.ApplicationStatus = ApplicationStatusRunning Then
    
        '
        ButtonValue = Main.getstreamtext(RNButton)
        If ButtonValue = FormButtonCancel Then
            'Hint = "200"
            '
            ' return empty cancels addon
            '
        Else
            'Hint = "300"
            FormID = Main.GetStreaminteger("formid")
            If (FormID = 0) Then
                'Hint = "400"
                '
                ' set defaults for search form
                '
                PrimaryFieldName = Main.GetSiteProperty(PropertyNamePrimaryFieldName, "email")
                PrimaryFieldListId = GetFieldId(PrimaryFieldName, PrimaryFieldList)
                PrimaryFieldValue = Main.GetSiteProperty(PropertyNamePrimaryFieldValue)
                '
                SecondaryFieldName = Main.GetSiteProperty(PropertyNameSecondaryFieldName, "")
                SecondaryFieldValue = Main.GetSiteProperty(PropertyNameSecondaryFieldValue)
                '
                FormID = FormSearch
            ElseIf (FormID = FormSearchList) Then
                'Hint = "500"
                '
                ' Process Search List Form
                '
                If ButtonValue = Formbuttonrefresh Then
                    '
                    ' Do a refresh -- just leave the formid alone
                    '
                    PrimaryFieldName = Main.GetSiteProperty(PropertyNamePrimaryFieldName)
                    PrimaryFieldListId = GetFieldId(PrimaryFieldName, PrimaryFieldList)
                    SecondaryFieldName = Main.GetSiteProperty(PropertyNameSecondaryFieldName)
                ElseIf ButtonValue = FormButtonNewSearch Then
                    PrimaryFieldName = Main.GetSiteProperty(PropertyNamePrimaryFieldName)
                    PrimaryFieldListId = GetFieldId(PrimaryFieldName, PrimaryFieldList)
                    PrimaryFieldValue = Main.GetSiteProperty(PropertyNamePrimaryFieldValue)
                    '
                    SecondaryFieldName = Main.GetSiteProperty(PropertyNameSecondaryFieldName)
                    SecondaryFieldValue = Main.GetSiteProperty(PropertyNameSecondaryFieldValue)
                    '
                    'TableList = Main.GetSiteProperty("DeDupSearchTableList")
                    FormID = FormSearch
                Else
                    '
                    ' hit from the dedup link
                    '
                    PrimaryFieldListId = Main.GetStreaminteger(RNPrimaryFieldListId)
                    PrimaryFieldName = GetFieldName(PrimaryFieldListId, PrimaryFieldList)
                    PrimaryFieldValue = Main.getstreamtext(RNPrimaryFieldValue)
                    '
                    SecondaryFieldId = Main.GetStreaminteger(RNSecondaryFieldId)
                    SecondaryFieldName = Main.getrecordname("content fields", SecondaryFieldId)
                    SecondaryFieldValue = Main.getstreamtext(RNSecondaryFieldValue)
                    '
                    FormID = FormSearchResults
                End If
            ElseIf (FormID = FormSearch) Then
                'Hint = "500"
                '
                ' Process Search Form
                '
                PrimaryFieldListId = Main.GetStreaminteger(RNPrimaryFieldListId)
                PrimaryFieldName = GetFieldName(PrimaryFieldListId, PrimaryFieldList)
                PrimaryFieldValue = Main.getstreamtext(RNPrimaryFieldValue)
                '
                SecondaryFieldId = Main.GetStreaminteger(RNSecondaryFieldId)
                SecondaryFieldName = Main.getrecordname("content fields", SecondaryFieldId)
                SecondaryFieldValue = Main.getstreamtext(RNSecondaryFieldValue)
                '
                Call Main.SetSiteProperty(PropertyNamePrimaryFieldName, CStr(PrimaryFieldName))
                Call Main.SetSiteProperty(PropertyNamePrimaryFieldValue, CStr(PrimaryFieldValue))
                Call Main.SetSiteProperty(PropertyNameSecondaryFieldName, CStr(SecondaryFieldName))
                Call Main.SetSiteProperty(PropertyNameSecondaryFieldValue, CStr(SecondaryFieldValue))
                
                If PrimaryFieldName = "" Then
                    ErrorString = ErrorString & "Please select a Search Field. This is the field that will be searched to locate duplicates. For instance, you believe there are many records with the same email address, select 'email' as the Search Field, and the address of the user with duplicates in the Search Value.<br />"
                    FormID = FormSearch
                Else
                    If PrimaryFieldValue = "" Then
                        FormID = FormSearchList
                    Else
                        FormID = FormSearchResults
                    End If
                End If
            ElseIf (FormID = FormSearchResults) Then
                'Hint = "600"
                '
                ' Process Search Results Form - target selected, now dedup
                '
                If ButtonValue = FormButtonNewSearch Then
                    'Hint = "700"
                    PrimaryFieldName = Main.GetSiteProperty(PropertyNamePrimaryFieldName)
                    PrimaryFieldListId = GetFieldId(PrimaryFieldName, PrimaryFieldList)
                    PrimaryFieldValue = Main.GetSiteProperty(PropertyNamePrimaryFieldValue)
                    '
                    SecondaryFieldName = Main.GetSiteProperty(PropertyNameSecondaryFieldName)
                    SecondaryFieldValue = Main.GetSiteProperty(PropertyNameSecondaryFieldValue)
                    '
                    'TableList = Main.GetSiteProperty("DeDupSearchTableList")
                    FormID = FormSearch
                Else
                    'Hint = "800"
                    PrimaryFieldListId = Main.GetStreaminteger(RNPrimaryFieldListId)
                    PrimaryFieldName = GetFieldName(PrimaryFieldListId, PrimaryFieldList)
                    PrimaryFieldValue = Main.getstreamtext(RNPrimaryFieldValue)
                    '
                    SecondaryFieldId = Main.GetStreaminteger(RNSecondaryFieldId)
                    SecondaryFieldName = Main.getrecordname("content fields", SecondaryFieldId)
                    SecondaryFieldValue = Main.getstreamtext(RNSecondaryFieldValue)
                    '
                    TargetID = Main.GetStreaminteger(RNTarget)
                    'TableList = Main.GetStreamText("TableList")
                    
                    If TargetID = 0 Then
                        FormID = FormSearchResults
                        ErrorString = ErrorString & "There was a problem identifying your target. It was zero and this is not allowed.<br />"
                    ElseIf PrimaryFieldName = "" Then
                        FormID = FormSearchResults
                        ErrorString = ErrorString & "There was a problem with the Search Field. It was blank and this is not allowed.<br />"
                    ElseIf PrimaryFieldValue = "" Then
                        FormID = FormSearchResults
                        ErrorString = ErrorString & "There was a problem with the Search Value. It was blank and this is not allowed.<br />"
                    Else
                        If LCase(PrimaryFieldValue) = "true" Then
                            Criteria = PrimaryFieldName & "<>0"
                            'Criteria = KeyField & "<>0"
                        ElseIf LCase(PrimaryFieldValue) = "false" Then
                            Criteria = "(" & PrimaryFieldName & "=0)or(" & PrimaryFieldName & " is null)"
                        ElseIf IsNumeric(PrimaryFieldValue) Then
                            Criteria = PrimaryFieldName & "=" & Main.EncodeSQLNumber(PrimaryFieldValue)
                        Else
                            Criteria = PrimaryFieldName & "=" & Main.EncodeSQLText(PrimaryFieldValue)
                        End If
                        If (SecondaryFieldName <> "") And (SecondaryFieldValue <> "") Then
                            Criteria = Criteria & "AND(" & SecondaryFieldName & "=" & Main.EncodeSQLText(SecondaryFieldValue) & ")"
                        End If
                        TargetMemberID = TargetID
                        '
                        ' Select a member to remove, and go through all the tables selected
                        '
                        CS = Main.OpenCSContent("people", "(" & Criteria & ")and(ID<>" & TargetMemberID & ")")
                        Do While Main.IsCSOK(CS)
                            MemberID = Main.getCSInteger(CS, "ID")
                            '
                            SQL = "update organizations set ContactMemberID=" & TargetMemberID & " where ContactMemberID=" & MemberID
                            Call Csv.ExecuteSQL("default", SQL)
                            '
                            SQL = "update ccpagecontent set ContactMemberID=" & TargetMemberID & " where ContactMemberID=" & MemberID
                            Call Csv.ExecuteSQL("default", SQL)
                            '
                            SQL = "update ccmemberrules set MemberID=" & TargetMemberID & " where MemberID=" & MemberID
                            Call Csv.ExecuteSQL("default", SQL)
                            '
                            SQL = "update ccemail set TestMemberID=" & TargetMemberID & " where TestMemberID=" & MemberID
                            Call Csv.ExecuteSQL("default", SQL)
                            '
                            SQL = "update ccemaillog set MemberID=" & TargetMemberID & " where MemberID=" & MemberID
                            Call Csv.ExecuteSQL("default", SQL)
                            '
                            If Csv.isSqlTableField("", "orders", "id") Then
                                SQL = "update orders set MemberID=" & TargetMemberID & " where MemberID=" & MemberID
                                Call Csv.ExecuteSQL("default", SQL)
                            End If
                            '
                            Call Main.NextCSRecord(CS)
                        Loop
                        Call Main.CloseCS(CS)
                        '
                        SQL = "Delete from ccMembers where " & Criteria & " and ID<>" & TargetID
                        Call Main.ExecuteSQL("default", SQL)
                        FormID = FormDeDupResults
                    End If
                End If
            ElseIf (FormID = FormDeDupResults) Then
                'Hint = "810"
                '
                ' Process de-Dup results form
                '
                If ButtonValue = FormButtonNewSearch Then
                    PrimaryFieldName = Main.GetSiteProperty(PropertyNamePrimaryFieldName)
                    PrimaryFieldListId = GetFieldId(PrimaryFieldName, PrimaryFieldList)
                    PrimaryFieldValue = Main.GetSiteProperty(PropertyNamePrimaryFieldValue)
                    '
                    SecondaryFieldName = Main.GetSiteProperty(PropertyNameSecondaryFieldName)
                    SecondaryFieldValue = Main.GetSiteProperty(PropertyNameSecondaryFieldValue)
                    '
                    FormID = FormSearch
                Else
                    '
                    ' Show results again
                    '
                    FormID = FormDeDupResults
                End If
            End If
            '
            '   Display next form
            '
            If FormID = FormSearchResults Then
                'Hint = "820"
                '
                ' ----- Search Results form
                '
                'If ErrorString = "" Then
                    RowCount = 0
                    If LCase(PrimaryFieldValue) = "true" Then
                        Criteria = "(" & PrimaryFieldName & "<>0)"
                    ElseIf LCase(PrimaryFieldValue) = "false" Then
                        Criteria = "((" & PrimaryFieldName & "=0)or(" & PrimaryFieldName & " is null))"
                    ElseIf IsNumeric(PrimaryFieldValue) Then
                        Criteria = "(" & PrimaryFieldName & "=" & Main.EncodeSQLNumber(PrimaryFieldValue) & ")"
                    Else
                        Criteria = "(" & PrimaryFieldName & "=" & Main.EncodeSQLText(PrimaryFieldValue) & ")"
                    End If
                    If (SecondaryFieldName <> "") And (SecondaryFieldValue <> "") Then
                        Criteria = Criteria & "AND(" & SecondaryFieldName & "=" & Main.EncodeSQLText(SecondaryFieldValue) & ")"
                    End If
                    CSPointer = Main.OpenCSContent(CNPeople, Criteria, "ID", False)
                    If Main.CSOK(CSPointer) Then
                        Results = Results & "<table width=""100%"" border=""0"" cellpadding=""5"" cellspacing=""0"" class=ccAdminHeader>"
                        Results = Results & "<tr>"
                        Results = Results & "<td class=ccAdminHeader><b></b></td>"
                        Results = Results & "<td class=ccAdminHeader><b></b></td>"
                        Results = Results & "<td class=ccAdminHeader><b>Key Field(s)</b></td>"
                        Results = Results & "<td class=ccAdminHeader><b>ID</b></td>"
                        Results = Results & "<td class=ccAdminHeader><b>Edit</b></td>"
                        Results = Results & "<td class=ccAdminHeader><b>Active</b></td>"
                        Results = Results & "<td class=ccAdminHeader><b>Username</b></td>"
                        'Results = Results  & "<td class=ccAdminHeader><b>Password</b></td>"
                        Results = Results & "<td class=ccAdminHeader><b>First Name</b></td>"
                        Results = Results & "<td class=ccAdminHeader><b>Last Name</b></td>"
                        Results = Results & "<td class=ccAdminHeader><b>Company</b></td>"
                        Results = Results & "<td class=ccAdminHeader><b>Email</b></td>"
                        Results = Results & "<td class=ccAdminHeader><b>Phone</b></td>"
                        Results = Results & "</tr>"
                        TargetMemberID = Commonx.GetTargetMemberID(Csv, Criteria)
                        Do While Main.CSOK(CSPointer)
                            MemberID = Main.getCSInteger(CSPointer, "ID")
                            locTarget = (TargetMemberID = MemberID)
                            'locTarget = isTarget(Main.GetCSInteger(CSPointer, "ID"))
                            If locTarget Then
                                Style = "background-color: #CF8686;color:white"
                            Else
                                If (RowCount Mod 2) = 1 Then
                                    Style = "background-color: #EBEBEB;color:black"
                                Else
                                    Style = "background-color: white;color:black"
                                    End If
                                End If
                            KeyValue = Main.GetCS(CSPointer, PrimaryFieldName)
                            If (SecondaryFieldName <> "") And (SecondaryFieldValue <> "") Then
                                KeyValue = KeyValue & "," & Main.GetCS(CSPointer, SecondaryFieldName)
                            End If
                            Results = Results & "<tr>"
                            Results = Results & "<td align=""right"" style=""" & Style & """>" & RowCount + 1 & "</td>"
                            Results = Results & "<td align=""right"" style=""" & Style & """>" & Main.GetFormInputRadioBox(RNTarget, MemberID, TargetMemberID) & "</td>"
                            Results = Results & "<td style=""" & Style & """><b>" & KeyValue & "</b></td>"
                            Results = Results & "<td style=""" & Style & """>" & Main.GetCS(CSPointer, "ID") & "</td>"
                            Results = Results & "<td style=""" & Style & """><a href=""?cid=" & Main.getCSInteger(CSPointer, "contentcontrolid") & "&id=" & MemberID & "&af=4"" target=""_blank"">edit</a></td>"
                            Results = Results & "<td style=""" & Style & """>" & Main.GetCS(CSPointer, "Active") & "</td>"
                            Results = Results & "<td style=""" & Style & """>" & Main.GetCS(CSPointer, "Username") & "</td>"
                            'Results = Results  & "<td style=""" & Style & """>" & Main.GetCS(CSPointer, "Password") & "</td>"
                            Results = Results & "<td style=""" & Style & """>" & Main.GetCS(CSPointer, "FirstName") & "</td>"
                            Results = Results & "<td style=""" & Style & """>" & Main.GetCS(CSPointer, "LastName") & "</td>"
                            Results = Results & "<td style=""" & Style & """>" & Main.GetCS(CSPointer, "Company") & "</td>"
                            Results = Results & "<td style=""" & Style & """>" & Main.GetCS(CSPointer, "email") & "</td>"
                            Results = Results & "<td style=""" & Style & """>" & Main.GetCS(CSPointer, "Phone") & "</td>"
                            Results = Results & "</tr>"
                            Call Main.NextCSRecord(CSPointer)
                            RowCount = RowCount + 1
                        Loop
                        Results = Results & "</table>"
                    End If
                    Call Main.CloseCS(CSPointer)
                    'Hint = "830"
    '                If ErrorString <> "" Then
    '                    Stream = Stream & "<tr>"
    '                    Stream = Stream & "<td class=""ccError"" style=""padding:20px"" colspan=""2"" width=""100%"">" & ErrorString & "</td>"
    '                    Stream = Stream & "</tr>"
    '                End If
                    If RowCount = 0 Then
                        FormID = FormSearch
                        ErrorString = ErrorString & "Your search (" & PrimaryFieldName & "='" & PrimaryFieldValue & "') returned no results.<br />"
                    ElseIf RowCount = 1 Then
                        Stream = Stream & "<tr><td colspan=2 style=""padding:20px"">" _
                            & "<p>Your search (" & PrimaryFieldName & "='" & PrimaryFieldValue & "') returned only one result, so does not need to be De-Duped.</p>" _
                            & "</td></tr>"
                        Stream = Stream & "<tr><td colspan=""2"" width=""100%"">" & Results & "</td></tr>"
                        'Stream = Stream & "<tr>"
                        'Stream = Stream & "<td><img src=""/ccLib/images/spacer.gif"" width=""100"" height=""1""></td>"
                        'Stream = Stream & "<td width=""100%"">" & Main.GetFormButton(FormButtonCancel) & Main.GetFormButton(FormButtonNewSearch) & "</td>"
                        'Stream = Stream & "</tr>"
                        Stream = Stream & Main.GetFormInputHidden("formid", FormSearchResults)
                        ButtonList = FormButtonCancel & "," & FormButtonNewSearch
                    Else
                        Description = Description _
                            & "<p>The following is a list of all records that match your search (" & PrimaryFieldName & "='" & PrimaryFieldValue & "'). In the first column is a checkbox you use to select the record which best matches the user. All others records will be removed, leaving only one record for this user.</p>" _
                            & "<p>Select the site tables that should be updated to reflect the new user.</p>"
                        Stream = Stream & "<tr><td colspan=""2"" width=""100%"">" & Results & "</td></tr>"
'                        '
'                        ' List all tables with a MemberID field, or special case
'                        '
'                        Stream = Stream & "<tr><td colspan=""2"" width=""100%""><b>Select the tables to be updated with the Target user where a duplicate is found.</b> Members can only be updated in these tables if the field is named 'memberid'. This is the case for all Contensive people fields. If there are other fields that need to be updated but have different names, the updates must be handled manually and you should not use the dedup tool.</td></tr>"
'                        CS = Main.OpenCSContent("Tables")
'                        TableCnt = 0
'                        Do While Main.IsCSOK(CS)
'                            TableName = Main.GetCSText(CS, "name")
'                            If ((TableList = "") Or (InStr(1, "," & TableList & ", ", "," & TableName & ", ", vbTextCompare) <> 0)) And (InStr(1, "," & DefaultOffTables & ", ", "," & TableName & ",", vbTextCompare) = 0) Then
'                                Checkbox = "<input type=checkbox name=TableList value=""" & TableName & """ checked>"
'                            Else
'                                Checkbox = "<input type=checkbox name=TableList value=""" & TableName & """>"
'                            End If
'                            If (InStr(1, SpecialCaseTables, TableName, vbTextCompare) <> 0) Or Main.IsSQLTableField("default", TableName, "memberid") Then
'                                Stream = Stream & "<tr>"
'                                Stream = Stream & "<td align=right >" & Checkbox & "</td>"
'                                Stream = Stream & "<td width=""100%"">" & TableName & "</td>"
'                                Stream = Stream & "</tr>"
'                                TableCnt = TableCnt + 1
'                            End If
'                            Main.NextCSRecord (CS)
'                        Loop
'                        Main.CloseCS (CS)
                        'Stream = Stream & "<tr>"
                        'Stream = Stream & "<td><img src=""/ccLib/images/spacer.gif"" width=""100"" height=""1""></td>"
                        'Stream = Stream & "<td width=""100%"">" & Main.GetFormButton(FormButtonCancel) & Main.GetFormButton(FormButtonNewSearch) & Main.GetFormButton(FormButtonDeDupUser) & "</td>"
                        'Stream = Stream & "</tr>"
                        ButtonList = FormButtonCancel & "," & FormButtonNewSearch & "," & FormButtonDeDupUser
                        Stream = Stream & Main.GetFormInputHidden(RNPrimaryFieldListId, PrimaryFieldListId)
                        Stream = Stream & Main.GetFormInputHidden(RNPrimaryFieldValue, PrimaryFieldValue)
                        SecondaryFieldId = Main.getrecordid("content fields", SecondaryFieldName)
                        Stream = Stream & Main.GetFormInputHidden(RNSecondaryFieldId, SecondaryFieldId)
                        Stream = Stream & Main.GetFormInputHidden(RNSecondaryFieldValue, SecondaryFieldValue)
                        Stream = Stream & Main.GetFormInputHidden("formid", FormSearchResults)
                    End If
                'End If
            End If
            If (FormID = FormSearchList) Then
                '
                ' ----- Search for duplicates, and list them
                '
                RowCount = 0
                If SecondaryFieldName <> "" Then
                    SQL = "select top 100 " & PrimaryFieldName & "," & SecondaryFieldName & ",count(id) as cnt from ccMembers where (contentcontrolid<>0) group by " & PrimaryFieldName & "," & SecondaryFieldName & " having (count(id)>1)and(" & PrimaryFieldName & " is not null)and(" & SecondaryFieldName & " is not null)"
                Else
                    SQL = "select top 100 " & PrimaryFieldName & ",count(id) as cnt from ccMembers where (contentcontrolid<>0) group by " & PrimaryFieldName & " having (count(id)>1)and(" & PrimaryFieldName & " is not null)"
                End If
                CS = Main.OpenCSSQL("default", SQL)
                If Main.IsCSOK(CS) Then
                    Stream = Stream & "<tr><td colspan=2 style=""padding:20px;"">"
                    '
                    Stream = Stream & "<table width=""100%"" border=""0"" cellpadding=""5"" cellspacing=""0"" class=ccAdminHeader>"
                    Stream = Stream & "<tr>"
                    Stream = Stream & "<td width=""50px"" class=ccAdminHeader>&nbsp;</td>"
                    'Stream = Stream & "<td width=""50px"" class=ccAdminHeader>&nbsp;</td>"
                    Stream = Stream & "<td width=""100px"" class=ccAdminHeader><b>Manual&nbsp;Dedup</b></td>"

                    If SecondaryFieldName <> "" Then
                        Stream = Stream & "<td width=""150px"" class=ccAdminHeader><b>Primary Field - " & PrimaryFieldName & "</b><br><img src=""/cclib/images/spacer.gif"" width=150 height=1></td>"
                        Stream = Stream & "<td width=""150px"" class=ccAdminHeader><b>Secondary Field - " & SecondaryFieldName & "</b><br><img src=""/cclib/images/spacer.gif"" width=150 height=1></td>"
                    Else
                        Stream = Stream & "<td width=""300px"" class=ccAdminHeader><b>Key Field - " & PrimaryFieldName & "</b><br><img src=""/cclib/images/spacer.gif"" width=300 height=1></td>"
                    End If
                    Stream = Stream & "<td width=""99%"" class=ccAdminHeader><b>Count</b></td>"
                    Stream = Stream & "</tr>"
                    'TargetMemberID = GetTargetMemberID()
                    RQS = Main.RefreshQueryString
                    Do While Main.IsCSOK(CS)
                        If SecondaryFieldName <> "" Then
                            PrimaryFieldValue = Main.GetCSText(CS, PrimaryFieldName)
                            SecondaryFieldValue = Main.GetCSText(CS, SecondaryFieldName)
                        Else
                            PrimaryFieldValue = Main.GetCSText(CS, PrimaryFieldName)
                        End If
                        ManualLink = RQS
                        ManualLink = ModifyQueryString(ManualLink, "formid", FormSearchList)
                        ManualLink = ModifyQueryString(ManualLink, RNPrimaryFieldListId, CStr(PrimaryFieldListId))
                        ManualLink = ModifyQueryString(ManualLink, RNPrimaryFieldValue, PrimaryFieldValue)
                        If SecondaryFieldName <> "" Then
                            
                            SecondaryFieldId = Main.getrecordid("content fields", SecondaryFieldName)
                            ManualLink = ModifyQueryString(ManualLink, RNSecondaryFieldId, CStr(SecondaryFieldId))
                            ManualLink = ModifyQueryString(ManualLink, RNSecondaryFieldValue, SecondaryFieldValue)
                        End If
                        If (RowCount Mod 2) = 1 Then
                            Style = "background-color: #EBEBEB;color:black"
                        Else
                            Style = "background-color: white;color:black"
                        End If
                        Stream = Stream & "<tr>"
                        Stream = Stream & "<td align=""right"" style=""" & Style & """>" & RowCount + 1 & "</td>"
                        'Stream = Stream & "<td align=""right"" style=""" & Style & """><input type=checkbox name=PrimaryFieldValue value=""" & kmaEncodeHTML(PrimaryFieldValue) & """></td>"
                        Stream = Stream & "<td style=""" & Style & """><a href=""?" & ManualLink & """>Dedup</a></td>"
                        If SecondaryFieldName <> "" Then
                            Stream = Stream & "<td style=""" & Style & """>" & PrimaryFieldValue & "&nbsp;" & "</td>"
                            Stream = Stream & "<td style=""" & Style & """>" & SecondaryFieldValue & "&nbsp;" & "</td>"
                        Else
                            Stream = Stream & "<td style=""" & Style & """>" & PrimaryFieldValue & "&nbsp;" & "</td>"
                        End If
                        Stream = Stream & "<td style=""" & Style & """>" & Main.GetCS(CS, "cnt") & "</td>"
                        Stream = Stream & "</tr>"
                        RowCount = RowCount + 1
                        Call Main.NextCSRecord(CS)
                    Loop
                    Stream = Stream & "</table>"
                    Stream = Stream & "</td></tr>"
                End If
                Call Main.CloseCS(CS)
                If RowCount = 0 Then
                    Stream = Stream & "<tr><td colspan=2 style=""padding:20px;"">There were no duplicates found.</td></tr>"
                End If
                ButtonList = FormButtonCancel & "," & FormButtonNewSearch & "," & Formbuttonrefresh
                Stream = Stream & Main.GetFormInputHidden("formid", FormSearchList)
            End If
            If (FormID = FormDeDupResults) Then
                'Hint = "835"
                '
                ' ----- deDup results
                '
                Stream = Stream & "<tr>"
                Stream = Stream & "<td colspan=2 style=""padding:20px;"">This user has been De-Dupped</td>"
                Stream = Stream & "</tr>"
                ButtonList = FormButtonCancel & "," & FormButtonNewSearch
                Stream = Stream & Main.GetFormInputHidden("formid", FormDeDupResults)
            End If
            If (FormID = FormSearch) Then
                'Hint = "840"
                '
                ' ----- Search form
                '
                PrimaryFieldListId = GetFieldId(PrimaryFieldName, PrimaryFieldList)
                SecondaryFieldId = Main.getrecordid("content fields", SecondaryFieldName)
                MemberCID = Main.GetContentID(CNPeople)
                Description = Description _
                    & "<p>To search for a user, select the field you want to search on, and the value you will search for. All results matching this selection will be displayed.</p>"
                Stream = Stream & "<tr>"
                Stream = Stream & "<td align=""right""><b>Primary Search Field:</b></td>"
                Stream = Stream & "<td width=""100%"">" & Main.GetFormInputSelectList(RNPrimaryFieldListId, PrimaryFieldListId, PrimaryFieldList) & "</td>"
                Stream = Stream & "</tr>"
                Stream = Stream & "<tr>"
                Stream = Stream & "<td align=""right""><b>Primary Search Value:</b></td>"
                Stream = Stream & "<td width=""100%"">" & Main.GetFormInputText(RNPrimaryFieldValue, PrimaryFieldValue, , 35) & "</td>"
                Stream = Stream & "</tr>"
                '
If False Then
                Stream = Stream & "<tr><td colspan=2>&nbsp;</td></tr>"
                '
                Stream = Stream & "<tr>"
                Stream = Stream & "<td align=""right""><b>Secondary Search Field:</b></td>"
                Stream = Stream & "<td width=""100%"">" & Main.GetFormInputSelect(RNSecondaryFieldId, SecondaryFieldId, CNFields, "ContentID=" & MemberCID) & "</td>"
                Stream = Stream & "</tr>"
                Stream = Stream & "</tr>"
                Stream = Stream & "<tr>"
                Stream = Stream & "<td align=""right""><b>Secondary Search Value:</b></td>"
                Stream = Stream & "<td width=""100%"">" & Main.GetFormInputText(RNSecondaryFieldValue, SecondaryFieldValue, , 35) & "</td>"
                Stream = Stream & "</tr>"
End If
                '
                ButtonList = FormButtonCancel & "," & FormButtonSearch
                Stream = Stream & Main.GetFormInputHidden("formid", FormSearch)
            End If
            '
            '
            If Stream <> "" Then
                GetForm = "" _
                    & "<table width=""100%"" border=""0"" cellpadding=""5"" cellspacing=""0"">" _
                    & Stream _
                    & "<tr>" _
                    & "<td width=200><img src=""/cclib/images/spacer.gif"" width=200 height=1></td>" _
                    & "<td width=""100%"">&nbsp;</td>" _
                    & "</tr>" _
                    & "</table>"
                If ErrorString <> "" Then
                    Description = Description & "<p class=ccError>" & ErrorString & "</p>"
                End If
                Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
                GetForm = AdminUI.GetBody(Main, caption, ButtonList, "", False, False, Description, "", 0, GetForm)
                Call Main.AddPageTitle("User DeDup")
            End If
        End If
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetForm, Hint='" & Hint & "'")
    'Call HandleError("CleanClass", "GetForm", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'========================================================================
'   Handle Class Error
'========================================================================
'
Private Sub HandleClassTrapError(MethodName As String)
    '
    Dim ErrNumber As Long
    Dim ErrSource As String
    Dim ErrDescription As String
    Dim ErrMsg As String
    '
    ErrNumber = Err.Number
    ErrSource = Err.Source
    ErrDescription = Err.Description
    ErrMsg = GetErrString(Err)
    Err.Clear
    On Error GoTo 0
    Call Err.Raise(ErrNumber, ErrSource, App.EXEName & ".ManualClass." & MethodName & " encountered a trap error: " & ErrDescription)
    '
End Sub
'
'
'
Private Function GetFieldName(Ptr As Long, FieldList As String) As String
    Dim fields() As String
    '
    If Ptr > 0 Then
        fields = Split(FieldList, ",")
        If Ptr <= (UBound(fields) + 1) Then
            GetFieldName = fields(Ptr - 1)
        End If
    End If
End Function
'
'
'
Private Function GetFieldId(FieldName As String, FieldList As String) As String
    '
    Dim Ptr As Long
    Dim t As String
    Dim fields() As String
    '
    GetFieldId = 0
    If FieldName <> "" Then
    fields = Split(LCase(FieldList), ",")
    t = LCase(FieldName)
    If UBound(fields) > 0 Then
        For Ptr = 0 To UBound(fields)
            If t = fields(Ptr) Then
                GetFieldId = Ptr + 1
            End If
        Next
    End If
    End If
End Function
'
'
'
Private Property Get PrimaryFieldList()
    PrimaryFieldList = Main.GetSiteProperty(PropertyNamePrimaryFieldList, PrimaryFieldListDefault)
    If PrimaryFieldList = "" Then
        PrimaryFieldList = PrimaryFieldListDefault
        Call Main.SetSiteProperty(PropertyNamePrimaryFieldList, PrimaryFieldList)
    End If
End Property

