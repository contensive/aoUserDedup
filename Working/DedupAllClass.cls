VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DedupAllClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'========================================================================
'   v3.3 Compatibility
'       To make an Add-on that works the same in v3.3 and v3.4, use this adapter instead of the execute above
'========================================================================
'
Public Function Execute(Csv As Object, Main As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Dim MonthNumber As Long
    Dim DayNumber As Long
    Dim Filename As String
    Dim caption As String
    Dim Description As String
    Dim AdminUI As Object
    Dim okToDedup As Boolean
    Dim DeleteMemberID As Long
    Dim TargetMemberID As Long
    Dim CS As Long
    Dim CSDelete As Long
    Dim KeyField As String
    Dim Common As New CommonClass
    Dim Criteria As String
    Dim PrimaryKeyValue As String
    Dim SecondaryKeyValue As String
    Dim RowCount As Long
    Dim SQL As String
    Dim PrimaryKey As String
    Dim SecondaryKey As String
    Dim UseSecondaryKey As Boolean
    Dim s As String
    Dim runningAsProcess As Boolean
    '
    okToDedup = True
    runningAsProcess = (Main Is Nothing)
    If Not runningAsProcess Then
        okToDedup = False
        If Not Main.isadmin() Then
            '
            '
            '
            Execute = "" _
                & "<div style=""background-color:white; padding: 20px;"">This tool requires administrator rights.</div>" _
                & ""
        Else
            '
            '
            '
            caption = "User DeDup All Tool"
            Description = "" _
                & "<p>Use this tool to automatically remote user duplicates based on email or username, and merge the duplicates into the target record.</p>" _
                & "<p>Please use the User Dedup Individuals tool to test the process used to detemine which record to keep when duplicates are found. Only continue if you understand that this tool DELETES PEOPLE records from your site.</p>" _
                & "<p>The selection of the merge target is based on:" _
                    & "<br>1) a record with a username and password" _
                    & "<br>2) a record marked as administrator" _
                    & "<br>3) the record with the most recent visit" _
                    & "<br>4) the oldest record.</p>"
            '
            ' hardcode to email for now -- can not allow anyone to change it to 'active' or other Db killing blunder
            '
            If (Main.getstreamtext("button") = ButtonOK) Then
                '
                ' they pressed OK
                '
                okToDedup = True
            Else
                '
                ' clicked from teh website, put up a form with an OK button
                '
                Description = Description _
                    & "<p>Press OK to proceed deduping all the people in your database based on the selections in the manual dedup tool.</p>" _
                    & ""
            End If
        End If
    End If
    '
    If okToDedup Then
        '
        ' background process compatible section
        '
        PrimaryKey = Csv.GetSiteProperty(PropertyNamePrimaryFieldName, "email", 0)
        SecondaryKey = Csv.GetSiteProperty(PropertyNameSecondaryFieldName, "", 0)
        UseSecondaryKey = (Trim(SecondaryKey) <> "")
        If (LCase(PrimaryKey) <> "email") And (LCase(PrimaryKey) = "username") Then
            '
            ' Bad Primary field
            '
            s = "<p>Automatic dedup could not run because Primary match field must be either 'email' or 'username'</p>"
        ElseIf (UseSecondaryKey) And (Not Csv.IsSQLTableField("default", "ccmembers", SecondaryKey)) Then
            '
            ' Bad seconary field
            '
            s = "<p>Automatic dedup could not run because the secondary match field [" & SecondaryKey & "] must be in the user table.</p>"
        Else
            '
            ' Select duplicate record sets
            '
            s = s & vbCrLf & Now() & " - Deduping based on the field [" & PrimaryKey & "]"
            RowCount = 0
            If Not UseSecondaryKey Then
                SQL = "select " & PrimaryKey & ",count(id) as cnt from ccMembers group by " & PrimaryKey & " having (count(id)>1)and(" & PrimaryKey & " is not null)"
            Else
                SQL = "select " & PrimaryKey & "," & SecondaryKey & ",count(id) as cnt from ccMembers group by " & PrimaryKey & "," & SecondaryKey & " having (count(id)>1)and(" & PrimaryKey & " is not null)and(" & SecondaryKey & " is not null)"
            End If
            CS = Csv.OpenCSSQL("default", SQL, 0)
            If Not Csv.IsCSOK(CS) Then
                s = s & vbCrLf & "....No duplicates were found"
            Else
                Do While Csv.IsCSOK(CS)
                    PrimaryKeyValue = Trim(Csv.GetCSText(CS, PrimaryKey))
                    If PrimaryKeyValue <> "" Then
                        Criteria = "(" & PrimaryKey & "=" & KmaEncodeSQLText(PrimaryKeyValue) & ")"
                        If UseSecondaryKey Then
                            SecondaryKeyValue = Trim(Csv.GetCSText(CS, SecondaryKey))
                            If SecondaryKeyValue <> "" Then
                                'If IsNumeric(SecondaryKeyValue) Then
                                '    Criteria = Criteria & "AND(" & SecondaryKey & "=" & KmaEncodeSQLNumber(SecondaryKeyValue) & ")"
                                'Else
                                    Criteria = Criteria & "AND(" & SecondaryKey & "=" & KmaEncodeSQLText(SecondaryKeyValue) & ")"
                                'End If
                            End If
                        End If
                        TargetMemberID = Common.GetTargetMemberID(Csv, Criteria)
                        If TargetMemberID <> 0 Then
                            '
                            ' Select a member to remove, and go through all the tables selected
                            '
                            s = s & vbCrLf & "....Deduping [" & PrimaryKeyValue & "], " & Csv.getCSInteger(CS, "cnt") & " records found, record [" & TargetMemberID & "] was selected to keep."
                            CSDelete = Csv.OpenCSContent("people", Criteria & "and(ID<>" & TargetMemberID & ")", , False)
                            Do While Csv.IsCSOK(CSDelete)
                                DeleteMemberID = Csv.getCSInteger(CSDelete, "ID")
                                '
                                SQL = "update ccpagecontent set ContactMemberID=" & TargetMemberID & " where ContactMemberID=" & DeleteMemberID
                                Call Csv.ExecuteSQL("default", SQL)
                                '
                                SQL = "update ccmemberrules set MemberID=" & TargetMemberID & " where MemberID=" & DeleteMemberID
                                Call Csv.ExecuteSQL("default", SQL)
                                '
                                SQL = "update ccemail set TestMemberID=" & TargetMemberID & " where TestMemberID=" & DeleteMemberID
                                Call Csv.ExecuteSQL("default", SQL)
                                '
                                SQL = "Delete from ccMembers where id=" & DeleteMemberID
                                Call Csv.ExecuteSQL("default", SQL)
                                '
                                Call Csv.NextCSRecord(CSDelete)
                            Loop
                            Call Csv.CloseCS(CSDelete)
                        End If
                    End If
                    RowCount = RowCount + 1
                    Call Csv.NextCSRecord(CS)
                Loop
            End If
            Call Csv.CloseCS(CS)
        End If
        '
        ' log the results
        '
        DayNumber = Day(Now)
        MonthNumber = Month(Now)
        Filename = Year(Now)
        If MonthNumber < 10 Then
            Filename = Filename & "0"
        End If
        Filename = Filename & MonthNumber
        If DayNumber < 10 Then
            Filename = Filename & "0"
        End If
        Filename = Filename & DayNumber
        Call Csv.appendVirtualFile("dedup\" & Filename & ".txt", s)
        '
    End If
    '
    '
    '
    If Not runningAsProcess Then
        '
        ' display form or results if not process
        '
        If s <> "" Then
            s = "<div style=""background-color:white;margin:0 20px 0 20px;;padding:20px 20px 20px 20px;""><pre>" & s & "</pre></div>"
        End If
        '
        Execute = "" _
            & "<table width=""100%"" border=""0"" cellpadding=""5"" cellspacing=""0"">" _
            & s _
            & "<tr>" _
            & "<td width=200><img src=""/cclib/images/spacer.gif"" width=200 height=1></td>" _
            & "<td width=""100%"">&nbsp;</td>" _
            & "</tr>" _
            & "</table>"
        Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
        Execute = AdminUI.GetBody(Main, caption, ButtonOK, "", False, False, Description, "", 0, Execute)
        Call Main.AddPageTitle("User DeDup")
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetTargetMemberID")
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
    Call Err.Raise(ErrNumber, ErrSource, App.EXEName & ".AutoProcessClass." & MethodName & " encountered a trap error: " & ErrDescription)
    '
End Sub



