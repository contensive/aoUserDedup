VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoProcessClass"
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
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Execute = Main(CsvObject, OptionString)
End Function
'
'
'
Public Function Main(Csv As Object, OptionString As String) As String
    On Error GoTo ErrorTrap
    '
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
    '
    ' hardcode to email for now -- can not allow anyone to change it to 'active' or other Db killing blunder
    '
    If kmaEncodeBoolean(Csv.GetSiteProperty("UserdedupProcessEnabled", "0", 0)) Then
        PrimaryKey = Csv.GetSiteProperty(PropertyNamePrimaryFieldName, "email", 0)
        SecondaryKey = Csv.GetSiteProperty(PropertyNameSecondaryFieldName, "", 0)
        UseSecondaryKey = (Trim(SecondaryKey) <> "")
        If (LCase(PrimaryKey) <> "email") And (LCase(PrimaryKey) = "username") Then
            '
            ' Bad Primary field
            '
            Main = "Automatic dedup could not run because Primary match field must be either 'email' or 'username'"
        ElseIf (UseSecondaryKey) And (Not Csv.isSqlTableField("default", "ccmembers", SecondaryKey)) Then
            '
            ' Bad seconary field
            '
            Main = "Automatic dedup could not run because the secondary match field [" & SecondaryKey & "] must be in the user table."
        Else
            '
            ' Select duplicate record sets
            '
            RowCount = 0
            If Not UseSecondaryKey Then
                SQL = "select " & PrimaryKey & ",count(id) as cnt from ccMembers group by " & PrimaryKey & " having (count(id)>1)and(" & PrimaryKey & " is not null)"
            Else
                SQL = "select " & PrimaryKey & "," & SecondaryKey & ",count(id) as cnt from ccMembers group by " & PrimaryKey & "," & SecondaryKey & " having (count(id)>1)and(" & PrimaryKey & " is not null)and(" & SecondaryKey & " is not null)"
            End If
            CS = Csv.OpenCSSQL("default", SQL, 0)
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
                            SQL = "update organizations set ContactMemberID=" & TargetMemberID & " where ContactMemberID=" & DeleteMemberID
                            Call Csv.ExecuteSQL("default", SQL)
                            '
                            SQL = "update ccemaillog set MemberID=" & TargetMemberID & " where MemberID=" & DeleteMemberID
                            Call Csv.ExecuteSQL("default", SQL)
                            '
                            SQL = "Delete from ccMembers where id=" & DeleteMemberID
                            Call Csv.ExecuteSQL("default", SQL)
                            '
                            If Csv.isSqlTableField("", "orders", "id") Then
                                SQL = "update orders set MemberID=" & TargetMemberID & " where MemberID=" & DeleteMemberID
                                Call Csv.ExecuteSQL("default", SQL)
                            End If
                            '
                            Call Csv.NextCSRecord(CSDelete)
                        Loop
                        Call Csv.CloseCS(CSDelete)
                    End If
                End If
                RowCount = RowCount + 1
                Call Csv.NextCSRecord(CS)
            Loop
            Call Csv.CloseCS(CS)
        End If
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


