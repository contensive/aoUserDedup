VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'
'
Friend Function GetTargetMemberID(Csv As Object, BaseCriteria As String) As Long
    On Error GoTo ErrorTrap
    '
    Dim CSMember As Long
    Dim Criteria As String
    Dim TestCriteria As String
    '
    ' Start the KeyCriteria
    '
    TestCriteria = "(" & BaseCriteria & ")"
    CSMember = Csv.OpenCSContent("people", TestCriteria, "ID")
    If Not Csv.IsCSOK(CSMember) Then
        '
        ' No users match the base criteria, exit
        '
        GetTargetMemberID = -1
    Else
        GetTargetMemberID = Csv.GetCSInteger(CSMember, "ID")
        Call Csv.NextCSRecord(CSMember)
        If Not Csv.IsCSOK(CSMember) Then
            '
            ' 1 match, exit with this match
            '
        Else
            '
            ' >1 match, go to next criteria
            '
            GetTargetMemberID = 0
            Criteria = Criteria & TestCriteria
        End If
    End If
    Call Csv.CloseCS(CSMember)
    '
    If GetTargetMemberID = 0 Then
        '
        ' Username and password Criteria
        '
        TestCriteria = "And(username is not null)And(password is not null)"
        CSMember = Csv.OpenCSContent("people", Criteria & TestCriteria)
        If Not Csv.IsCSOK(CSMember) Then
            '
            ' No matches, ignore this criteria and go on to the next
            '
            GetTargetMemberID = 0
        Else
            GetTargetMemberID = Csv.GetCSInteger(CSMember, "ID")
            Call Csv.NextCSRecord(CSMember)
            If Not Csv.IsCSOK(CSMember) Then
                '
                ' 1 match, exit with this match
                '
            Else
                '
                ' >1 match, go to next criteria
                '
                GetTargetMemberID = 0
                Criteria = Criteria & TestCriteria
            End If
        End If
        Call Csv.CloseCS(CSMember)
    End If
    '
    If GetTargetMemberID = 0 Then
        '
        ' Admin or Developer set
        '
        TestCriteria = "And((Admin<>0)or(Developer<>0))"
        CSMember = Csv.OpenCSContent("people", Criteria & TestCriteria)
        If Not Csv.IsCSOK(CSMember) Then
            '
            ' No matches, ignore this criteria and go on to the next
            '
            GetTargetMemberID = 0
        Else
            GetTargetMemberID = Csv.GetCSInteger(CSMember, "ID")
            Call Csv.NextCSRecord(CSMember)
            If Not Csv.IsCSOK(CSMember) Then
                '
                ' 1 match, exit with this match
                '
            Else
                '
                ' >1 match, go to next criteria
                '
                GetTargetMemberID = 0
                Criteria = Criteria & TestCriteria
            End If
        End If
        Call Csv.CloseCS(CSMember)
    End If
    '
    If GetTargetMemberID = 0 Then
        '
        ' Get the match with the last visit -- or the lowest ID
        '
        CSMember = Csv.OpenCSContent("people", Criteria, "LastVisit desc,ID")
        If Not Csv.IsCSOK(CSMember) Then
            '
            ' No matches, (very weird case)
            '
            GetTargetMemberID = 0
        Else
            GetTargetMemberID = Csv.GetCSInteger(CSMember, "ID")
        End If
        Call Csv.CloseCS(CSMember)
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetTargetMemberID")
    'Call HandleError("CleanClass", "isTarget", Err.Number, Err.Source, Err.Description, True, False)
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
    Call Err.Raise(ErrNumber, ErrSource, App.EXEName & ".CommonClass." & MethodName & " encountered a trap error: " & ErrDescription)
    '
End Sub

