Attribute VB_Name = "modOLACalculation"
Option Compare Database
Option Explicit

' ============================================================
' OLA Calculation Module for Technical Review Stage
' ============================================================
' Handles the complex OLA calculation logic for departments
' within the Technical Review stage (and other stages)
' ============================================================

' ============================================================
' FUNCTION: GetDepartmentOLAStatus
' Calculates OLA status for a specific department within a stage
' for a given policy, considering all rounds cumulatively
' ============================================================
Public Function GetDepartmentOLAStatus(policyID As String, departmentName As String) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsPolicy As DAO.Recordset
    Dim startDate As Variant
    Dim endDate As Variant
    Dim totalDays As Double
    Dim olaDays As Double
    Dim warningPct As Double
    Dim isCompleted As Boolean

    Set db = CurrentDb

    ' Get OLA settings for this department
    Dim rsSLA As DAO.Recordset
    Set rsSLA = db.OpenRecordset( _
        "SELECT OLA_Days, WarningPercentage FROM tblSLASetup WHERE ItemName='" & _
        Replace(departmentName, "'", "''") & "' AND ItemType='Department'", dbOpenSnapshot)

    If rsSLA.EOF Then
        GetDepartmentOLAStatus = "غير معرف"
        rsSLA.Close
        Exit Function
    End If

    olaDays = rsSLA!OLA_Days
    warningPct = rsSLA!WarningPercentage
    rsSLA.Close

    ' Calculate cumulative time for this department across all rounds
    totalDays = CalculateCumulativeDepartmentDays(policyID, departmentName)
    isCompleted = IsDepartmentCompleted(policyID, departmentName)

    If totalDays = -1 Then
        GetDepartmentOLAStatus = "لم يبدأ"
        Exit Function
    End If

    If isCompleted Then
        GetDepartmentOLAStatus = "مكتمل"
        Exit Function
    End If

    If totalDays > olaDays Then
        GetDepartmentOLAStatus = "متجاوز OLA"
    ElseIf totalDays >= olaDays * warningPct Then
        GetDepartmentOLAStatus = "قريب من التجاوز"
    Else
        GetDepartmentOLAStatus = "ضمن OLA"
    End If

    Set db = Nothing
End Function

' ============================================================
' FUNCTION: CalculateCumulativeDepartmentDays
' Calculates total business days a department has been active
' across all rounds for a policy
' Returns -1 if department hasn't started
' ============================================================
Public Function CalculateCumulativeDepartmentDays(policyID As String, departmentName As String) As Double
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim totalDays As Double
    Dim hasStarted As Boolean

    Set db = CurrentDb
    totalDays = 0
    hasStarted = False

    ' Get the current stage of the policy
    Dim rsPolicy As DAO.Recordset
    Set rsPolicy = db.OpenRecordset( _
        "SELECT CurrentStage, CurrentStageStart FROM tblPolicies WHERE PolicyID='" & _
        Replace(policyID, "'", "''") & "'", dbOpenSnapshot)

    If rsPolicy.EOF Then
        CalculateCumulativeDepartmentDays = -1
        rsPolicy.Close
        Exit Function
    End If

    Dim currentStage As String
    currentStage = Nz(rsPolicy!CurrentStage, "")
    rsPolicy.Close

    ' Get all workflow events for this policy ordered chronologically
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblWorkflowLog WHERE PolicyID='" & _
        Replace(policyID, "'", "''") & _
        "' ORDER BY LogID", dbOpenSnapshot)

    ' Track department activation periods
    Dim deptStartDate As Variant
    deptStartDate = Null

    Do While Not rs.EOF
        Dim newStage As String
        Dim receivedFrom As String

        newStage = Nz(rs!NewStage, "")
        receivedFrom = Nz(rs!ReceivedFromDepartment, "")

        ' Handle Technical Review stage logic
        If currentStage = "مرحلة المراجعة الفنية" Or _
           Nz(rs!PreviousStage, "") = "مرحلة المراجعة الفنية" Then

            Call ProcessTechnicalReviewOLA(departmentName, newStage, receivedFrom, _
                rs!EntryDateTime, deptStartDate, totalDays, hasStarted)
        End If

        ' Handle General Review stage logic
        If currentStage = "مرحلة المراجعة العمومية" Or _
           Nz(rs!PreviousStage, "") = "مرحلة المراجعة العمومية" Then

            Call ProcessGeneralReviewOLA(departmentName, newStage, receivedFrom, _
                rs!EntryDateTime, deptStartDate, totalDays, hasStarted)
        End If

        ' Handle Approvals stage logic
        If currentStage = "مرحلة الإعتمادات" Or _
           Nz(rs!PreviousStage, "") = "مرحلة الإعتمادات" Then

            Call ProcessApprovalsOLA(departmentName, newStage, receivedFrom, _
                rs!EntryDateTime, deptStartDate, totalDays, hasStarted)
        End If

        rs.MoveNext
    Loop

    rs.Close

    ' If department is still active (no end date), add time until now
    If Not IsNull(deptStartDate) Then
        totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), Now())
        hasStarted = True
    End If

    If Not hasStarted Then
        CalculateCumulativeDepartmentDays = -1
    Else
        CalculateCumulativeDepartmentDays = totalDays
    End If

    Set rs = Nothing
    Set db = Nothing
End Function

' ============================================================
' SUB: ProcessTechnicalReviewOLA
' Handles OLA logic specific to Technical Review stage
' ============================================================
Private Sub ProcessTechnicalReviewOLA(departmentName As String, _
    newStage As String, receivedFrom As String, _
    eventDateTime As Date, ByRef deptStartDate As Variant, _
    ByRef totalDays As Double, ByRef hasStarted As Boolean)

    Select Case departmentName
        Case "الإدارة الفنية"
            ' Technical dept starts when entering Technical Review stage
            If newStage = "مرحلة المراجعة الفنية" Then
                deptStartDate = eventDateTime
                hasStarted = True
            End If
            ' Ends when "received from Technical dept" is logged
            If receivedFrom = "الإدارة الفنية" And Not IsNull(deptStartDate) Then
                totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), eventDateTime)
                deptStartDate = Null
            End If
            ' Restarts after both Governance and Legal respond
            If receivedFrom = "إدارة الحوكمة والامتثال" Or _
               receivedFrom = "الإدارة العامة للقانونية والتشريعات" Then
                ' Check if both have responded - if so, Technical restarts
                If BothGovernanceAndLegalResponded(receivedFrom, eventDateTime) Then
                    deptStartDate = eventDateTime
                    hasStarted = True
                End If
            End If

        Case "الإدارة المالكة"
            ' Owner dept starts after Technical dept responds
            If receivedFrom = "الإدارة الفنية" Then
                deptStartDate = eventDateTime
                hasStarted = True
            End If
            ' Ends when "received from Owner dept" is logged
            If receivedFrom = "الإدارة المالكة" And Not IsNull(deptStartDate) Then
                totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), eventDateTime)
                deptStartDate = Null
            End If

        Case "إدارة الحوكمة والامتثال"
            ' Governance starts after Owner dept responds
            If receivedFrom = "الإدارة المالكة" Then
                deptStartDate = eventDateTime
                hasStarted = True
            End If
            ' Ends when "received from Governance" is logged
            If receivedFrom = "إدارة الحوكمة والامتثال" And Not IsNull(deptStartDate) Then
                totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), eventDateTime)
                deptStartDate = Null
            End If

        Case "الإدارة العامة للقانونية والتشريعات"
            ' Legal starts after Owner dept responds (same time as Governance)
            If receivedFrom = "الإدارة المالكة" Then
                deptStartDate = eventDateTime
                hasStarted = True
            End If
            ' Ends when "received from Legal" is logged
            If receivedFrom = "الإدارة العامة للقانونية والتشريعات" And Not IsNull(deptStartDate) Then
                totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), eventDateTime)
                deptStartDate = Null
            End If
    End Select
End Sub

' ============================================================
' SUB: ProcessGeneralReviewOLA
' Handles OLA logic for General Review stage
' ============================================================
Private Sub ProcessGeneralReviewOLA(departmentName As String, _
    newStage As String, receivedFrom As String, _
    eventDateTime As Date, ByRef deptStartDate As Variant, _
    ByRef totalDays As Double, ByRef hasStarted As Boolean)

    Select Case departmentName
        Case "الإدارة المالكة"
            If newStage = "مرحلة المراجعة العمومية" Then
                deptStartDate = eventDateTime
                hasStarted = True
            End If
            If receivedFrom = "الإدارة المالكة" And Not IsNull(deptStartDate) Then
                totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), eventDateTime)
                deptStartDate = Null
            End If

        Case "الإدارة العامة للتميز"
            If newStage = "مرحلة المراجعة العمومية" Then
                deptStartDate = eventDateTime
                hasStarted = True
            End If
            If receivedFrom = "الإدارة العامة للتميز" And Not IsNull(deptStartDate) Then
                totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), eventDateTime)
                deptStartDate = Null
            End If

        Case "الإدارة العامة للقانونية والتشريعات"
            If newStage = "مرحلة المراجعة العمومية" Then
                deptStartDate = eventDateTime
                hasStarted = True
            End If
            If receivedFrom = "الإدارة العامة للقانونية والتشريعات" And Not IsNull(deptStartDate) Then
                totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), eventDateTime)
                deptStartDate = Null
            End If
    End Select
End Sub

' ============================================================
' SUB: ProcessApprovalsOLA
' Handles OLA logic for Approvals stage
' ============================================================
Private Sub ProcessApprovalsOLA(departmentName As String, _
    newStage As String, receivedFrom As String, _
    eventDateTime As Date, ByRef deptStartDate As Variant, _
    ByRef totalDays As Double, ByRef hasStarted As Boolean)

    Select Case departmentName
        Case "الإدارة المالكة"
            If newStage = "مرحلة الإعتمادات" Then
                deptStartDate = eventDateTime
                hasStarted = True
            End If
            If receivedFrom = "الإدارة المالكة" And Not IsNull(deptStartDate) Then
                totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), eventDateTime)
                deptStartDate = Null
            End If

        Case "الإدارة العامة للتميز"
            If newStage = "مرحلة الإعتمادات" Then
                deptStartDate = eventDateTime
                hasStarted = True
            End If
            If receivedFrom = "الإدارة العامة للتميز" And Not IsNull(deptStartDate) Then
                totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), eventDateTime)
                deptStartDate = Null
            End If

        Case "الإدارة العامة للقانونية والتشريعات"
            If newStage = "مرحلة الإعتمادات" Then
                deptStartDate = eventDateTime
                hasStarted = True
            End If
            If receivedFrom = "الإدارة العامة للقانونية والتشريعات" And Not IsNull(deptStartDate) Then
                totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), eventDateTime)
                deptStartDate = Null
            End If

        Case "اعتماد معالي المحافظ"
            If newStage = "مرحلة الإعتمادات" Then
                deptStartDate = eventDateTime
                hasStarted = True
            End If
            If receivedFrom = "اعتماد معالي المحافظ" And Not IsNull(deptStartDate) Then
                totalDays = totalDays + BusinessDaysFraction(CDate(deptStartDate), eventDateTime)
                deptStartDate = Null
            End If
    End Select
End Sub

' ============================================================
' FUNCTION: BothGovernanceAndLegalResponded
' Checks if both Governance and Legal have responded
' (simplified - checks last received entries)
' ============================================================
Private Function BothGovernanceAndLegalResponded(latestDept As String, eventDateTime As Date) As Boolean
    ' This is a simplified check - in production, you'd track
    ' the specific round context
    BothGovernanceAndLegalResponded = False

    ' If current event is Governance, check if Legal already responded
    ' If current event is Legal, check if Governance already responded
    Dim otherDept As String
    If latestDept = "إدارة الحوكمة والامتثال" Then
        otherDept = "الإدارة العامة للقانونية والتشريعات"
    ElseIf latestDept = "الإدارة العامة للقانونية والتشريعات" Then
        otherDept = "إدارة الحوكمة والامتثال"
    Else
        Exit Function
    End If

    ' Check if the other department has a "received from" entry
    ' after the last "received from Owner" entry
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 EntryDateTime FROM tblWorkflowLog " & _
        "WHERE ReceivedFromDepartment='" & Replace(otherDept, "'", "''") & "' " & _
        "AND EntryDateTime <= #" & Format(eventDateTime, "yyyy-mm-dd hh:nn:ss") & "# " & _
        "ORDER BY LogID DESC", dbOpenSnapshot)

    If Not rs.EOF Then
        BothGovernanceAndLegalResponded = True
    End If

    rs.Close
    Set rs = Nothing
End Function

' ============================================================
' FUNCTION: IsDepartmentCompleted
' Checks if a department has completed its current task
' ============================================================
Public Function IsDepartmentCompleted(policyID As String, departmentName As String) As Boolean
    Dim rs As DAO.Recordset

    ' Check if there's a "received from" entry for this department
    ' that is the latest action for this department
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 ReceivedFromDepartment FROM tblWorkflowLog " & _
        "WHERE PolicyID='" & Replace(policyID, "'", "''") & "' " & _
        "AND ReceivedFromDepartment='" & Replace(departmentName, "'", "''") & "' " & _
        "ORDER BY LogID DESC", dbOpenSnapshot)

    If Not rs.EOF Then
        ' Department has responded at least once
        ' Check if it was re-activated after that
        ' (simplified check)
        IsDepartmentCompleted = True
    Else
        IsDepartmentCompleted = False
    End If

    rs.Close
    Set rs = Nothing
End Function
