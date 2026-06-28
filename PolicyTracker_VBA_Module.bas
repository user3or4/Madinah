Attribute VB_Name = "modPolicyTracker"
Option Compare Database
Option Explicit

' ============================================================
' Policy Tracker Management System - VBA Module
' ============================================================
' This module contains all VBA code for the system:
' - BusinessDaysFraction: Calculate working days excluding Fri/Sat
' - OLA Status calculation
' - Form event handlers
' - Query creation
' - Form creation
' ============================================================

' ============================================================
' FUNCTION: BusinessDaysFraction
' Calculates working days between two dates excluding Friday & Saturday
' 1 day = 8 hours, supports fractions (4 hours = 0.5 day)
' ============================================================
Public Function BusinessDaysFraction(StartDateTime As Date, EndDateTime As Date) As Double
    Dim totalMinutes As Double
    Dim currentDate As Date
    Dim startDate As Date
    Dim endDate As Date
    Dim startTime As Date
    Dim endTime As Date
    Dim workStartTime As Date
    Dim workEndTime As Date
    Dim dayMinutes As Double
    Dim totalWorkMinutes As Double

    If StartDateTime >= EndDateTime Then
        BusinessDaysFraction = 0
        Exit Function
    End If

    workStartTime = TimeSerial(8, 0, 0)  ' 8:00 AM
    workEndTime = TimeSerial(16, 0, 0)   ' 4:00 PM (8 hours)

    startDate = DateValue(StartDateTime)
    endDate = DateValue(EndDateTime)
    startTime = TimeValue(StartDateTime)
    endTime = TimeValue(EndDateTime)

    totalWorkMinutes = 0
    currentDate = startDate

    Do While currentDate <= endDate
        Dim dayOfWeek As Integer
        dayOfWeek = Weekday(currentDate, vbSunday)

        ' Skip Friday (6) and Saturday (7)
        If dayOfWeek <> 6 And dayOfWeek <> 7 Then
            Dim dayStart As Date
            Dim dayEnd As Date

            If currentDate = startDate And currentDate = endDate Then
                ' Same day
                dayStart = startTime
                If dayStart < workStartTime Then dayStart = workStartTime
                dayEnd = endTime
                If dayEnd > workEndTime Then dayEnd = workEndTime
                If dayEnd > dayStart Then
                    totalWorkMinutes = totalWorkMinutes + DateDiff("n", dayStart, dayEnd)
                End If
            ElseIf currentDate = startDate Then
                ' First day
                dayStart = startTime
                If dayStart < workStartTime Then dayStart = workStartTime
                If workEndTime > dayStart Then
                    totalWorkMinutes = totalWorkMinutes + DateDiff("n", dayStart, workEndTime)
                End If
            ElseIf currentDate = endDate Then
                ' Last day
                dayEnd = endTime
                If dayEnd > workEndTime Then dayEnd = workEndTime
                If dayEnd > workStartTime Then
                    totalWorkMinutes = totalWorkMinutes + DateDiff("n", workStartTime, dayEnd)
                End If
            Else
                ' Full working day
                totalWorkMinutes = totalWorkMinutes + 480 ' 8 hours = 480 minutes
            End If
        End If

        currentDate = currentDate + 1
    Loop

    ' Convert minutes to days (480 minutes = 1 day)
    BusinessDaysFraction = Round(totalWorkMinutes / 480, 2)
End Function

' ============================================================
' FUNCTION: GetOLAStatus
' Returns OLA status for a given item
' ============================================================
Public Function GetOLAStatus(itemName As String, startDate As Variant, isCompleted As Boolean) As String
    If IsNull(startDate) Or IsEmpty(startDate) Then
        GetOLAStatus = Chr(1604) & Chr(1605) & " " & Chr(1610) & Chr(1576) & Chr(1583) & Chr(1571)
        GetOLAStatus = "لم يبدأ"
        Exit Function
    End If

    If isCompleted Then
        GetOLAStatus = "مكتمل"
        Exit Function
    End If

    Dim olaDays As Double
    Dim warningPct As Double
    Dim rs As DAO.Recordset

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT OLA_Days, WarningPercentage FROM tblSLASetup WHERE ItemName='" & _
        Replace(itemName, "'", "''") & "'", dbOpenSnapshot)

    If rs.EOF Then
        GetOLAStatus = "غير معرف"
        rs.Close
        Set rs = Nothing
        Exit Function
    End If

    olaDays = rs!OLA_Days
    warningPct = rs!WarningPercentage
    rs.Close
    Set rs = Nothing

    Dim elapsed As Double
    elapsed = BusinessDaysFraction(CDate(startDate), Now())

    If elapsed > olaDays Then
        GetOLAStatus = "متجاوز OLA"
    ElseIf elapsed >= olaDays * warningPct Then
        GetOLAStatus = "قريب من التجاوز"
    Else
        GetOLAStatus = "ضمن OLA"
    End If
End Function

' ============================================================
' FUNCTION: GetStageSLAStatus
' Returns OLA status for current stage of a policy
' ============================================================
Public Function GetStageSLAStatus(policyID As String) As String
    Dim rs As DAO.Recordset
    Dim rsPolicy As DAO.Recordset

    Set rsPolicy = CurrentDb.OpenRecordset( _
        "SELECT CurrentStage, CurrentStageStart FROM tblPolicies WHERE PolicyID='" & _
        Replace(policyID, "'", "''") & "'", dbOpenSnapshot)

    If rsPolicy.EOF Then
        GetStageSLAStatus = "غير موجود"
        rsPolicy.Close
        Set rsPolicy = Nothing
        Exit Function
    End If

    Dim currentStage As String
    currentStage = Nz(rsPolicy!CurrentStage, "")

    ' Check if closed stage
    If currentStage = "معتمدة" Or currentStage = "ملغي / مرفوض" Then
        GetStageSLAStatus = "مكتمل"
        rsPolicy.Close
        Set rsPolicy = Nothing
        Exit Function
    End If

    GetStageSLAStatus = GetOLAStatus(currentStage, rsPolicy!CurrentStageStart, False)

    rsPolicy.Close
    Set rsPolicy = Nothing
End Function

' ============================================================
' SUB: AddNewPolicy
' Adds a new policy and creates initial workflow log entry
' ============================================================
Public Sub AddNewPolicy(policyID As String, policyName As String, _
    sector As String, ownerDept As String, responsiblePerson As String, _
    notes As String)

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsLog As DAO.Recordset
    Dim nowDT As Date

    nowDT = Now()
    Set db = CurrentDb

    ' Check if PolicyID already exists
    Set rs = db.OpenRecordset("SELECT PolicyID FROM tblPolicies WHERE PolicyID='" & _
        Replace(policyID, "'", "''") & "'", dbOpenSnapshot)
    If Not rs.EOF Then
        MsgBox "رقم السياسة موجود مسبقاً: " & policyID, vbExclamation, "خطأ"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close

    ' Add to tblPolicies
    Set rs = db.OpenRecordset("tblPolicies", dbOpenDynaset)
    rs.AddNew
    rs!policyID = policyID
    rs!policyName = policyName
    rs!sector = sector
    rs!OwnerDepartment = ownerDept
    rs!ResponsiblePerson = responsiblePerson
    rs!CurrentStage = "مرحلة التسجيل"
    rs!ReceivedFrom = Null
    rs!CreatedDate = nowDT
    rs!CurrentStageStart = nowDT
    rs!LastUpdate = nowDT
    rs!CurrentSLAStatus = "ضمن OLA"
    rs!notes = notes
    rs.Update
    rs.Close

    ' Add initial workflow log
    Set rsLog = db.OpenRecordset("tblWorkflowLog", dbOpenDynaset)
    rsLog.AddNew
    rsLog!policyID = policyID
    rsLog!policyName = policyName
    rsLog!sector = sector
    rsLog!OwnerDepartment = ownerDept
    rsLog!PreviousStage = Null
    rsLog!NewStage = "مرحلة التسجيل"
    rsLog!ReceivedFromDepartment = Null
    rsLog!EntryDateTime = nowDT
    rsLog!ExitDateTime = Null
    rsLog!DurationDays = Null
    rsLog!RoundNo = 1
    rsLog!StageNotes = Null
    rsLog!UpdateDate = DateValue(nowDT)
    rsLog!UpdateTime = TimeValue(nowDT)
    rsLog.Update
    rsLog.Close

    Set rs = Nothing
    Set rsLog = Nothing
    Set db = Nothing

    MsgBox "تم إضافة السياسة بنجاح: " & policyID, vbInformation, "نجاح"
End Sub

' ============================================================
' SUB: UpdatePolicyStage
' Changes the stage of a policy and logs the transition
' ============================================================
Public Sub UpdatePolicyStage(policyID As String, newStage As String, stageNotes As String)
    Dim db As DAO.Database
    Dim rsPolicy As DAO.Recordset
    Dim rsLog As DAO.Recordset
    Dim nowDT As Date
    Dim previousStage As String
    Dim policyName As String
    Dim sector As String
    Dim ownerDept As String
    Dim roundNo As Long

    nowDT = Now()
    Set db = CurrentDb

    ' Get current policy info
    Set rsPolicy = db.OpenRecordset( _
        "SELECT * FROM tblPolicies WHERE PolicyID='" & _
        Replace(policyID, "'", "''") & "'", dbOpenDynaset)

    If rsPolicy.EOF Then
        MsgBox "السياسة غير موجودة: " & policyID, vbExclamation, "خطأ"
        rsPolicy.Close
        Exit Sub
    End If

    previousStage = Nz(rsPolicy!CurrentStage, "")
    policyName = Nz(rsPolicy!policyName, "")
    sector = Nz(rsPolicy!sector, "")
    ownerDept = Nz(rsPolicy!OwnerDepartment, "")

    If previousStage = newStage Then
        MsgBox "المرحلة الجديدة هي نفس المرحلة الحالية!", vbExclamation, "تنبيه"
        rsPolicy.Close
        Exit Sub
    End If

    ' Calculate round number
    Dim rsRound As DAO.Recordset
    Set rsRound = db.OpenRecordset( _
        "SELECT MAX(RoundNo) AS MaxRound FROM tblWorkflowLog WHERE PolicyID='" & _
        Replace(policyID, "'", "''") & "' AND NewStage IS NOT NULL", dbOpenSnapshot)
    If rsRound.EOF Or IsNull(rsRound!MaxRound) Then
        roundNo = 1
    Else
        roundNo = rsRound!MaxRound
        ' If going backwards, increment round
        Dim prevOrder As Integer, newOrder As Integer
        prevOrder = GetStageOrder(previousStage)
        newOrder = GetStageOrder(newStage)
        If newOrder <= prevOrder And newStage <> "ملغي / مرفوض" Then
            roundNo = roundNo + 1
        End If
    End If
    rsRound.Close

    ' Close the last open workflow log entry for this policy (stage changes only)
    Set rsLog = db.OpenRecordset( _
        "SELECT * FROM tblWorkflowLog WHERE PolicyID='" & _
        Replace(policyID, "'", "''") & _
        "' AND ExitDateTime IS NULL AND NewStage IS NOT NULL ORDER BY LogID DESC", dbOpenDynaset)

    If Not rsLog.EOF Then
        rsLog.Edit
        rsLog!ExitDateTime = nowDT
        rsLog!DurationDays = BusinessDaysFraction(rsLog!EntryDateTime, nowDT)
        rsLog.Update
    End If
    rsLog.Close

    ' Add new workflow log entry
    Set rsLog = db.OpenRecordset("tblWorkflowLog", dbOpenDynaset)
    rsLog.AddNew
    rsLog!policyID = policyID
    rsLog!policyName = policyName
    rsLog!sector = sector
    rsLog!OwnerDepartment = ownerDept
    rsLog!PreviousStage = previousStage
    rsLog!NewStage = newStage
    rsLog!ReceivedFromDepartment = Null
    rsLog!EntryDateTime = nowDT
    rsLog!ExitDateTime = Null
    rsLog!DurationDays = Null
    rsLog!RoundNo = roundNo
    rsLog!StageNotes = stageNotes
    rsLog!UpdateDate = DateValue(nowDT)
    rsLog!UpdateTime = TimeValue(nowDT)
    rsLog.Update
    rsLog.Close

    ' Update tblPolicies
    rsPolicy.Edit
    rsPolicy!CurrentStage = newStage
    rsPolicy!CurrentStageStart = nowDT
    rsPolicy!LastUpdate = nowDT
    rsPolicy!ReceivedFrom = Null  ' Clear ReceivedFrom on stage change
    rsPolicy!CurrentSLAStatus = GetOLAStatus(newStage, nowDT, _
        (newStage = "معتمدة" Or newStage = "ملغي / مرفوض"))
    rsPolicy.Update
    rsPolicy.Close

    Set rsPolicy = Nothing
    Set rsLog = Nothing
    Set rsRound = Nothing
    Set db = Nothing

    MsgBox "تم تحديث المرحلة بنجاح" & vbCrLf & _
        "من: " & previousStage & vbCrLf & _
        "إلى: " & newStage, vbInformation, "نجاح"
End Sub

' ============================================================
' SUB: RegisterReceivedFrom
' Records a "received from" event without changing the stage
' ============================================================
Public Sub RegisterReceivedFrom(policyID As String, receivedFromDept As String, stageNotes As String)
    Dim db As DAO.Database
    Dim rsPolicy As DAO.Recordset
    Dim rsLog As DAO.Recordset
    Dim nowDT As Date

    nowDT = Now()
    Set db = CurrentDb

    ' Get current policy info
    Set rsPolicy = db.OpenRecordset( _
        "SELECT * FROM tblPolicies WHERE PolicyID='" & _
        Replace(policyID, "'", "''") & "'", dbOpenDynaset)

    If rsPolicy.EOF Then
        MsgBox "السياسة غير موجودة: " & policyID, vbExclamation, "خطأ"
        rsPolicy.Close
        Exit Sub
    End If

    Dim currentStage As String
    Dim policyName As String
    Dim sector As String
    Dim ownerDept As String

    currentStage = Nz(rsPolicy!CurrentStage, "")
    policyName = Nz(rsPolicy!policyName, "")
    sector = Nz(rsPolicy!sector, "")
    ownerDept = Nz(rsPolicy!OwnerDepartment, "")

    ' Add workflow log entry (received from event - no stage change)
    Set rsLog = db.OpenRecordset("tblWorkflowLog", dbOpenDynaset)
    rsLog.AddNew
    rsLog!policyID = policyID
    rsLog!policyName = policyName
    rsLog!sector = sector
    rsLog!OwnerDepartment = ownerDept
    rsLog!PreviousStage = currentStage
    rsLog!NewStage = Null  ' No stage change
    rsLog!ReceivedFromDepartment = receivedFromDept
    rsLog!EntryDateTime = nowDT
    rsLog!ExitDateTime = Null
    rsLog!DurationDays = Null
    rsLog!RoundNo = Null
    rsLog!StageNotes = stageNotes
    rsLog!UpdateDate = DateValue(nowDT)
    rsLog!UpdateTime = TimeValue(nowDT)
    rsLog.Update
    rsLog.Close

    ' Update tblPolicies - only LastUpdate, clear ReceivedFrom
    rsPolicy.Edit
    rsPolicy!LastUpdate = nowDT
    rsPolicy!ReceivedFrom = Null
    rsPolicy.Update
    rsPolicy.Close

    Set rsPolicy = Nothing
    Set rsLog = Nothing
    Set db = Nothing

    MsgBox "تم تسجيل الاستلام من: " & receivedFromDept, vbInformation, "نجاح"
End Sub

' ============================================================
' FUNCTION: GetStageOrder
' Returns the order number for a stage name
' ============================================================
Public Function GetStageOrder(stageName As String) As Integer
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT StageOrder FROM tblStages WHERE StageName='" & _
        Replace(stageName, "'", "''") & "'", dbOpenSnapshot)

    If rs.EOF Then
        GetStageOrder = 0
    Else
        GetStageOrder = rs!StageOrder
    End If
    rs.Close
    Set rs = Nothing
End Function

' ============================================================
' FUNCTION: GetReceivedFromOptions
' Returns departments available for "Received From" based on current stage
' ============================================================
Public Function GetReceivedFromOptions(currentStage As String) As String
    Select Case currentStage
        Case "مرحلة المراجعة الفنية"
            GetReceivedFromOptions = "'الإدارة الفنية','الإدارة المالكة','إدارة الحوكمة والامتثال','الإدارة العامة للقانونية والتشريعات'"
        Case "مرحلة المراجعة العمومية"
            GetReceivedFromOptions = "'الإدارة المالكة','الإدارة العامة للتميز','الإدارة العامة للقانونية والتشريعات'"
        Case "مرحلة الإعتمادات"
            GetReceivedFromOptions = "'الإدارة المالكة','الإدارة العامة للتميز','الإدارة العامة للقانونية والتشريعات','اعتماد معالي المحافظ'"
        Case Else
            GetReceivedFromOptions = ""
    End Select
End Function

' ============================================================
' FUNCTION: GetDashboardValue
' Returns dashboard metric values
' ============================================================
Public Function GetDashboardValue(metricName As String) As Variant
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = CurrentDb

    Select Case metricName
        Case "Total"
            sql = "SELECT COUNT(*) AS Cnt FROM tblPolicies"
        Case "Registration"
            sql = "SELECT COUNT(*) AS Cnt FROM tblPolicies WHERE CurrentStage='مرحلة التسجيل'"
        Case "TechnicalReview"
            sql = "SELECT COUNT(*) AS Cnt FROM tblPolicies WHERE CurrentStage='مرحلة المراجعة الفنية'"
        Case "GeneralReview"
            sql = "SELECT COUNT(*) AS Cnt FROM tblPolicies WHERE CurrentStage='مرحلة المراجعة العمومية'"
        Case "Approvals"
            sql = "SELECT COUNT(*) AS Cnt FROM tblPolicies WHERE CurrentStage='مرحلة الإعتمادات'"
        Case "Approved"
            sql = "SELECT COUNT(*) AS Cnt FROM tblPolicies WHERE CurrentStage='معتمدة'"
        Case "Cancelled"
            sql = "SELECT COUNT(*) AS Cnt FROM tblPolicies WHERE CurrentStage='ملغي / مرفوض'"
        Case "OverdueOLA"
            sql = "SELECT COUNT(*) AS Cnt FROM tblPolicies WHERE CurrentSLAStatus='متجاوز OLA'"
        Case "NearOverdue"
            sql = "SELECT COUNT(*) AS Cnt FROM tblPolicies WHERE CurrentSLAStatus='قريب من التجاوز'"
        Case Else
            GetDashboardValue = 0
            Exit Function
    End Select

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then
        GetDashboardValue = rs!Cnt
    Else
        GetDashboardValue = 0
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' ============================================================
' FUNCTION: GetAverageDurationByStage
' Returns average duration for a specific stage
' ============================================================
Public Function GetAverageDurationByStage(stageName As String) As Double
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT AVG(DurationDays) AS AvgDays FROM tblWorkflowLog " & _
        "WHERE NewStage='" & Replace(stageName, "'", "''") & _
        "' AND DurationDays IS NOT NULL AND NewStage IS NOT NULL", dbOpenSnapshot)

    If Not rs.EOF And Not IsNull(rs!AvgDays) Then
        GetAverageDurationByStage = Round(rs!AvgDays, 2)
    Else
        GetAverageDurationByStage = 0
    End If
    rs.Close
    Set rs = Nothing
End Function

' ============================================================
' FUNCTION: GetTopDelayDepartment
' Returns the department causing most delays
' ============================================================
Public Function GetTopDelayDepartment() As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT TOP 1 ReceivedFromDepartment, COUNT(*) AS Cnt " & _
        "FROM tblWorkflowLog " & _
        "WHERE ReceivedFromDepartment IS NOT NULL " & _
        "GROUP BY ReceivedFromDepartment " & _
        "ORDER BY COUNT(*) DESC", dbOpenSnapshot)

    If Not rs.EOF Then
        GetTopDelayDepartment = rs!ReceivedFromDepartment & " (" & rs!Cnt & ")"
    Else
        GetTopDelayDepartment = "لا يوجد"
    End If
    rs.Close
    Set rs = Nothing
End Function

' ============================================================
' SUB: CreateAllQueries
' Creates all required queries in the database
' ============================================================
Public Sub CreateAllQueries()
    Dim db As DAO.Database
    Set db = CurrentDb

    On Error Resume Next

    ' Delete existing queries first
    db.QueryDefs.Delete "qryOpenPolicies"
    db.QueryDefs.Delete "qryDashboardByStage"
    db.QueryDefs.Delete "qryWorkflowLogOrdered"
    db.QueryDefs.Delete "qryReceivedFromSummary"
    db.QueryDefs.Delete "qryAverageDurationByStage"
    db.QueryDefs.Delete "qryOverdueOLA"
    db.QueryDefs.Delete "qryAllPolicies"
    db.QueryDefs.Delete "qryStageList"
    db.QueryDefs.Delete "qryDepartmentList"

    On Error GoTo 0

    Dim qd As DAO.QueryDef

    ' qryOpenPolicies
    Set qd = db.CreateQueryDef("qryOpenPolicies", _
        "SELECT PolicyID, PolicyName, Sector, OwnerDepartment, ResponsiblePerson, " & _
        "CurrentStage, ReceivedFrom, CreatedDate, CurrentStageStart, LastUpdate, " & _
        "CurrentSLAStatus, Notes " & _
        "FROM tblPolicies " & _
        "WHERE CurrentStage <> 'معتمدة' AND CurrentStage <> 'ملغي / مرفوض' " & _
        "ORDER BY LastUpdate DESC;")

    ' qryDashboardByStage
    Set qd = db.CreateQueryDef("qryDashboardByStage", _
        "SELECT CurrentStage, COUNT(*) AS PolicyCount " & _
        "FROM tblPolicies " & _
        "GROUP BY CurrentStage " & _
        "ORDER BY CurrentStage;")

    ' qryWorkflowLogOrdered
    Set qd = db.CreateQueryDef("qryWorkflowLogOrdered", _
        "SELECT LogID, PolicyID, PolicyName, Sector, OwnerDepartment, " & _
        "PreviousStage, NewStage, ReceivedFromDepartment, " & _
        "EntryDateTime, ExitDateTime, DurationDays, RoundNo, " & _
        "StageNotes, UpdateDate, UpdateTime " & _
        "FROM tblWorkflowLog " & _
        "ORDER BY PolicyID, LogID;")

    ' qryReceivedFromSummary
    Set qd = db.CreateQueryDef("qryReceivedFromSummary", _
        "SELECT ReceivedFromDepartment, COUNT(*) AS ReceiveCount " & _
        "FROM tblWorkflowLog " & _
        "WHERE ReceivedFromDepartment IS NOT NULL " & _
        "GROUP BY ReceivedFromDepartment " & _
        "ORDER BY COUNT(*) DESC;")

    ' qryAverageDurationByStage
    Set qd = db.CreateQueryDef("qryAverageDurationByStage", _
        "SELECT NewStage AS StageName, " & _
        "AVG(DurationDays) AS AvgDurationDays, " & _
        "MIN(DurationDays) AS MinDurationDays, " & _
        "MAX(DurationDays) AS MaxDurationDays, " & _
        "COUNT(*) AS TransitionCount " & _
        "FROM tblWorkflowLog " & _
        "WHERE DurationDays IS NOT NULL AND NewStage IS NOT NULL " & _
        "GROUP BY NewStage;")

    ' qryOverdueOLA
    Set qd = db.CreateQueryDef("qryOverdueOLA", _
        "SELECT PolicyID, PolicyName, CurrentStage, CurrentStageStart, " & _
        "CurrentSLAStatus, LastUpdate " & _
        "FROM tblPolicies " & _
        "WHERE CurrentSLAStatus IN ('متجاوز OLA','قريب من التجاوز') " & _
        "ORDER BY CurrentStageStart;")

    ' qryAllPolicies - for dropdowns
    Set qd = db.CreateQueryDef("qryAllPolicies", _
        "SELECT PolicyID, PolicyName, PolicyID & ' - ' & PolicyName AS DisplayText " & _
        "FROM tblPolicies ORDER BY PolicyID;")

    ' qryStageList - for dropdowns
    Set qd = db.CreateQueryDef("qryStageList", _
        "SELECT StageName FROM tblStages ORDER BY StageOrder;")

    ' qryDepartmentList
    Set qd = db.CreateQueryDef("qryDepartmentList", _
        "SELECT DepartmentName FROM tblDepartments ORDER BY DepartmentID;")

    Set qd = Nothing
    Set db = Nothing

    MsgBox "تم إنشاء جميع الاستعلامات بنجاح!", vbInformation, "نجاح"
End Sub

' ============================================================
' SUB: RefreshAllSLAStatuses
' Updates SLA status for all open policies
' ============================================================
Public Sub RefreshAllSLAStatuses()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblPolicies WHERE CurrentStage <> 'معتمدة' " & _
        "AND CurrentStage <> 'ملغي / مرفوض'", dbOpenDynaset)

    Do While Not rs.EOF
        Dim newStatus As String
        newStatus = GetOLAStatus(rs!CurrentStage, rs!CurrentStageStart, False)

        If Nz(rs!CurrentSLAStatus, "") <> newStatus Then
            rs.Edit
            rs!CurrentSLAStatus = newStatus
            rs.Update
        End If

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

' ============================================================
' SUB: SetupSystem
' Initial setup - creates queries and refreshes SLA
' ============================================================
Public Sub SetupSystem()
    CreateAllQueries
    RefreshAllSLAStatuses
    MsgBox "تم إعداد النظام بنجاح!" & vbCrLf & _
        "يمكنك الآن استخدام النماذج للعمل.", vbInformation, "إعداد النظام"
End Sub
