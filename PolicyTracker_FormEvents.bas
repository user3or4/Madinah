Attribute VB_Name = "modFormEvents"
Option Compare Database
Option Explicit

' ============================================================
' Form Event Handlers for Policy Tracker Management System
' ============================================================
' These procedures should be called from Form events
' Copy the relevant code to each form's code module
' ============================================================

' ============================================================
' frmMainMenu - Button Click Events
' ============================================================
' Place this code in frmMainMenu's Form module:
'
' Private Sub btnAddPolicy_Click()
'     DoCmd.OpenForm "frmAddPolicy", acNormal
' End Sub
'
' Private Sub btnUpdatePolicy_Click()
'     DoCmd.OpenForm "frmUpdatePolicy", acNormal
' End Sub
'
' Private Sub btnViewPolicies_Click()
'     DoCmd.OpenForm "frmViewPolicies", acNormal
' End Sub
'
' Private Sub btnWorkflowLog_Click()
'     DoCmd.OpenForm "frmWorkflowLog", acNormal
' End Sub
'
' Private Sub btnSLASetup_Click()
'     DoCmd.OpenForm "frmSLASetup", acNormal
' End Sub
'
' Private Sub btnDashboard_Click()
'     DoCmd.OpenForm "frmDashboard", acNormal
' End Sub
'
' Private Sub btnExit_Click()
'     DoCmd.Quit
' End Sub
'
' Private Sub Form_Open(Cancel As Integer)
'     RefreshAllSLAStatuses
' End Sub

' ============================================================
' frmAddPolicy - Save Button
' ============================================================
' Place this code in frmAddPolicy's Form module:
'
' Private Sub btnSave_Click()
'     ' Validate required fields
'     If Nz(Me.txtPolicyID, "") = "" Then
'         MsgBox "يرجى إدخال رقم السياسة", vbExclamation, "تنبيه"
'         Me.txtPolicyID.SetFocus
'         Exit Sub
'     End If
'
'     If Nz(Me.txtPolicyName, "") = "" Then
'         MsgBox "يرجى إدخال اسم السياسة", vbExclamation, "تنبيه"
'         Me.txtPolicyName.SetFocus
'         Exit Sub
'     End If
'
'     If Nz(Me.cboSector, "") = "" Then
'         MsgBox "يرجى اختيار القطاع", vbExclamation, "تنبيه"
'         Me.cboSector.SetFocus
'         Exit Sub
'     End If
'
'     If Nz(Me.cboOwnerDepartment, "") = "" Then
'         MsgBox "يرجى اختيار الإدارة المالكة", vbExclamation, "تنبيه"
'         Me.cboOwnerDepartment.SetFocus
'         Exit Sub
'     End If
'
'     If Nz(Me.txtResponsiblePerson, "") = "" Then
'         MsgBox "يرجى إدخال اسم المسؤول", vbExclamation, "تنبيه"
'         Me.txtResponsiblePerson.SetFocus
'         Exit Sub
'     End If
'
'     AddNewPolicy Me.txtPolicyID, Me.txtPolicyName, _
'         Me.cboSector, Me.cboOwnerDepartment, _
'         Me.txtResponsiblePerson, Nz(Me.txtNotes, "")
'
'     ' Clear fields after save
'     Me.txtPolicyID = ""
'     Me.txtPolicyName = ""
'     Me.cboSector = Null
'     Me.cboOwnerDepartment = Null
'     Me.txtResponsiblePerson = ""
'     Me.txtNotes = ""
'     Me.txtPolicyID.SetFocus
' End Sub
'
' Private Sub btnClose_Click()
'     DoCmd.Close acForm, Me.Name
' End Sub

' ============================================================
' frmUpdatePolicy - Event Handlers
' ============================================================
' Place this code in frmUpdatePolicy's Form module:
'
' Private Sub cboPolicy_AfterUpdate()
'     If IsNull(Me.cboPolicy) Then Exit Sub
'
'     Dim rs As DAO.Recordset
'     Set rs = CurrentDb.OpenRecordset( _
'         "SELECT * FROM tblPolicies WHERE PolicyID='" & _
'         Replace(Me.cboPolicy, "'", "''") & "'", dbOpenSnapshot)
'
'     If Not rs.EOF Then
'         Me.lblCurrentStage.Caption = Nz(rs!CurrentStage, "")
'         Me.lblSector.Caption = Nz(rs!Sector, "")
'         Me.lblOwnerDepartment.Caption = Nz(rs!OwnerDepartment, "")
'         Me.lblResponsiblePerson.Caption = Nz(rs!ResponsiblePerson, "")
'         Me.lblSLAStatus.Caption = Nz(rs!CurrentSLAStatus, "")
'         Me.lblCreatedDate.Caption = Format(rs!CreatedDate, "yyyy-mm-dd hh:nn")
'         Me.lblStageStart.Caption = Format(rs!CurrentStageStart, "yyyy-mm-dd hh:nn")
'
'         ' Update ReceivedFrom dropdown based on current stage
'         UpdateReceivedFromDropdown Nz(rs!CurrentStage, "")
'
'         ' Enable/disable stage change based on current stage
'         Dim currentStage As String
'         currentStage = Nz(rs!CurrentStage, "")
'         If currentStage = "معتمدة" Or currentStage = "ملغي / مرفوض" Then
'             Me.cboNewStage.Enabled = False
'             Me.cboReceivedFrom.Enabled = False
'             Me.btnUpdateStage.Enabled = False
'             Me.btnRegisterReceived.Enabled = False
'         Else
'             Me.cboNewStage.Enabled = True
'             Me.cboReceivedFrom.Enabled = True
'             Me.btnUpdateStage.Enabled = True
'             Me.btnRegisterReceived.Enabled = True
'         End If
'     End If
'
'     rs.Close
'     Set rs = Nothing
' End Sub
'
' Private Sub UpdateReceivedFromDropdown(currentStage As String)
'     Dim options As String
'     options = GetReceivedFromOptions(currentStage)
'
'     If options = "" Then
'         Me.cboReceivedFrom.RowSource = ""
'         Me.cboReceivedFrom.Enabled = False
'     Else
'         Me.cboReceivedFrom.RowSourceType = "Value List"
'         Me.cboReceivedFrom.RowSource = options
'         Me.cboReceivedFrom.Enabled = True
'     End If
'     Me.cboReceivedFrom = Null
' End Sub
'
' Private Sub btnUpdateStage_Click()
'     If IsNull(Me.cboPolicy) Then
'         MsgBox "يرجى اختيار سياسة أولاً", vbExclamation, "تنبيه"
'         Exit Sub
'     End If
'
'     If IsNull(Me.cboNewStage) Then
'         MsgBox "يرجى اختيار المرحلة الجديدة", vbExclamation, "تنبيه"
'         Exit Sub
'     End If
'
'     UpdatePolicyStage Me.cboPolicy, Me.cboNewStage, Nz(Me.txtNotes, "")
'
'     ' Refresh display
'     Me.cboNewStage = Null
'     Me.txtNotes = ""
'     cboPolicy_AfterUpdate ' Refresh info
' End Sub
'
' Private Sub btnRegisterReceived_Click()
'     If IsNull(Me.cboPolicy) Then
'         MsgBox "يرجى اختيار سياسة أولاً", vbExclamation, "تنبيه"
'         Exit Sub
'     End If
'
'     If IsNull(Me.cboReceivedFrom) Then
'         MsgBox "يرجى اختيار الجهة", vbExclamation, "تنبيه"
'         Exit Sub
'     End If
'
'     RegisterReceivedFrom Me.cboPolicy, Me.cboReceivedFrom, Nz(Me.txtNotes, "")
'
'     ' Refresh display
'     Me.cboReceivedFrom = Null
'     Me.txtNotes = ""
'     cboPolicy_AfterUpdate ' Refresh info
' End Sub
'
' Private Sub btnClose_Click()
'     DoCmd.Close acForm, Me.Name
' End Sub

' ============================================================
' frmDashboard - Load Event
' ============================================================
' Place this code in frmDashboard's Form module:
'
' Private Sub Form_Load()
'     RefreshDashboard
' End Sub
'
' Private Sub btnRefresh_Click()
'     RefreshDashboard
' End Sub
'
' Private Sub RefreshDashboard()
'     RefreshAllSLAStatuses
'
'     Me.lblTotalPolicies.Caption = GetDashboardValue("Total")
'     Me.lblRegistration.Caption = GetDashboardValue("Registration")
'     Me.lblTechnicalReview.Caption = GetDashboardValue("TechnicalReview")
'     Me.lblGeneralReview.Caption = GetDashboardValue("GeneralReview")
'     Me.lblApprovals.Caption = GetDashboardValue("Approvals")
'     Me.lblApproved.Caption = GetDashboardValue("Approved")
'     Me.lblCancelled.Caption = GetDashboardValue("Cancelled")
'     Me.lblOverdueOLA.Caption = GetDashboardValue("OverdueOLA")
'     Me.lblNearOverdue.Caption = GetDashboardValue("NearOverdue")
'
'     Me.lblAvgRegistration.Caption = GetAverageDurationByStage("مرحلة التسجيل") & " يوم"
'     Me.lblAvgTechnical.Caption = GetAverageDurationByStage("مرحلة المراجعة الفنية") & " يوم"
'     Me.lblAvgGeneral.Caption = GetAverageDurationByStage("مرحلة المراجعة العمومية") & " يوم"
'     Me.lblAvgApprovals.Caption = GetAverageDurationByStage("مرحلة الإعتمادات") & " يوم"
'
'     Me.lblTopDelay.Caption = GetTopDelayDepartment()
' End Sub
'
' Private Sub btnClose_Click()
'     DoCmd.Close acForm, Me.Name
' End Sub
