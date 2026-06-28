Attribute VB_Name = "modFormCreator"
Option Compare Database
Option Explicit

' ============================================================
' Form Creator Module
' Creates all forms programmatically
' Run CreateAllForms from the Immediate window
' ============================================================

Public Sub CreateAllForms()
    On Error GoTo ErrHandler

    CreateMainMenu
    CreateAddPolicyForm
    CreateUpdatePolicyForm
    CreateViewPoliciesForm
    CreateWorkflowLogForm
    CreateSLASetupForm
    CreateDashboardForm

    MsgBox "تم إنشاء جميع النماذج بنجاح!" & vbCrLf & _
        "يرجى إغلاق وإعادة فتح قاعدة البيانات.", vbInformation, "نجاح"
    Exit Sub

ErrHandler:
    MsgBox "خطأ: " & Err.Description & " (رقم: " & Err.Number & ")", vbCritical, "خطأ"
End Sub

' ============================================================
' frmMainMenu
' ============================================================
Private Sub CreateMainMenu()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmMainMenu"
    On Error GoTo 0

    Dim frm As Form
    Set frm = CreateForm

    frm.Caption = "نظام متابعة السياسات - Policy Tracker Management System"
    frm.DefaultView = 0 ' Single Form
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.ScrollBars = 0 ' Neither
    frm.Section(acDetail).Height = 7000
    frm.Width = 8000
    frm.Section(acDetail).BackColor = RGB(240, 248, 255)

    ' Title
    Dim lblTitle As Control
    Set lblTitle = CreateControl(frm.Name, acLabel, acDetail, , , 500, 200, 7000, 600)
    lblTitle.Caption = "نظام متابعة السياسات"
    lblTitle.FontSize = 22
    lblTitle.FontBold = True
    lblTitle.ForeColor = RGB(0, 51, 102)
    lblTitle.TextAlign = 2 ' Center

    Dim lblSubtitle As Control
    Set lblSubtitle = CreateControl(frm.Name, acLabel, acDetail, , , 500, 800, 7000, 400)
    lblSubtitle.Caption = "Policy Tracker Management System"
    lblSubtitle.FontSize = 14
    lblSubtitle.ForeColor = RGB(100, 100, 100)
    lblSubtitle.TextAlign = 2

    ' Buttons
    Dim btnNames() As String
    btnNames = Split("btnAddPolicy,btnUpdatePolicy,btnViewPolicies,btnWorkflowLog,btnSLASetup,btnDashboard,btnExit", ",")
    Dim btnCaptions() As String
    btnCaptions = Split("إضافة سياسة جديدة,تحديث سياسة,عرض السياسات,سجل الحركات,إعدادات OLA/SLA,لوحة المتابعة,خروج", ",")
    Dim btnForms() As String
    btnForms = Split("frmAddPolicy,frmUpdatePolicy,frmViewPolicies,frmWorkflowLog,frmSLASetup,frmDashboard,EXIT", ",")

    Dim i As Integer
    Dim topPos As Long
    For i = 0 To 6
        topPos = 1500 + (i * 750)
        Dim btn As Control
        Set btn = CreateControl(frm.Name, acCommandButton, acDetail, , , 1500, topPos, 5000, 550)
        btn.Name = btnNames(i)
        btn.Caption = btnCaptions(i)
        btn.FontSize = 12
        btn.FontBold = True

        If i = 6 Then
            btn.ForeColor = RGB(255, 255, 255)
            btn.BackColor = RGB(192, 0, 0)
        Else
            btn.ForeColor = RGB(255, 255, 255)
            btn.BackColor = RGB(0, 102, 153)
        End If

        ' Add click event
        Dim eventCode As String
        If btnForms(i) = "EXIT" Then
            eventCode = "Private Sub " & btnNames(i) & "_Click()" & vbCrLf & _
                "    DoCmd.Quit" & vbCrLf & "End Sub"
        Else
            eventCode = "Private Sub " & btnNames(i) & "_Click()" & vbCrLf & _
                "    DoCmd.OpenForm """ & btnForms(i) & """, acNormal" & vbCrLf & "End Sub"
        End If

        frm.Module.InsertLines frm.Module.CountOfLines + 1, eventCode & vbCrLf
    Next i

    ' Add Form_Open event to refresh SLA
    Dim openEvent As String
    openEvent = "Private Sub Form_Open(Cancel As Integer)" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    RefreshAllSLAStatuses" & vbCrLf & _
        "End Sub"
    frm.Module.InsertLines frm.Module.CountOfLines + 1, openEvent

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name
    DoCmd.Rename "frmMainMenu", acForm, frm.Name
End Sub

' ============================================================
' frmAddPolicy
' ============================================================
Private Sub CreateAddPolicyForm()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmAddPolicy"
    On Error GoTo 0

    Dim frm As Form
    Set frm = CreateForm

    frm.Caption = "إضافة سياسة جديدة"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 0
    frm.Section(acDetail).Height = 5500
    frm.Width = 7500
    frm.Section(acDetail).BackColor = RGB(245, 245, 250)

    ' Title
    Dim lblTitle As Control
    Set lblTitle = CreateControl(frm.Name, acLabel, acDetail, , , 200, 100, 7000, 500)
    lblTitle.Caption = "إضافة سياسة جديدة"
    lblTitle.FontSize = 18
    lblTitle.FontBold = True
    lblTitle.ForeColor = RGB(0, 51, 102)

    ' Fields
    Dim fieldNames() As String
    fieldNames = Split("txtPolicyID,txtPolicyName,cboSector,cboOwnerDepartment,txtResponsiblePerson,txtNotes", ",")
    Dim fieldLabels() As String
    fieldLabels = Split("رقم السياسة:,اسم السياسة:,القطاع:,الإدارة المالكة:,المسؤول:,ملاحظات:", ",")

    Dim y As Long
    y = 800

    ' PolicyID
    Dim lbl As Control
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 350)
    lbl.Caption = "رقم السياسة:"
    lbl.FontSize = 11
    Dim txt As Control
    Set txt = CreateControl(frm.Name, acTextBox, acDetail, , , 2300, y, 4500, 350)
    txt.Name = "txtPolicyID"
    txt.FontSize = 11
    y = y + 500

    ' PolicyName
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 350)
    lbl.Caption = "اسم السياسة:"
    lbl.FontSize = 11
    Set txt = CreateControl(frm.Name, acTextBox, acDetail, , , 2300, y, 4500, 350)
    txt.Name = "txtPolicyName"
    txt.FontSize = 11
    y = y + 500

    ' Sector (Combo)
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 350)
    lbl.Caption = "القطاع:"
    lbl.FontSize = 11
    Dim cbo As Control
    Set cbo = CreateControl(frm.Name, acComboBox, acDetail, , , 2300, y, 4500, 350)
    cbo.Name = "cboSector"
    cbo.RowSourceType = "Value List"
    cbo.RowSource = "'قطاع التقنية','قطاع الحوكمة','قطاع المخاطر','قطاع الموارد البشرية','قطاع البيئة','قطاع المالية'"
    cbo.FontSize = 11
    cbo.LimitToList = True
    y = y + 500

    ' OwnerDepartment (Combo)
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 350)
    lbl.Caption = "الإدارة المالكة:"
    lbl.FontSize = 11
    Set cbo = CreateControl(frm.Name, acComboBox, acDetail, , , 2300, y, 4500, 350)
    cbo.Name = "cboOwnerDepartment"
    cbo.RowSourceType = "Value List"
    cbo.RowSource = "'الإدارة المالكة','الإدارة الفنية','إدارة الحوكمة والامتثال','الإدارة العامة للقانونية والتشريعات','الإدارة العامة للتميز'"
    cbo.FontSize = 11
    cbo.LimitToList = True
    y = y + 500

    ' ResponsiblePerson
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 350)
    lbl.Caption = "المسؤول:"
    lbl.FontSize = 11
    Set txt = CreateControl(frm.Name, acTextBox, acDetail, , , 2300, y, 4500, 350)
    txt.Name = "txtResponsiblePerson"
    txt.FontSize = 11
    y = y + 500

    ' Notes
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 350)
    lbl.Caption = "ملاحظات:"
    lbl.FontSize = 11
    Set txt = CreateControl(frm.Name, acTextBox, acDetail, , , 2300, y, 4500, 700)
    txt.Name = "txtNotes"
    txt.FontSize = 11
    y = y + 850

    ' Buttons
    Dim btnSave As Control
    Set btnSave = CreateControl(frm.Name, acCommandButton, acDetail, , , 2300, y, 2000, 450)
    btnSave.Name = "btnSave"
    btnSave.Caption = "حفظ"
    btnSave.FontSize = 12
    btnSave.FontBold = True
    btnSave.ForeColor = RGB(255, 255, 255)
    btnSave.BackColor = RGB(0, 128, 0)

    Dim btnClose As Control
    Set btnClose = CreateControl(frm.Name, acCommandButton, acDetail, , , 4500, y, 2000, 450)
    btnClose.Name = "btnClose"
    btnClose.Caption = "إغلاق"
    btnClose.FontSize = 12
    btnClose.FontBold = True
    btnClose.ForeColor = RGB(255, 255, 255)
    btnClose.BackColor = RGB(128, 128, 128)

    ' Add VBA code
    Dim code As String
    code = "Private Sub btnSave_Click()" & vbCrLf & _
        "    If Nz(Me.txtPolicyID, """") = """" Then" & vbCrLf & _
        "        MsgBox ""يرجى إدخال رقم السياسة"", vbExclamation, ""تنبيه""" & vbCrLf & _
        "        Me.txtPolicyID.SetFocus: Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    If Nz(Me.txtPolicyName, """") = """" Then" & vbCrLf & _
        "        MsgBox ""يرجى إدخال اسم السياسة"", vbExclamation, ""تنبيه""" & vbCrLf & _
        "        Me.txtPolicyName.SetFocus: Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    If Nz(Me.cboSector, """") = """" Then" & vbCrLf & _
        "        MsgBox ""يرجى اختيار القطاع"", vbExclamation, ""تنبيه""" & vbCrLf & _
        "        Me.cboSector.SetFocus: Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    If Nz(Me.cboOwnerDepartment, """") = """" Then" & vbCrLf & _
        "        MsgBox ""يرجى اختيار الإدارة المالكة"", vbExclamation, ""تنبيه""" & vbCrLf & _
        "        Me.cboOwnerDepartment.SetFocus: Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    If Nz(Me.txtResponsiblePerson, """") = """" Then" & vbCrLf & _
        "        MsgBox ""يرجى إدخال اسم المسؤول"", vbExclamation, ""تنبيه""" & vbCrLf & _
        "        Me.txtResponsiblePerson.SetFocus: Exit Sub" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    AddNewPolicy Me.txtPolicyID, Me.txtPolicyName, Me.cboSector, Me.cboOwnerDepartment, Me.txtResponsiblePerson, Nz(Me.txtNotes, """")" & vbCrLf & _
        "    Me.txtPolicyID = """": Me.txtPolicyName = """": Me.cboSector = Null" & vbCrLf & _
        "    Me.cboOwnerDepartment = Null: Me.txtResponsiblePerson = """": Me.txtNotes = """"" & vbCrLf & _
        "    Me.txtPolicyID.SetFocus" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub btnClose_Click()" & vbCrLf & _
        "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
        "End Sub"

    frm.Module.InsertLines frm.Module.CountOfLines + 1, code

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name
    DoCmd.Rename "frmAddPolicy", acForm, frm.Name
End Sub

' ============================================================
' frmUpdatePolicy
' ============================================================
Private Sub CreateUpdatePolicyForm()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmUpdatePolicy"
    On Error GoTo 0

    Dim frm As Form
    Set frm = CreateForm

    frm.Caption = "تحديث سياسة"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 0
    frm.Section(acDetail).Height = 8500
    frm.Width = 8000
    frm.Section(acDetail).BackColor = RGB(245, 245, 250)

    ' Title
    Dim lblTitle As Control
    Set lblTitle = CreateControl(frm.Name, acLabel, acDetail, , , 200, 100, 7500, 500)
    lblTitle.Caption = "تحديث سياسة"
    lblTitle.FontSize = 18
    lblTitle.FontBold = True
    lblTitle.ForeColor = RGB(0, 51, 102)

    Dim y As Long
    y = 700

    ' Policy Selection
    Dim lbl As Control
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 350)
    lbl.Caption = "اختر السياسة:"
    lbl.FontSize = 11
    lbl.FontBold = True

    Dim cboPolicy As Control
    Set cboPolicy = CreateControl(frm.Name, acComboBox, acDetail, , , 2300, y, 5200, 350)
    cboPolicy.Name = "cboPolicy"
    cboPolicy.RowSourceType = "Table/Query"
    cboPolicy.RowSource = "SELECT PolicyID, PolicyID & ' - ' & PolicyName FROM tblPolicies ORDER BY PolicyID"
    cboPolicy.ColumnCount = 2
    cboPolicy.ColumnWidths = "1500;5000"
    cboPolicy.BoundColumn = 1
    cboPolicy.FontSize = 11
    cboPolicy.LimitToList = True
    y = y + 500

    ' Info Display Section
    Dim lblSection As Control
    Set lblSection = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 7500, 350)
    lblSection.Caption = "معلومات السياسة الحالية"
    lblSection.FontSize = 12
    lblSection.FontBold = True
    lblSection.ForeColor = RGB(0, 102, 153)
    y = y + 400

    ' Current Stage
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 300)
    lbl.Caption = "المرحلة الحالية:"
    lbl.FontSize = 10
    Dim lblInfo As Control
    Set lblInfo = CreateControl(frm.Name, acLabel, acDetail, , , 2300, y, 5200, 300)
    lblInfo.Name = "lblCurrentStage"
    lblInfo.Caption = " "
    lblInfo.FontSize = 10
    lblInfo.FontBold = True
    lblInfo.ForeColor = RGB(0, 0, 180)
    y = y + 350

    ' Sector
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 300)
    lbl.Caption = "القطاع:"
    lbl.FontSize = 10
    Set lblInfo = CreateControl(frm.Name, acLabel, acDetail, , , 2300, y, 5200, 300)
    lblInfo.Name = "lblSector"
    lblInfo.Caption = " "
    lblInfo.FontSize = 10
    y = y + 350

    ' Owner Department
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 300)
    lbl.Caption = "الإدارة المالكة:"
    lbl.FontSize = 10
    Set lblInfo = CreateControl(frm.Name, acLabel, acDetail, , , 2300, y, 5200, 300)
    lblInfo.Name = "lblOwnerDepartment"
    lblInfo.Caption = " "
    lblInfo.FontSize = 10
    y = y + 350

    ' Responsible Person
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 300)
    lbl.Caption = "المسؤول:"
    lbl.FontSize = 10
    Set lblInfo = CreateControl(frm.Name, acLabel, acDetail, , , 2300, y, 5200, 300)
    lblInfo.Name = "lblResponsiblePerson"
    lblInfo.Caption = " "
    lblInfo.FontSize = 10
    y = y + 350

    ' SLA Status
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 300)
    lbl.Caption = "حالة OLA:"
    lbl.FontSize = 10
    Set lblInfo = CreateControl(frm.Name, acLabel, acDetail, , , 2300, y, 5200, 300)
    lblInfo.Name = "lblSLAStatus"
    lblInfo.Caption = " "
    lblInfo.FontSize = 10
    lblInfo.FontBold = True
    y = y + 350

    ' Created Date
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 300)
    lbl.Caption = "تاريخ الإنشاء:"
    lbl.FontSize = 10
    Set lblInfo = CreateControl(frm.Name, acLabel, acDetail, , , 2300, y, 5200, 300)
    lblInfo.Name = "lblCreatedDate"
    lblInfo.Caption = " "
    lblInfo.FontSize = 10
    y = y + 350

    ' Stage Start
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 300)
    lbl.Caption = "بداية المرحلة:"
    lbl.FontSize = 10
    Set lblInfo = CreateControl(frm.Name, acLabel, acDetail, , , 2300, y, 5200, 300)
    lblInfo.Name = "lblStageStart"
    lblInfo.Caption = " "
    lblInfo.FontSize = 10
    y = y + 500

    ' Action Section - Change Stage
    Set lblSection = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 7500, 350)
    lblSection.Caption = "تغيير المرحلة"
    lblSection.FontSize = 12
    lblSection.FontBold = True
    lblSection.ForeColor = RGB(0, 128, 0)
    y = y + 400

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 350)
    lbl.Caption = "المرحلة الجديدة:"
    lbl.FontSize = 11

    Dim cboNewStage As Control
    Set cboNewStage = CreateControl(frm.Name, acComboBox, acDetail, , , 2300, y, 3500, 350)
    cboNewStage.Name = "cboNewStage"
    cboNewStage.RowSourceType = "Table/Query"
    cboNewStage.RowSource = "SELECT StageName FROM tblStages ORDER BY StageOrder"
    cboNewStage.FontSize = 11
    cboNewStage.LimitToList = True

    Dim btnUpdateStage As Control
    Set btnUpdateStage = CreateControl(frm.Name, acCommandButton, acDetail, , , 6000, y, 1500, 350)
    btnUpdateStage.Name = "btnUpdateStage"
    btnUpdateStage.Caption = "تحديث المرحلة"
    btnUpdateStage.FontSize = 10
    btnUpdateStage.ForeColor = RGB(255, 255, 255)
    btnUpdateStage.BackColor = RGB(0, 128, 0)
    y = y + 500

    ' Action Section - Received From
    Set lblSection = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 7500, 350)
    lblSection.Caption = "تسجيل استلام من جهة"
    lblSection.FontSize = 12
    lblSection.FontBold = True
    lblSection.ForeColor = RGB(153, 102, 0)
    y = y + 400

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 350)
    lbl.Caption = "تم الاستلام من:"
    lbl.FontSize = 11

    Dim cboReceivedFrom As Control
    Set cboReceivedFrom = CreateControl(frm.Name, acComboBox, acDetail, , , 2300, y, 3500, 350)
    cboReceivedFrom.Name = "cboReceivedFrom"
    cboReceivedFrom.RowSourceType = "Value List"
    cboReceivedFrom.FontSize = 11
    cboReceivedFrom.LimitToList = True

    Dim btnRegisterReceived As Control
    Set btnRegisterReceived = CreateControl(frm.Name, acCommandButton, acDetail, , , 6000, y, 1500, 350)
    btnRegisterReceived.Name = "btnRegisterReceived"
    btnRegisterReceived.Caption = "تسجيل الاستلام"
    btnRegisterReceived.FontSize = 10
    btnRegisterReceived.ForeColor = RGB(255, 255, 255)
    btnRegisterReceived.BackColor = RGB(153, 102, 0)
    y = y + 500

    ' Notes
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 2000, 350)
    lbl.Caption = "ملاحظات:"
    lbl.FontSize = 11

    Dim txtNotes As Control
    Set txtNotes = CreateControl(frm.Name, acTextBox, acDetail, , , 2300, y, 5200, 600)
    txtNotes.Name = "txtNotes"
    txtNotes.FontSize = 11
    y = y + 750

    ' Close button
    Dim btnClose As Control
    Set btnClose = CreateControl(frm.Name, acCommandButton, acDetail, , , 3000, y, 2000, 450)
    btnClose.Name = "btnClose"
    btnClose.Caption = "إغلاق"
    btnClose.FontSize = 12
    btnClose.ForeColor = RGB(255, 255, 255)
    btnClose.BackColor = RGB(128, 128, 128)

    ' Add VBA code for the form
    Dim code As String
    code = "Private Sub cboPolicy_AfterUpdate()" & vbCrLf & _
        "    If IsNull(Me.cboPolicy) Then Exit Sub" & vbCrLf & _
        "    Dim rs As DAO.Recordset" & vbCrLf & _
        "    Set rs = CurrentDb.OpenRecordset(""SELECT * FROM tblPolicies WHERE PolicyID='"" & Replace(Me.cboPolicy, ""'"", ""''"") & ""'"", dbOpenSnapshot)" & vbCrLf & _
        "    If Not rs.EOF Then" & vbCrLf & _
        "        Me.lblCurrentStage.Caption = Nz(rs!CurrentStage, """")" & vbCrLf & _
        "        Me.lblSector.Caption = Nz(rs!Sector, """")" & vbCrLf & _
        "        Me.lblOwnerDepartment.Caption = Nz(rs!OwnerDepartment, """")" & vbCrLf & _
        "        Me.lblResponsiblePerson.Caption = Nz(rs!ResponsiblePerson, """")" & vbCrLf & _
        "        Me.lblSLAStatus.Caption = Nz(rs!CurrentSLAStatus, """")" & vbCrLf & _
        "        Me.lblCreatedDate.Caption = Format(rs!CreatedDate, ""yyyy-mm-dd hh:nn"")" & vbCrLf & _
        "        Me.lblStageStart.Caption = Format(rs!CurrentStageStart, ""yyyy-mm-dd hh:nn"")" & vbCrLf & _
        "        UpdateReceivedFromDropdown Nz(rs!CurrentStage, """")" & vbCrLf & _
        "        Dim cs As String: cs = Nz(rs!CurrentStage, """")" & vbCrLf & _
        "        If cs = ""معتمدة"" Or cs = ""ملغي / مرفوض"" Then" & vbCrLf & _
        "            Me.cboNewStage.Enabled = False: Me.cboReceivedFrom.Enabled = False" & vbCrLf & _
        "            Me.btnUpdateStage.Enabled = False: Me.btnRegisterReceived.Enabled = False" & vbCrLf & _
        "        Else" & vbCrLf & _
        "            Me.cboNewStage.Enabled = True: Me.cboReceivedFrom.Enabled = True" & vbCrLf & _
        "            Me.btnUpdateStage.Enabled = True: Me.btnRegisterReceived.Enabled = True" & vbCrLf & _
        "        End If" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    rs.Close: Set rs = Nothing" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub UpdateReceivedFromDropdown(currentStage As String)" & vbCrLf & _
        "    Dim options As String" & vbCrLf & _
        "    options = GetReceivedFromOptions(currentStage)" & vbCrLf & _
        "    If options = """" Then" & vbCrLf & _
        "        Me.cboReceivedFrom.RowSource = """": Me.cboReceivedFrom.Enabled = False" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        Me.cboReceivedFrom.RowSourceType = ""Value List""" & vbCrLf & _
        "        Me.cboReceivedFrom.RowSource = options: Me.cboReceivedFrom.Enabled = True" & vbCrLf & _
        "    End If" & vbCrLf & _
        "    Me.cboReceivedFrom = Null" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub btnUpdateStage_Click()" & vbCrLf & _
        "    If IsNull(Me.cboPolicy) Then MsgBox ""يرجى اختيار سياسة أولاً"", vbExclamation, ""تنبيه"": Exit Sub" & vbCrLf & _
        "    If IsNull(Me.cboNewStage) Then MsgBox ""يرجى اختيار المرحلة الجديدة"", vbExclamation, ""تنبيه"": Exit Sub" & vbCrLf & _
        "    UpdatePolicyStage Me.cboPolicy, Me.cboNewStage, Nz(Me.txtNotes, """")" & vbCrLf & _
        "    Me.cboNewStage = Null: Me.txtNotes = """": cboPolicy_AfterUpdate" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub btnRegisterReceived_Click()" & vbCrLf & _
        "    If IsNull(Me.cboPolicy) Then MsgBox ""يرجى اختيار سياسة أولاً"", vbExclamation, ""تنبيه"": Exit Sub" & vbCrLf & _
        "    If IsNull(Me.cboReceivedFrom) Then MsgBox ""يرجى اختيار الجهة"", vbExclamation, ""تنبيه"": Exit Sub" & vbCrLf & _
        "    RegisterReceivedFrom Me.cboPolicy, Me.cboReceivedFrom, Nz(Me.txtNotes, """")" & vbCrLf & _
        "    Me.cboReceivedFrom = Null: Me.txtNotes = """": cboPolicy_AfterUpdate" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf

    code = code & "Private Sub btnClose_Click()" & vbCrLf & _
        "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
        "End Sub"

    frm.Module.InsertLines frm.Module.CountOfLines + 1, code

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name
    DoCmd.Rename "frmUpdatePolicy", acForm, frm.Name
End Sub

' ============================================================
' frmViewPolicies
' ============================================================
Private Sub CreateViewPoliciesForm()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmViewPolicies"
    On Error GoTo 0

    Dim frm As Form
    Set frm = CreateForm(, "tblPolicies")

    frm.Caption = "عرض السياسات"
    frm.DefaultView = 2 ' Datasheet
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = False
    frm.NavigationButtons = True

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name
    DoCmd.Rename "frmViewPolicies", acForm, frm.Name
End Sub

' ============================================================
' frmWorkflowLog
' ============================================================
Private Sub CreateWorkflowLogForm()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmWorkflowLog"
    On Error GoTo 0

    Dim frm As Form
    Set frm = CreateForm

    frm.Caption = "سجل الحركات - Workflow Log"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 0
    frm.Section(acDetail).Height = 7500
    frm.Width = 8000
    frm.Section(acDetail).BackColor = RGB(245, 245, 250)

    ' Title
    Dim lblTitle As Control
    Set lblTitle = CreateControl(frm.Name, acLabel, acDetail, , , 200, 100, 7500, 500)
    lblTitle.Caption = "سجل الحركات"
    lblTitle.FontSize = 18
    lblTitle.FontBold = True
    lblTitle.ForeColor = RGB(0, 51, 102)

    ' Policy filter
    Dim lbl As Control
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, 700, 2000, 350)
    lbl.Caption = "فلترة بالسياسة:"
    lbl.FontSize = 11

    Dim cboFilter As Control
    Set cboFilter = CreateControl(frm.Name, acComboBox, acDetail, , , 2300, 700, 3500, 350)
    cboFilter.Name = "cboFilterPolicy"
    cboFilter.RowSourceType = "Table/Query"
    cboFilter.RowSource = "SELECT PolicyID, PolicyID & ' - ' & PolicyName FROM tblPolicies ORDER BY PolicyID"
    cboFilter.ColumnCount = 2
    cboFilter.ColumnWidths = "1500;5000"
    cboFilter.BoundColumn = 1
    cboFilter.FontSize = 11

    Dim btnFilter As Control
    Set btnFilter = CreateControl(frm.Name, acCommandButton, acDetail, , , 6000, 700, 1500, 350)
    btnFilter.Name = "btnFilter"
    btnFilter.Caption = "فلترة"
    btnFilter.FontSize = 10
    btnFilter.ForeColor = RGB(255, 255, 255)
    btnFilter.BackColor = RGB(0, 102, 153)

    ' Subform for log data
    Dim sfrmLog As Control
    Set sfrmLog = CreateControl(frm.Name, acSubform, acDetail, , , 200, 1200, 7500, 5500)
    sfrmLog.Name = "sfrmWorkflowLog"
    sfrmLog.SourceObject = "Table.tblWorkflowLog"

    ' Close button
    Dim btnClose As Control
    Set btnClose = CreateControl(frm.Name, acCommandButton, acDetail, , , 3000, 6900, 2000, 450)
    btnClose.Name = "btnClose"
    btnClose.Caption = "إغلاق"
    btnClose.FontSize = 12
    btnClose.ForeColor = RGB(255, 255, 255)
    btnClose.BackColor = RGB(128, 128, 128)

    Dim code As String
    code = "Private Sub btnFilter_Click()" & vbCrLf & _
        "    If IsNull(Me.cboFilterPolicy) Then" & vbCrLf & _
        "        Me.sfrmWorkflowLog.SourceObject = ""Table.tblWorkflowLog""" & vbCrLf & _
        "    Else" & vbCrLf & _
        "        Me.sfrmWorkflowLog.SourceObject = ""Table.tblWorkflowLog""" & vbCrLf & _
        "        Me.sfrmWorkflowLog.Form.Filter = ""PolicyID='"" & Me.cboFilterPolicy & ""'""" & vbCrLf & _
        "        Me.sfrmWorkflowLog.Form.FilterOn = True" & vbCrLf & _
        "    End If" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub btnClose_Click()" & vbCrLf & _
        "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
        "End Sub"

    frm.Module.InsertLines frm.Module.CountOfLines + 1, code

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name
    DoCmd.Rename "frmWorkflowLog", acForm, frm.Name
End Sub

' ============================================================
' frmSLASetup
' ============================================================
Private Sub CreateSLASetupForm()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmSLASetup"
    On Error GoTo 0

    Dim frm As Form
    Set frm = CreateForm(, "tblSLASetup")

    frm.Caption = "إعدادات OLA/SLA"
    frm.DefaultView = 2 ' Datasheet
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    frm.AllowEdits = True
    frm.NavigationButtons = True

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name
    DoCmd.Rename "frmSLASetup", acForm, frm.Name
End Sub

' ============================================================
' frmDashboard
' ============================================================
Private Sub CreateDashboardForm()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmDashboard"
    On Error GoTo 0

    Dim frm As Form
    Set frm = CreateForm

    frm.Caption = "لوحة المتابعة - Dashboard"
    frm.DefaultView = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 2 ' Vertical only
    frm.Section(acDetail).Height = 10000
    frm.Width = 8500
    frm.Section(acDetail).BackColor = RGB(240, 248, 255)

    ' Title
    Dim lblTitle As Control
    Set lblTitle = CreateControl(frm.Name, acLabel, acDetail, , , 200, 100, 8000, 500)
    lblTitle.Caption = "لوحة المتابعة - Dashboard"
    lblTitle.FontSize = 20
    lblTitle.FontBold = True
    lblTitle.ForeColor = RGB(0, 51, 102)
    lblTitle.TextAlign = 2

    Dim y As Long
    y = 700

    ' Section: Policy Counts
    Dim lblSection As Control
    Set lblSection = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 8000, 400)
    lblSection.Caption = "إحصائيات السياسات"
    lblSection.FontSize = 14
    lblSection.FontBold = True
    lblSection.ForeColor = RGB(0, 102, 153)
    y = y + 450

    ' Dashboard cards
    Dim cardLabels() As String
    cardLabels = Split("إجمالي السياسات,مرحلة التسجيل,المراجعة الفنية,المراجعة العمومية,مرحلة الإعتمادات,معتمدة,ملغي / مرفوض", ",")
    Dim cardNames() As String
    cardNames = Split("lblTotalPolicies,lblRegistration,lblTechnicalReview,lblGeneralReview,lblApprovals,lblApproved,lblCancelled", ",")

    Dim i As Integer
    Dim xPos As Long

    For i = 0 To 6
        xPos = 200 + (i Mod 4) * 2050

        If i = 4 Then y = y + 850

        ' Card background label
        Dim lbl As Control
        Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , xPos, y, 1950, 350)
        lbl.Caption = cardLabels(i)
        lbl.FontSize = 9
        lbl.TextAlign = 2
        lbl.ForeColor = RGB(80, 80, 80)

        ' Card value
        Dim lblVal As Control
        Set lblVal = CreateControl(frm.Name, acLabel, acDetail, , , xPos, y + 350, 1950, 400)
        lblVal.Name = cardNames(i)
        lblVal.Caption = "0"
        lblVal.FontSize = 20
        lblVal.FontBold = True
        lblVal.TextAlign = 2
        lblVal.ForeColor = RGB(0, 51, 102)
    Next i

    y = y + 950

    ' Section: OLA Status
    Set lblSection = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 8000, 400)
    lblSection.Caption = "حالة OLA"
    lblSection.FontSize = 14
    lblSection.FontBold = True
    lblSection.ForeColor = RGB(192, 0, 0)
    y = y + 450

    ' Overdue OLA
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 3800, 350)
    lbl.Caption = "متجاوز OLA:"
    lbl.FontSize = 11
    Set lblVal = CreateControl(frm.Name, acLabel, acDetail, , , 4100, y, 1500, 350)
    lblVal.Name = "lblOverdueOLA"
    lblVal.Caption = "0"
    lblVal.FontSize = 14
    lblVal.FontBold = True
    lblVal.ForeColor = RGB(192, 0, 0)
    y = y + 400

    ' Near Overdue
    Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 3800, 350)
    lbl.Caption = "قريب من التجاوز:"
    lbl.FontSize = 11
    Set lblVal = CreateControl(frm.Name, acLabel, acDetail, , , 4100, y, 1500, 350)
    lblVal.Name = "lblNearOverdue"
    lblVal.Caption = "0"
    lblVal.FontSize = 14
    lblVal.FontBold = True
    lblVal.ForeColor = RGB(200, 150, 0)
    y = y + 500

    ' Section: Average Durations
    Set lblSection = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 8000, 400)
    lblSection.Caption = "متوسط المدد لكل مرحلة"
    lblSection.FontSize = 14
    lblSection.FontBold = True
    lblSection.ForeColor = RGB(0, 128, 0)
    y = y + 450

    Dim stageLabels() As String
    stageLabels = Split("مرحلة التسجيل:,المراجعة الفنية:,المراجعة العمومية:,مرحلة الإعتمادات:", ",")
    Dim avgNames() As String
    avgNames = Split("lblAvgRegistration,lblAvgTechnical,lblAvgGeneral,lblAvgApprovals", ",")

    For i = 0 To 3
        Set lbl = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 3800, 350)
        lbl.Caption = stageLabels(i)
        lbl.FontSize = 11
        Set lblVal = CreateControl(frm.Name, acLabel, acDetail, , , 4100, y, 3000, 350)
        lblVal.Name = avgNames(i)
        lblVal.Caption = "0 يوم"
        lblVal.FontSize = 11
        lblVal.FontBold = True
        y = y + 400
    Next i

    y = y + 200

    ' Top Delay Department
    Set lblSection = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 8000, 400)
    lblSection.Caption = "أكثر جهة تسبب تأخير"
    lblSection.FontSize = 14
    lblSection.FontBold = True
    lblSection.ForeColor = RGB(153, 0, 0)
    y = y + 450

    Set lblVal = CreateControl(frm.Name, acLabel, acDetail, , , 200, y, 7500, 400)
    lblVal.Name = "lblTopDelay"
    lblVal.Caption = "لا يوجد"
    lblVal.FontSize = 13
    lblVal.FontBold = True
    lblVal.ForeColor = RGB(153, 0, 0)
    y = y + 600

    ' Buttons
    Dim btnRefresh As Control
    Set btnRefresh = CreateControl(frm.Name, acCommandButton, acDetail, , , 2000, y, 2000, 450)
    btnRefresh.Name = "btnRefresh"
    btnRefresh.Caption = "تحديث"
    btnRefresh.FontSize = 12
    btnRefresh.FontBold = True
    btnRefresh.ForeColor = RGB(255, 255, 255)
    btnRefresh.BackColor = RGB(0, 102, 153)

    Dim btnClose As Control
    Set btnClose = CreateControl(frm.Name, acCommandButton, acDetail, , , 4500, y, 2000, 450)
    btnClose.Name = "btnClose"
    btnClose.Caption = "إغلاق"
    btnClose.FontSize = 12
    btnClose.FontBold = True
    btnClose.ForeColor = RGB(255, 255, 255)
    btnClose.BackColor = RGB(128, 128, 128)

    ' Add VBA code
    Dim code As String
    code = "Private Sub Form_Load()" & vbCrLf & _
        "    RefreshDashboard" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub btnRefresh_Click()" & vbCrLf & _
        "    RefreshDashboard" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub RefreshDashboard()" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    RefreshAllSLAStatuses" & vbCrLf & _
        "    Me.lblTotalPolicies.Caption = GetDashboardValue(""Total"")" & vbCrLf & _
        "    Me.lblRegistration.Caption = GetDashboardValue(""Registration"")" & vbCrLf & _
        "    Me.lblTechnicalReview.Caption = GetDashboardValue(""TechnicalReview"")" & vbCrLf & _
        "    Me.lblGeneralReview.Caption = GetDashboardValue(""GeneralReview"")" & vbCrLf & _
        "    Me.lblApprovals.Caption = GetDashboardValue(""Approvals"")" & vbCrLf & _
        "    Me.lblApproved.Caption = GetDashboardValue(""Approved"")" & vbCrLf & _
        "    Me.lblCancelled.Caption = GetDashboardValue(""Cancelled"")" & vbCrLf & _
        "    Me.lblOverdueOLA.Caption = GetDashboardValue(""OverdueOLA"")" & vbCrLf & _
        "    Me.lblNearOverdue.Caption = GetDashboardValue(""NearOverdue"")" & vbCrLf & _
        "    Me.lblAvgRegistration.Caption = GetAverageDurationByStage(""مرحلة التسجيل"") & "" يوم""" & vbCrLf & _
        "    Me.lblAvgTechnical.Caption = GetAverageDurationByStage(""مرحلة المراجعة الفنية"") & "" يوم""" & vbCrLf & _
        "    Me.lblAvgGeneral.Caption = GetAverageDurationByStage(""مرحلة المراجعة العمومية"") & "" يوم""" & vbCrLf & _
        "    Me.lblAvgApprovals.Caption = GetAverageDurationByStage(""مرحلة الإعتمادات"") & "" يوم""" & vbCrLf & _
        "    Me.lblTopDelay.Caption = GetTopDelayDepartment()" & vbCrLf & _
        "End Sub" & vbCrLf & vbCrLf & _
        "Private Sub btnClose_Click()" & vbCrLf & _
        "    DoCmd.Close acForm, Me.Name" & vbCrLf & _
        "End Sub"

    frm.Module.InsertLines frm.Module.CountOfLines + 1, code

    DoCmd.Save acForm, frm.Name
    DoCmd.Close acForm, frm.Name
    DoCmd.Rename "frmDashboard", acForm, frm.Name
End Sub
