'#-------------------------------------------------------------------------------------
'# Created Date: 05/02/2009 11:20:09 AM
'# Created User: Thiên Huỳnh
'# Modify Date: 05/02/2009 11:20:09 AM
'# Modify User: Thiên Huỳnh
'#-------------------------------------------------------------------------------------
Public Class D02F6004
	Private _bSaved As Boolean = False
	Public ReadOnly Property bSaved() As Boolean
		Get
			Return _bSaved
		   End Get
	End Property


    Dim bLoadFormState As Boolean = False
	Private _FormState As EnumFormState
    Private _ReportCode As String
    

    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
	bLoadFormState = True
	LoadInfoGeneral()
            _FormState = value
            LoadTDBCombo()
            Select Case _FormState
                Case EnumFormState.FormAdd
                    CheckIdTextBox(txtReportCode, txtReportCode.MaxLength)
                    LoadDefault()
                    btnNext.Enabled = False
                    txtReportCode.Focus()
                    '  btnDefineData.Enabled = False
                Case EnumFormState.FormEdit
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    LoadEdit()
                Case EnumFormState.FormView

                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    btnSave.Enabled = False
                    LoadEdit()
            End Select
            ResetSelection(1)
            ResetLevel(1)
        End Set
    End Property

    Public Property m_ReportCode() As String
        Get
            Return _ReportCode
        End Get
        Set(ByVal value As String)
            _ReportCode = value
        End Set
    End Property

#Region "LoadCombo"

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Load tdbcReportID
        If _FormState = EnumFormState.FormAdd Then
            LoadtdbcReportID()
        End If
        'Load cac Cbo Add
        LoadTDBCUnit()
        LoadtdbcAmountFormat()
        LoadtdbcNegative()

        'Load tdbcGroup, tdbcSelection
        sSQL = " Select Code, Description" & UnicodeJoin(gbUnicode) & " Description From D02V3046 "
        sSQL += " Where Language = " & SQLString(gsLanguage)
        sSQL += " Order by  Code"
        Dim dtSelection As DataTable
        dtSelection = ReturnDataTable(sSQL)
        'Load 5 tdbcGroup
        LoadDataSource(tdbcLevel01, dtSelection.DefaultView.ToTable, gbUnicode)
        LoadDataSource(tdbcLevel02, dtSelection.DefaultView.ToTable, gbUnicode)
        LoadDataSource(tdbcLevel03, dtSelection.DefaultView.ToTable, gbUnicode)
        LoadDataSource(tdbcLevel04, dtSelection.DefaultView.ToTable, gbUnicode)
        LoadDataSource(tdbcLevel05, dtSelection.DefaultView.ToTable, gbUnicode)
        'Load 5 tdbcSelection
        LoadDataSource(tdbcSelection01, dtSelection.DefaultView.ToTable, gbUnicode)
        LoadDataSource(tdbcSelection02, dtSelection.DefaultView.ToTable, gbUnicode)
        LoadDataSource(tdbcSelection03, dtSelection.DefaultView.ToTable, gbUnicode)
        LoadDataSource(tdbcSelection04, dtSelection.DefaultView.ToTable, gbUnicode)
        LoadDataSource(tdbcSelection05, dtSelection.DefaultView.ToTable, gbUnicode)
    End Sub

    Private Sub LoadtdbcReportID()
        Dim sSQL As String = ""
        'Load tdbcReportID
        sSQL = " Select ReportID, " & IIf(geLanguage = EnumLanguage.Vietnamese, "Description", "Description01").ToString & UnicodeJoin(gbUnicode) & " as ReportName "
        sSQL &= " From D02V5555 Where ReportType = 'D02F6005'"
        sSQL &= " Order by ReportID"
        LoadDataSource(tdbcReportID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadTDBCUnit()
        Dim dt As New DataTable
        Dim dr As DataRow
        dt.Columns.Add("Code", System.Type.GetType("System.String"))
        dt.Columns.Add("Value", System.Type.GetType("System.String"))

        dr = dt.NewRow
        dr.Item("Code") = "0"
        dr.Item("Value") = "1"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr.Item("Code") = "1"
        dr.Item("Value") = "1.000"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr.Item("Code") = "2"
        dr.Item("Value") = "1.000.000"
        dt.Rows.Add(dr)

        LoadDataSource(tdbcUnit, dt, gbUnicode)
    End Sub

    Private Sub LoadtdbcAmountFormat()
        Dim dt As New DataTable
        Dim dr As DataRow
        dt.Columns.Add("Code", System.Type.GetType("System.String"))
        dt.Columns.Add("Value", System.Type.GetType("System.String"))

        dr = dt.NewRow
        dr.Item("Code") = "0"
        dr.Item("Value") = "1"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr.Item("Code") = "1"
        dr.Item("Value") = "0.1"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr.Item("Code") = "2"
        dr.Item("Value") = "0.01"
        dt.Rows.Add(dr)

        LoadDataSource(tdbcDecimalNegative, dt, gbUnicode)
    End Sub

    Private Sub LoadtdbcNegative()
        Dim dt As New DataTable
        Dim dr As DataRow
        dt.Columns.Add("Code", System.Type.GetType("System.String"))
        dt.Columns.Add("Value", System.Type.GetType("System.String"))

        dr = dt.NewRow
        dr.Item("Code") = "0"
        dr.Item("Value") = "None"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr.Item("Code") = "1"
        dr.Item("Value") = "-123"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr.Item("Code") = "2"
        dr.Item("Value") = "(123)"
        dt.Rows.Add(dr)

        LoadDataSource(tdbcBracketPlaces, dt, gbUnicode)
    End Sub

#End Region

#Region "Form Load"

    Private Sub LoadEdit()
        txtReportCode.ReadOnly = True
        Dim sSQL As String
        Dim dt As DataTable
        Dim _ReportID As String = ""

        sSQL = "Select ReportName1" & UnicodeJoin(gbUnicode) & " as ReportName1, ReportName2" & UnicodeJoin(gbUnicode) & " as ReportName2, * From D02T3040  WITH (NOLOCK) Where ReportCode= " & SQLString(_ReportCode)
        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            With dt.Rows(0)

                txtReportCode.Text = .Item("ReportCode").ToString
                txtReportName1.Text = .Item("ReportName1").ToString
                txtReportName2.Text = .Item("ReportName2").ToString
                tdbcUnit.SelectedValue = .Item("AmountFormat").ToString
                tdbcDecimalNegative.SelectedValue = .Item("DecimalPlaces").ToString
                tdbcBracketPlaces.SelectedValue = .Item("BracketNegative").ToString
                chkCustomized.Checked = CBool(.Item("Customized"))
                _ReportID = .Item("ReportID").ToString
                txtCustomizedReportID.Text = .Item("ReportID").ToString
                chkDisabled.Checked = CBool(.Item("Disabled"))

                If CBool(.Item("Customized")) Then
                    tdbcReportID.Visible = False
                    txtReportName.Visible = False
                    txtCustomizedReportID.Visible = True
                Else
                    tdbcReportID.Visible = True
                    txtReportName.Visible = True
                    txtCustomizedReportID.Visible = False
                End If
                tdbcSelection01.Text = .Item("Selection01").ToString
                tdbcSelection02.Text = .Item("Selection02").ToString
                tdbcSelection03.Text = .Item("Selection03").ToString
                tdbcSelection04.Text = .Item("Selection04").ToString
                tdbcSelection05.Text = .Item("Selection05").ToString

                tdbcLevel01.Text = .Item("Level01").ToString
                tdbcLevel02.Text = .Item("Level02").ToString
                tdbcLevel03.Text = .Item("Level03").ToString
                tdbcLevel04.Text = .Item("Level04").ToString
                tdbcLevel05.Text = .Item("Level05").ToString
            End With
        End If
        LoadtdbcReportID()
        tdbcReportID.Text = _ReportID
    End Sub

    Private Sub LoadDefault()
        tdbcUnit.Text = "1"
        tdbcBracketPlaces.Text = "None"
        tdbcDecimalNegative.Text = "1"
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_muc_bao_cao_nghiep_vu_tac_dong") & " - D02F6004" & UnicodeCaption(gbUnicode) 'Danh móc bÀo cÀo nghiÖp vó tÀc ¢èng - D02F6004
        '================================================================ 
        lblReportCode.Text = rl3("_Ma_bao_cao") 'Mã báo cáo
        lblReportName1.Text = rl3("Ten_bao_cao") & " 1" 'Tên báo cáo 1
        lblReportName2.Text = rl3("Ten_bao_cao") & " 2" 'Tên báo cáo 2
        lblReportID.Text = rl3("Dang_bao_cao") 'Dạng báo cáo
        lblUnit.Text = rl3("Don_vi_tinh") 'Đơn vị tính
        lblBarcketPlaces.Text = rl3("Hien_thi_so_am") 'Hiển thị số âm
        lblDecimalNegative.Text = rl3("Dinh_dang_so_le") 'Định dạng số lẻ
        lblSelection01.Text = rl3("Tieu_thuc") & " 1" 'Tiêu thức 1
        lblSelection02.Text = rl3("Tieu_thuc") & " 2" 'Tiêu thức 2
        lblSelection03.Text = rl3("Tieu_thuc") & " 3" 'Tiêu thức 3
        lblSelection04.Text = rl3("Tieu_thuc") & " 4" 'Tiêu thức 4
        lblSelection05.Text = rl3("Tieu_thuc") & " 5" 'Tiêu thức 5
        lblLevel01.Text = rl3("Nhom") & " 1" 'Nhóm 1
        lblLevel02.Text = rl3("Nhom") & " 2" 'Nhóm 2
        lblLevel03.Text = rl3("Nhom") & " 3" 'Nhóm 3
        lblLevel04.Text = rl3("Nhom") & " 4" 'Nhóm 4
        lblLevel05.Text = rl3("Nhom") & " 5" 'Nhóm 5
        '================================================================ 
        btnSave.Text = rl3("_Luu") '&Lưu
        btnNext.Text = rl3("_Nhap_tiep") 'Nhập &tiếp
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        chkDisabled.Text = rl3("Khong_su_dung") 'Không sử dụng
        chkCustomized.Text = rl3("Dac_thu") 'Đặc thù
        '================================================================ 
        grpSelection.Text = "1. " & rl3("Chon_tieu_thuc") 'Chọn tiêu thức
        grpGroup.Text = "2. " & rl3("Chon_nhom") 'Chọn nhóm
        grpCustomized.Text = "1. " & rl3("Chon_dang_bao_cao") 'Chọn dạng báo cáo
        '================================================================ 
        TabInfo.Text = "1. " & rl3("Thong_tin_chung") '1. Thông tin chung
        TabSelection.Text = "2. " & rl3("Tieu_thuc_va_nhom") '2. Tiêu thức và nhóm
        '================================================================ 
        tdbcReportID.Columns("ReportID").Caption = rl3("Ma") 'Mã 
        tdbcReportID.Columns("ReportName").Caption = rl3("Ten") 'Tên
        tdbcSelection01.Columns("Code").Caption = rl3("Ma") 'Mã 
        tdbcSelection01.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcSelection02.Columns("Code").Caption = rl3("Ma") 'Mã 
        tdbcSelection02.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcSelection03.Columns("Code").Caption = rl3("Ma") 'Mã 
        tdbcSelection03.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcSelection04.Columns("Code").Caption = rl3("Ma") 'Mã 
        tdbcSelection04.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcSelection05.Columns("Code").Caption = rl3("Ma") 'Mã 
        tdbcSelection05.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcLevel01.Columns("Code").Caption = rl3("Ma") 'Mã 
        tdbcLevel01.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcLevel02.Columns("Code").Caption = rl3("Ma") 'Mã 
        tdbcLevel02.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcLevel03.Columns("Code").Caption = rl3("Ma") 'Mã 
        tdbcLevel03.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcLevel04.Columns("Code").Caption = rl3("Ma") 'Mã 
        tdbcLevel04.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbcLevel05.Columns("Code").Caption = rl3("Ma") 'Mã 
        tdbcLevel05.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải

        btnDefineData.Text = rl3("Dinh_nghia__du_lieu") 'Định nghĩa &dữ liệu
    End Sub

    Private Sub D02F6004_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If

        If e.Alt Then
            If e.KeyCode = Keys.NumPad1 Or e.KeyCode = Keys.D1 Then
                Application.DoEvents()
                tabMain.SelectedTab = TabInfo
                If txtReportCode.Enabled Then
                    txtReportCode.Focus()
                Else
                    txtReportName1.Focus()
                End If
                Application.DoEvents()
            End If
            If e.KeyCode = Keys.NumPad2 Or e.KeyCode = Keys.D2 Then
                Application.DoEvents()
                tabMain.SelectedTab = TabSelection
                tdbcSelection01.Focus()
                Application.DoEvents()
            End If
        End If
    End Sub

    Private Sub D02F6004_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	If bLoadFormState = False Then FormState = _formState
        Me.Cursor = Cursors.WaitCursor
        Loadlanguage()
        InputbyUnicode(Me, gbUnicode)
        SetBackColorObligatory()
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub SetBackColorObligatory()
        txtReportCode.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcReportID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcLevel01.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        txtCustomizedReportID.BackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

#End Region

#Region "Button Click"

    Private Function AllowSave() As Boolean
        If txtReportCode.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rl3("_Ma_bao_cao"))
            tabMain.SelectedTab = TabInfo
            txtReportCode.Focus()
            Return False
        End If
        If chkCustomized.Checked Then
            If txtCustomizedReportID.Text.Trim = "" Then
                D99C0008.MsgNotYetEnter(rl3("Dang_bao_cao"))
                tabMain.SelectedTab = TabInfo
                txtCustomizedReportID.Focus()
                Return False
            End If
        Else
            If tdbcReportID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rl3("Dang_bao_cao"))
                tabMain.SelectedTab = TabInfo
                tdbcReportID.Focus()
                Return False
            End If
        End If
        If tdbcLevel01.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Nhom") & " 1")
            tabMain.SelectedTab = TabSelection
            tdbcLevel01.Focus()
            Return False
        End If
        If tdbcSelection01.Text <> "" Then
            If tdbcSelection02.Text <> "" And tdbcSelection02.Enabled = True Then
                If tdbcSelection01.Text = tdbcSelection02.Text Then
                    D99C0008.MsgL3(rl3("Tieu_thuc_nay_da_bi_trung"))
                    tabMain.SelectedTab = TabSelection
                    tdbcSelection02.Focus()
                    Return False
                End If
                If tdbcSelection03.Text <> "" And tdbcSelection03.Enabled = True Then
                    If tdbcSelection02.Text = tdbcSelection03.Text Then
                        D99C0008.MsgL3(rl3("Tieu_thuc_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcSelection03.Focus()
                        Return False
                    End If
                End If
                If tdbcSelection04.Text <> "" And tdbcSelection04.Enabled = True Then
                    If tdbcSelection02.Text = tdbcSelection04.Text Then
                        D99C0008.MsgL3(rl3("Tieu_thuc_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcSelection04.Focus()
                        Return False
                    End If
                End If
                If tdbcSelection05.Text <> "" And tdbcSelection05.Enabled = True Then
                    If tdbcSelection02.Text = tdbcSelection05.Text Then
                        D99C0008.MsgL3(rl3("Tieu_thuc_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcSelection05.Focus()
                        Return False
                    End If
                End If
            End If
            If tdbcSelection03.Text <> "" And tdbcSelection03.Enabled = True Then
                If tdbcSelection01.Text = tdbcSelection03.Text Then
                    D99C0008.MsgL3(rl3("Tieu_thuc_nay_da_bi_trung"))
                    tabMain.SelectedTab = TabSelection
                    tdbcSelection03.Focus()
                    Return False
                End If
                If tdbcSelection04.Text <> "" And tdbcSelection04.Enabled = True Then
                    If tdbcSelection03.Text = tdbcSelection04.Text Then
                        D99C0008.MsgL3(rl3("Tieu_thuc_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcSelection04.Focus()
                        Return False
                    End If
                End If
                If tdbcSelection05.Text <> "" And tdbcSelection05.Enabled = True Then
                    If tdbcSelection03.Text = tdbcSelection05.Text Then
                        D99C0008.MsgL3(rl3("Tieu_thuc_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcSelection05.Focus()
                        Return False
                    End If
                End If
            End If
            If tdbcSelection04.Text <> "" And tdbcSelection04.Enabled = True Then
                If tdbcSelection04.Text = tdbcSelection01.Text Then
                    D99C0008.MsgL3(rl3("Tieu_thuc_nay_da_bi_trung"))
                    tabMain.SelectedTab = TabSelection
                    tdbcSelection04.Focus()
                    Return False
                End If
                If tdbcSelection05.Text <> "" And tdbcSelection05.Enabled = True Then
                    If tdbcSelection04.Text = tdbcSelection05.Text Then
                        D99C0008.MsgL3(rl3("Tieu_thuc_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcSelection05.Focus()
                        Return False
                    End If
                End If
            End If
            If tdbcSelection05.Text <> "" And tdbcSelection05.Enabled = True Then
                If tdbcSelection05.Text = tdbcSelection01.Text Then
                    D99C0008.MsgL3(rl3("Tieu_thuc_nay_da_bi_trung"))
                    tabMain.SelectedTab = TabSelection
                    tdbcSelection05.Focus()
                    Return False
                End If
            End If
        End If

        If tdbcLevel01.Text <> "" Then
            If tdbcLevel02.Text <> "" And tdbcLevel02.Enabled = True Then
                If tdbcLevel01.Text = tdbcLevel02.Text Then
                    D99C0008.MsgL3(rl3("Nhom_nay_da_bi_trung"))
                    tabMain.SelectedTab = TabSelection
                    tdbcLevel02.Focus()
                    Return False
                End If
                If tdbcLevel03.Text <> "" And tdbcLevel03.Enabled = True Then
                    If tdbcLevel02.Text = tdbcLevel03.Text Then
                        D99C0008.MsgL3(rl3("Nhom_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcLevel03.Focus()
                        Return False
                    End If
                End If
                If tdbcLevel04.Text <> "" And tdbcLevel04.Enabled = True Then
                    If tdbcLevel02.Text = tdbcLevel04.Text Then
                        D99C0008.MsgL3(rl3("Nhom_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcLevel04.Focus()
                        Return False
                    End If
                End If
                If tdbcLevel05.Text <> "" And tdbcLevel05.Enabled = True Then
                    If tdbcLevel02.Text = tdbcLevel05.Text Then
                        D99C0008.MsgL3(rl3("Nhom_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcLevel05.Focus()
                        Return False
                    End If
                End If
            End If
            If tdbcLevel03.Text <> "" And tdbcLevel03.Enabled = True Then
                If tdbcLevel01.Text = tdbcLevel03.Text Then
                    D99C0008.MsgL3(rl3("Nhom_nay_da_bi_trung"))
                    tabMain.SelectedTab = TabSelection
                    tdbcLevel03.Focus()
                    Return False
                End If
                If tdbcLevel04.Text <> "" And tdbcLevel04.Enabled = True Then
                    If tdbcLevel03.Text = tdbcLevel04.Text Then
                        D99C0008.MsgL3(rl3("Nhom_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcLevel04.Focus()
                        Return False
                    End If
                End If
                If tdbcLevel05.Text <> "" And tdbcLevel05.Enabled = True Then
                    If tdbcLevel03.Text = tdbcLevel05.Text Then
                        D99C0008.MsgL3(rl3("Nhom_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcLevel05.Focus()
                        Return False
                    End If
                End If
            End If
            If tdbcLevel04.Text <> "" And tdbcLevel04.Enabled = True Then
                If tdbcLevel04.Text = tdbcLevel01.Text Then
                    D99C0008.MsgL3(rl3("Nhom_nay_da_bi_trung"))
                    tabMain.SelectedTab = TabSelection
                    tdbcLevel04.Focus()
                    Return False
                End If
                If tdbcLevel05.Text <> "" And tdbcLevel05.Enabled = True Then
                    If tdbcLevel04.Text = tdbcLevel05.Text Then
                        D99C0008.MsgL3(rl3("Nhom_nay_da_bi_trung"))
                        tabMain.SelectedTab = TabSelection
                        tdbcLevel05.Focus()
                        Return False
                    End If
                End If
            End If
            If tdbcLevel05.Text <> "" And tdbcLevel05.Enabled = True Then
                If tdbcLevel05.Text = tdbcLevel01.Text Then
                    D99C0008.MsgL3(rl3("Nhom_nay_da_bi_trung"))
                    tabMain.SelectedTab = TabSelection
                    tdbcLevel05.Focus()
                    Return False
                End If
            End If
        End If

        If _FormState = EnumFormState.FormAdd Then
            If IsExistKey("D02T3040", "ReportCode", txtReportCode.Text) = True Then
                D99C0008.MsgDuplicatePKey()
                tabMain.SelectedTab = TabInfo
                txtReportCode.Focus()
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub chkCustomized_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkCustomized.Click
        If chkCustomized.Checked Then
            tdbcReportID.Visible = False
            txtReportName.Visible = False
            txtCustomizedReportID.Visible = True
            txtCustomizedReportID.Text = ""
            txtCustomizedReportID.Focus()
        Else
            tdbcReportID.Visible = True
            txtReportName.Visible = True
            txtCustomizedReportID.Visible = False
            tdbcReportID.Text = ""
            tdbcReportID.Focus()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not AllowSave() Then Exit Sub

        'Kiểm tra Ngày phiếu có phù hợp với kỳ kế toán hiện tại không (gọi hàm CheckVoucherDateInPeriod)

        btnSave.Enabled = False
        btnClose.Enabled = False

        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        Select Case _FormState
            Case EnumFormState.FormAdd
                sSQL.Append(SQLInsertD02T3040)

                'Lưu LastKey của Số phiếu xuống Database (gọi hàm CreateIGEVoucherNo bật cờ True)
                'Kiểm tra trùng Số phiếu (gọi hàm CheckDuplicateVoucherNo)
                'Nếu tra trùng Số phiếu thì bật
                'btnSave.Enabled = True
                'btnClose.Enabled = True

            Case EnumFormState.FormEdit
                sSQL.Append(SQLUpdateD02T3040)
        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            _bSaved = True
            btnClose.Enabled = True
            Select Case _FormState
                Case EnumFormState.FormAdd
                    _ReportCode = txtReportCode.Text
                    btnNext.Enabled = True
                    btnNext.Focus()
                Case EnumFormState.FormEdit
                    btnSave.Enabled = True
                    btnClose.Focus()
            End Select
        Else
            SaveNotOK()
            _bSaved = False
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        _ReportCode = txtReportCode.Text
        Me.Close()
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click

        btnSave.Enabled = True
        btnNext.Enabled = False
        _FormState = EnumFormState.FormAdd
        txtReportCode.Text = ""
        txtReportName1.Text = ""
        txtReportName2.Text = ""
        chkDisabled.Checked = False
        tdbcReportID.Text = ""
        txtCustomizedReportID.Text = ""

        tdbcSelection01.SelectedValue = ""
        tdbcSelection02.SelectedValue = ""
        tdbcSelection03.SelectedValue = ""
        tdbcSelection04.SelectedValue = ""
        tdbcSelection05.SelectedValue = ""

        tdbcLevel01.SelectedValue = ""
        tdbcLevel02.SelectedValue = ""
        tdbcLevel03.SelectedValue = ""
        tdbcLevel04.SelectedValue = ""
        tdbcLevel05.SelectedValue = ""
        '-----------------------------
        tabMain.SelectedTab = TabInfo
        txtReportCode.Focus()
        ResetLevel(1)
        ResetSelection(1)
    End Sub

#End Region

#Region "Events tdbcReportID with txtReportName"

    Private Sub tdbcReportID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcReportID.SelectedValueChanged
        If tdbcReportID.Text <> "" Then
            txtReportName.Text = tdbcReportID.Columns(1).Value.ToString
        Else
            txtReportName.Text = ""
        End If
    End Sub

    Private Sub tdbcReportID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcReportID.LostFocus
        If tdbcReportID.FindStringExact(tdbcReportID.Text) = -1 Then
            tdbcReportID.Text = ""
            txtReportName.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcUnit"

    Private Sub tdbcUnit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcUnit.LostFocus
        If tdbcUnit.FindStringExact(tdbcUnit.Text) = -1 Then tdbcUnit.Text = ""
    End Sub

#End Region

#Region "Events tdbcDecimalNegative"

    Private Sub tdbcDecimalNegative_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcDecimalNegative.LostFocus
        If tdbcDecimalNegative.FindStringExact(tdbcDecimalNegative.Text) = -1 Then tdbcDecimalNegative.Text = ""
    End Sub

#End Region

#Region "Events tdbcBracketPlaces"

    Private Sub tdbcBracketPlaces_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcBracketPlaces.SelectedValueChanged
        If tdbcBracketPlaces.Text = "-123" Then
            lblexp.Text = rl3("Vi_du") & ": - 55,555.555   ---->   - 55,555.555"
        ElseIf tdbcBracketPlaces.Text = "(123)" Then
            lblexp.Text = rl3("Vi_du") & ": - 55,555.555   ---->   (55,555.555)"
        Else
            lblexp.Text = rl3("Vi_du") & ": - 55,555.555   ---->   55,555.555"
        End If
    End Sub

    Private Sub tdbcBracketPlaces_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcBracketPlaces.LostFocus
        If tdbcBracketPlaces.FindStringExact(tdbcBracketPlaces.Text) = -1 Then tdbcBracketPlaces.Text = ""
    End Sub

#End Region

#Region "Events tdbcSelection01 with txtSelection01Name"

    Private Sub tdbcSelection01_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSelection01.LostFocus
        If tdbcSelection01.FindStringExact(tdbcSelection01.Text) = -1 Then
            tdbcSelection01.Text = ""
            txtSelection01Name.Text = ""
        End If
        ResetSelection(1)
    End Sub

    Private Sub tdbcSelection01_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcSelection01.SelectedValueChanged
        txtSelection01Name.Text = tdbcSelection01.Columns(1).Value.ToString
        EnabledSelection(1)
    End Sub

#End Region

#Region "Events tdbcSelection02 with txtSelection02Name"

    Private Sub tdbcSelection02_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSelection02.LostFocus
        If tdbcSelection02.FindStringExact(tdbcSelection02.Text) = -1 Then
            tdbcSelection02.Text = ""
            txtSelection02Name.Text = ""
        End If
        ResetSelection(2)
    End Sub

    Private Sub tdbcSelection02_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcSelection02.SelectedValueChanged
        txtSelection02Name.Text = tdbcSelection02.Columns(1).Value.ToString
        EnabledSelection(2)
    End Sub

#End Region

#Region "Events tdbcSelection03 with txtSelection03Name"

    Private Sub tdbcSelection03_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSelection03.LostFocus
        If tdbcSelection03.FindStringExact(tdbcSelection03.Text) = -1 Then
            tdbcSelection03.Text = ""
            txtSelection03Name.Text = ""
        End If
        ResetSelection(3)
    End Sub

    Private Sub tdbcSelection03_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcSelection03.SelectedValueChanged
        txtSelection03Name.Text = tdbcSelection03.Columns(1).Value.ToString
        EnabledSelection(3)
    End Sub

#End Region

#Region "Events tdbcSelection04 with txtSelection04Name"

    Private Sub tdbcSelection04_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSelection04.LostFocus
        If tdbcSelection04.FindStringExact(tdbcSelection04.Text) = -1 Then
            tdbcSelection04.Text = ""
            txtSelection04Name.Text = ""
        End If
        ResetSelection(4)
    End Sub

    Private Sub tdbcSelection04_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcSelection04.SelectedValueChanged
        txtSelection04Name.Text = tdbcSelection04.Columns(1).Value.ToString
        EnabledSelection(4)
    End Sub

#End Region

#Region "Events tdbcSelection05 with txtSelection05Name"

    Private Sub tdbcSelection05_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcSelection05.SelectedValueChanged
        txtSelection05Name.Text = tdbcSelection05.Columns(1).Value.ToString
    End Sub

    Private Sub tdbcSelection05_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcSelection05.LostFocus
        If tdbcSelection05.FindStringExact(tdbcSelection05.Text) = -1 Then
            tdbcSelection05.Text = ""
            txtSelection05Name.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcLevel01 with txtLevel01Name"

    Private Sub tdbcLevel01_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcLevel01.LostFocus
        If tdbcLevel01.FindStringExact(tdbcLevel01.Text) = -1 Then
            tdbcLevel01.Text = ""
            txtLevel01Name.Text = ""
        End If
        EnabledLevel(1)
    End Sub

    Private Sub tdbcLevel01_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcLevel01.SelectedValueChanged
        txtLevel01Name.Text = tdbcLevel01.Columns(1).Value.ToString
        ResetLevel(1)
    End Sub
#End Region

#Region "Events tdbcLevel02 with txtLevel02Name"

    Private Sub tdbcLevel02_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcLevel02.LostFocus
        If tdbcLevel02.FindStringExact(tdbcLevel02.Text) = -1 Then
            tdbcLevel02.Text = ""
            txtLevel02Name.Text = ""
        End If
        EnabledLevel(2)
    End Sub

    Private Sub tdbcLevel02_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcLevel02.SelectedValueChanged
        txtLevel02Name.Text = tdbcLevel02.Columns(1).Value.ToString
        ResetLevel(2)
    End Sub
#End Region

#Region "Events tdbcLevel03 with txtLevel03Name"

    Private Sub tdbcLevel03_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcLevel03.LostFocus
        If tdbcLevel03.FindStringExact(tdbcLevel03.Text) = -1 Then
            tdbcLevel03.Text = ""
            txtLevel03Name.Text = ""
        End If
        EnabledLevel(3)
    End Sub

    Private Sub tdbcLevel03_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcLevel03.SelectedValueChanged
        txtLevel03Name.Text = tdbcLevel03.Columns(1).Value.ToString
        ResetLevel(3)
    End Sub

#End Region

#Region "Events tdbcLevel04 with txtLevel04Name"

    Private Sub tdbcLevel04_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcLevel04.LostFocus
        If tdbcLevel04.FindStringExact(tdbcLevel04.Text) = -1 Then
            tdbcLevel04.Text = ""
            txtLevel04Name.Text = ""
        End If
        EnabledLevel(4)
    End Sub

    Private Sub tdbcLevel04_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcLevel04.SelectedValueChanged
        txtLevel04Name.Text = tdbcLevel04.Columns(1).Value.ToString
        ResetLevel(4)
    End Sub
#End Region

#Region "Events tdbcLevel05 with txtLevel05Name"

    Private Sub tdbcLevel05_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcLevel05.SelectedValueChanged
        txtLevel05Name.Text = tdbcLevel05.Columns(1).Value.ToString
    End Sub

    Private Sub tdbcLevel05_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcLevel05.LostFocus
        If tdbcLevel05.FindStringExact(tdbcLevel05.Text) = -1 Then
            tdbcLevel05.Text = ""
            txtLevel05Name.Text = ""
        End If
    End Sub

#End Region

    Private Sub ResetSelection(ByVal NoSelection As Integer)
        Select Case NoSelection
            Case 1
                If tdbcSelection01.Text = "" Then
                    tdbcSelection02.Text = ""
                    tdbcSelection03.Text = ""
                    tdbcSelection04.Text = ""
                    tdbcSelection05.Text = ""
                    txtSelection01Name.Text = ""
                    txtSelection02Name.Text = ""
                    txtSelection03Name.Text = ""
                    txtSelection04Name.Text = ""
                    txtSelection05Name.Text = ""
                    tdbcSelection02.Enabled = False
                    tdbcSelection03.Enabled = False
                    tdbcSelection04.Enabled = False
                    tdbcSelection05.Enabled = False
                Else
                    tdbcSelection02.Enabled = True
                    If tdbcSelection02.Text = "" Then
                        tdbcSelection03.Text = ""
                        tdbcSelection04.Text = ""
                        tdbcSelection05.Text = ""
                        txtSelection02Name.Text = ""
                        txtSelection03Name.Text = ""
                        txtSelection04Name.Text = ""
                        txtSelection05Name.Text = ""
                        tdbcSelection03.Enabled = False
                        tdbcSelection04.Enabled = False
                        tdbcSelection05.Enabled = False
                    Else
                        tdbcSelection03.Enabled = True
                        If tdbcSelection03.Text = "" Then
                            tdbcSelection04.Text = ""
                            tdbcSelection05.Text = ""
                            txtSelection03Name.Text = ""
                            txtSelection04Name.Text = ""
                            txtSelection05Name.Text = ""
                            tdbcSelection04.Enabled = False
                            tdbcSelection05.Enabled = False
                        Else
                            tdbcSelection04.Enabled = True
                            If tdbcSelection04.Text = "" Then
                                tdbcSelection05.Text = ""
                                txtSelection04Name.Text = ""
                                txtSelection05Name.Text = ""
                                tdbcSelection05.Enabled = False
                            Else
                                tdbcSelection05.Enabled = True
                            End If
                        End If
                    End If
                End If
            Case 2
                If tdbcSelection02.Text = "" Then
                    tdbcSelection03.Text = ""
                    tdbcSelection04.Text = ""
                    tdbcSelection05.Text = ""
                    txtSelection02Name.Text = ""
                    txtSelection03Name.Text = ""
                    txtSelection04Name.Text = ""
                    txtSelection05Name.Text = ""
                    tdbcSelection03.Enabled = False
                    tdbcSelection04.Enabled = False
                    tdbcSelection05.Enabled = False
                Else
                    tdbcSelection03.Enabled = True
                    If tdbcSelection03.Text = "" Then
                        tdbcSelection04.Text = ""
                        tdbcSelection05.Text = ""
                        txtSelection03Name.Text = ""
                        txtSelection04Name.Text = ""
                        txtSelection05Name.Text = ""
                        tdbcSelection04.Enabled = False
                        tdbcSelection05.Enabled = False
                    Else
                        tdbcSelection04.Enabled = True
                        If tdbcSelection04.Text = "" Then
                            tdbcSelection05.Text = ""
                            txtSelection04Name.Text = ""
                            txtSelection05Name.Text = ""
                            tdbcSelection05.Enabled = False
                        Else
                            tdbcSelection05.Enabled = True
                        End If
                    End If
                End If
            Case 3
                If tdbcSelection03.Text = "" Then
                    tdbcSelection04.Text = ""
                    tdbcSelection05.Text = ""
                    txtSelection03Name.Text = ""
                    txtSelection04Name.Text = ""
                    txtSelection05Name.Text = ""
                    tdbcSelection04.Enabled = False
                    tdbcSelection05.Enabled = False
                Else
                    tdbcSelection04.Enabled = True
                    If tdbcSelection04.Text = "" Then
                        tdbcSelection05.Text = ""
                        txtSelection04Name.Text = ""
                        txtSelection05Name.Text = ""
                        tdbcSelection05.Enabled = False
                    Else
                    End If
                End If
            Case 4
                If tdbcSelection04.Text = "" Then
                    tdbcSelection05.Text = ""
                    txtSelection04Name.Text = ""
                    txtSelection05Name.Text = ""
                    tdbcSelection05.Enabled = False
                Else
                    tdbcSelection05.Enabled = True
                End If
        End Select
    End Sub

    Private Sub EnabledSelection(ByVal ID As Integer)
        Select Case ID
            Case 1
                If tdbcSelection01.Text = "" Then
                    tdbcSelection02.Enabled = False
                    tdbcSelection03.Enabled = False
                    tdbcSelection04.Enabled = False
                    tdbcSelection05.Enabled = False
                Else
                    tdbcSelection02.Enabled = True
                End If
            Case 2
                If tdbcSelection02.Text = "" Then
                    tdbcSelection03.Enabled = False
                    tdbcSelection04.Enabled = False
                    tdbcSelection05.Enabled = False
                Else
                    tdbcSelection03.Enabled = True
                End If
            Case 3
                If tdbcSelection03.Text = "" Then
                    tdbcSelection04.Enabled = False
                    tdbcSelection05.Enabled = False
                Else
                    tdbcSelection04.Enabled = True
                End If
            Case 4
                If tdbcSelection04.Text = "" Then
                    tdbcSelection05.Enabled = False
                Else
                    tdbcSelection05.Enabled = True
                End If
        End Select
    End Sub

    Private Sub ResetLevel(ByVal ID As Integer)
        Select Case ID
            Case 1
                If tdbcLevel01.Text = "" Then
                    tdbcLevel02.Text = ""
                    tdbcLevel03.Text = ""
                    tdbcLevel04.Text = ""
                    tdbcLevel05.Text = ""
                    txtLevel01Name.Text = ""
                    txtLevel02Name.Text = ""
                    txtLevel03Name.Text = ""
                    txtLevel04Name.Text = ""
                    txtLevel05Name.Text = ""
                    tdbcLevel02.Enabled = False
                    tdbcLevel03.Enabled = False
                    tdbcLevel04.Enabled = False
                    tdbcLevel05.Enabled = False
                Else
                    tdbcLevel02.Enabled = True
                    If tdbcLevel02.Text = "" Then
                        tdbcLevel03.Text = ""
                        tdbcLevel04.Text = ""
                        tdbcLevel05.Text = ""
                        txtLevel02Name.Text = ""
                        txtLevel03Name.Text = ""
                        txtLevel04Name.Text = ""
                        txtLevel05Name.Text = ""
                        tdbcLevel03.Enabled = False
                        tdbcLevel04.Enabled = False
                        tdbcLevel05.Enabled = False
                    Else
                        tdbcLevel03.Enabled = True
                        If tdbcLevel03.Text = "" Then
                            tdbcLevel04.Text = ""
                            tdbcLevel05.Text = ""
                            txtLevel03Name.Text = ""
                            txtLevel04Name.Text = ""
                            txtLevel05Name.Text = ""
                            tdbcLevel04.Enabled = False
                            tdbcLevel05.Enabled = False
                        Else
                            tdbcLevel04.Enabled = True
                            If tdbcLevel04.Text = "" Then
                                tdbcLevel05.Text = ""
                                txtLevel04Name.Text = ""
                                txtLevel05Name.Text = ""
                                tdbcLevel05.Enabled = False
                            Else
                                tdbcLevel05.Enabled = True
                            End If
                        End If
                    End If
                End If
            Case 2
                If tdbcLevel02.Text = "" Then
                    tdbcLevel03.Text = ""
                    tdbcLevel04.Text = ""
                    tdbcLevel05.Text = ""
                    txtLevel02Name.Text = ""
                    txtLevel03Name.Text = ""
                    txtLevel04Name.Text = ""
                    txtLevel05Name.Text = ""
                    tdbcLevel03.Enabled = False
                    tdbcLevel04.Enabled = False
                    tdbcLevel05.Enabled = False
                Else
                    tdbcLevel03.Enabled = True
                    If tdbcLevel03.Text = "" Then
                        tdbcLevel04.Text = ""
                        tdbcLevel05.Text = ""
                        txtLevel03Name.Text = ""
                        txtLevel04Name.Text = ""
                        txtLevel05Name.Text = ""
                        tdbcLevel04.Enabled = False
                        tdbcLevel05.Enabled = False
                    Else
                        tdbcLevel04.Enabled = True
                        If tdbcLevel04.Text = "" Then
                            tdbcLevel05.Text = ""
                            txtLevel04Name.Text = ""
                            txtLevel05Name.Text = ""
                            tdbcLevel05.Enabled = False
                        Else
                            tdbcLevel05.Enabled = True
                        End If
                    End If
                End If
            Case 3
                If tdbcLevel03.Text = "" Then
                    tdbcLevel04.Text = ""
                    tdbcLevel05.Text = ""
                    txtSelection03Name.Text = ""
                    txtLevel04Name.Text = ""
                    txtLevel05Name.Text = ""
                    tdbcLevel04.Enabled = False
                    tdbcLevel05.Enabled = False
                Else
                    tdbcLevel04.Enabled = True
                    If tdbcLevel04.Text = "" Then
                        tdbcLevel05.Text = ""
                        txtLevel04Name.Text = ""
                        txtLevel05Name.Text = ""
                        tdbcLevel05.Enabled = False
                    Else
                    End If
                End If
            Case 4
                If tdbcLevel04.Text = "" Then
                    tdbcLevel05.Text = ""
                    txtLevel04Name.Text = ""
                    txtLevel05Name.Text = ""
                    tdbcLevel05.Enabled = False
                Else
                    tdbcLevel05.Enabled = True
                End If
        End Select
    End Sub

    Private Sub EnabledLevel(ByVal ID As Integer)
        Select Case ID
            Case 1
                If tdbcLevel01.Text = "" Then
                    tdbcLevel02.Enabled = False
                    tdbcLevel03.Enabled = False
                    tdbcLevel04.Enabled = False
                    tdbcLevel05.Enabled = False
                Else
                    tdbcLevel02.Enabled = True
                End If
            Case 2
                If tdbcLevel02.Text = "" Then
                    tdbcLevel03.Enabled = False
                    tdbcLevel04.Enabled = False
                    tdbcLevel05.Enabled = False
                Else
                    tdbcLevel03.Enabled = True
                End If
            Case 3
                If tdbcLevel03.Text = "" Then
                    tdbcLevel04.Enabled = False
                    tdbcLevel05.Enabled = False
                Else
                    tdbcLevel04.Enabled = True
                End If
            Case 4
                If tdbcLevel04.Text = "" Then
                    tdbcLevel05.Enabled = False
                Else
                    tdbcLevel05.Enabled = True
                End If
        End Select
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T3030
    '# Created User: Thiên Huỳnh
    '# Created Date: 05/01/2009 08:09:59
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T3040() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T3040(")
        sSQL.Append("ReportCode, ReportName1, ReportName1U, ReportName2U, Selection01, Selection02, ")
        sSQL.Append("Selection03, Selection04, Selection05, Level01, Level02, ")
        sSQL.Append("Level03, BracketNegative, DecimalPlaces, AmountFormat, Customized, ")
        sSQL.Append("CreateUserID, CreateDate, LastModifyUserID, LastModifyDate, Disabled, ")
        sSQL.Append("ReportID, Level04, Level05")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(txtReportCode.Text) & COMMA) 'ReportCode [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString("") & COMMA) 'ReportName1, varchar[150], NOT NULL
        sSQL.Append(SQLStringUnicode(txtReportName1.Text, gbUnicode, True) & COMMA) 'ReportName1, varchar[150], NOT NULL
        sSQL.Append(SQLStringUnicode(txtReportName2.Text, gbUnicode, True) & COMMA) 'ReportName2, varchar[150], NULL
        sSQL.Append(SQLString(tdbcSelection01.Text) & COMMA) 'Selection01, varchar[20], NULL
        sSQL.Append(SQLString(IIf(tdbcSelection02.Enabled, tdbcSelection02.Text, "")) & COMMA) 'Selection02, varchar[20], NULL
        sSQL.Append(SQLString(IIf(tdbcSelection03.Enabled, tdbcSelection03.Text, "")) & COMMA) 'Selection03, varchar[20], NULL
        sSQL.Append(SQLString(IIf(tdbcSelection04.Enabled, tdbcSelection04.Text, "")) & COMMA) 'Selection04, varchar[20], NULL
        sSQL.Append(SQLString(IIf(tdbcSelection05.Enabled, tdbcSelection05.Text, "")) & COMMA) 'Selection05, varchar[20], NULL
        sSQL.Append(SQLString(tdbcLevel01.Text) & COMMA) 'Level01, varchar[20], NULL
        sSQL.Append(SQLString(IIf(tdbcLevel02.Enabled, tdbcLevel02.Text, "")) & COMMA) 'Level02, varchar[20], NULL
        sSQL.Append(SQLString(IIf(tdbcLevel02.Enabled, tdbcLevel03.Text, "")) & COMMA) 'Level03, varchar[20], NULL
        sSQL.Append(SQLNumber(tdbcBracketPlaces.Columns("Code").Value.ToString) & COMMA) 'BracketNegative, tinyint, NULL
        sSQL.Append(SQLNumber(tdbcDecimalNegative.Columns("Code").Value.ToString) & COMMA) 'DecimalPlaces, tinyint, NULL
        sSQL.Append(SQLNumber(tdbcUnit.Columns("Code").Value.ToString) & COMMA) 'AmountFormat, tinyint, NULL
        sSQL.Append(SQLNumber(IIf(chkCustomized.Checked, "1", "0").ToString) & COMMA) 'Customized, tinyint, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
        sSQL.Append(SQLNumber(IIf(chkDisabled.Checked, "1", "0").ToString) & COMMA) 'Disabled, tinyint, NOT NULL
        If chkCustomized.Checked Then
            sSQL.Append(SQLString(txtCustomizedReportID.Text) & COMMA) 'ReportID, varchar[20], NULL
        Else
            sSQL.Append(SQLString(tdbcReportID.Text) & COMMA) 'ReportID, varchar[20], NULL
        End If
        sSQL.Append(SQLString(IIf(tdbcLevel04.Enabled, tdbcLevel04.Text, "")) & COMMA) 'Level04, varchar[20], NOT NULL
        sSQL.Append(SQLString(IIf(tdbcLevel05.Enabled, tdbcLevel05.Text, ""))) 'Level05, varchar[20], NOT NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T3040
    '# Created User: Thiên Huỳnh
    '# Created Date: 05/02/2009 04:03:40
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T3040() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T3040 Set ")
        sSQL.Append("ReportName1U = " & SQLStringUnicode(txtReportName1.Text, gbUnicode, True) & COMMA) 'varchar[150], NOT NULL
        sSQL.Append("ReportName2U = " & SQLStringUnicode(txtReportName2.Text, gbUnicode, True) & COMMA) 'varchar[150], NULL
        sSQL.Append("Selection01 = " & SQLString(tdbcSelection01.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("Selection02 = " & SQLString(IIf(tdbcSelection02.Enabled, tdbcSelection02.Text, "")) & COMMA) 'varchar[20], NULL
        sSQL.Append("Selection03 = " & SQLString(IIf(tdbcSelection03.Enabled, tdbcSelection03.Text, "")) & COMMA) 'varchar[20], NULL
        sSQL.Append("Selection04 = " & SQLString(IIf(tdbcSelection04.Enabled, tdbcSelection04.Text, "")) & COMMA) 'varchar[20], NULL
        sSQL.Append("Selection05 = " & SQLString(IIf(tdbcSelection05.Enabled, tdbcSelection05.Text, "")) & COMMA) 'varchar[20], NULL
        sSQL.Append("Level01 = " & SQLString(tdbcLevel01.Text) & COMMA) 'varchar[20], NULL
        sSQL.Append("Level02 = " & SQLString(IIf(tdbcLevel02.Enabled, tdbcLevel02.Text, "")) & COMMA) 'varchar[20], NULL
        sSQL.Append("Level03 = " & SQLString(IIf(tdbcLevel03.Enabled, tdbcLevel03.Text, "")) & COMMA) 'varchar[20], NULL
        sSQL.Append("BracketNegative = " & SQLNumber(tdbcBracketPlaces.Columns("Code").Value.ToString) & COMMA) 'tinyint, NULL
        sSQL.Append("DecimalPlaces = " & SQLNumber(tdbcDecimalNegative.Columns("Code").Value.ToString) & COMMA) 'tinyint, NULL
        sSQL.Append("AmountFormat = " & SQLNumber(tdbcUnit.Columns("Code").Value.ToString) & COMMA) 'tinyint, NULL
        sSQL.Append("Customized = " & SQLNumber(IIf(chkCustomized.Checked, "1", "0").ToString) & COMMA) 'tinyint, NULL
        sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
        sSQL.Append("Disabled = " & SQLNumber(IIf(chkDisabled.Checked, "1", "0").ToString) & COMMA) 'tinyint, NOT NULL
        If chkCustomized.Checked Then
            sSQL.Append("ReportID = " & SQLString(txtCustomizedReportID.Text) & COMMA) 'varchar[20], NULL
        Else
            sSQL.Append("ReportID = " & SQLString(tdbcReportID.Text) & COMMA) 'varchar[20], NULL
        End If
        sSQL.Append("Level04 = " & SQLString(IIf(tdbcLevel04.Enabled, tdbcLevel04.Text, "")) & COMMA) 'varchar[20], NOT NULL
        sSQL.Append("Level05 = " & SQLString(IIf(tdbcLevel05.Enabled, tdbcLevel05.Text, ""))) 'varchar[20], NOT NULL
        sSQL.Append(" Where ")
        sSQL.Append("ReportCode = " & SQLString(txtReportCode.Text))

        Return sSQL
    End Function

    Private Sub btnDefineData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDefineData.Click
        'Dim frm As New D89F9300
        'With frm
        '    .FormName = "D89F9300"
        '    .FormStatus = _FormState
        '    .FormPermission = "D02F6004" 'Màn hình phân quyền
        '    .Key01ID = "D02F6004" 'Tên Form cha
        '    .Key02ID = txtReportCode.Text 'Mã báo cáo
        '    .Key03ID = txtReportName1.Text 'Tiêu đề báo cáo
        '    .Key04ID = ""
        '    .Key05ID = ""
        '    .ShowDialog()
        '    .Dispose()
        'End With

        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "FormIDPermission", "D02F6004")
        SetProperties(arrPro, "LoadStatus", _FormState)
        SetProperties(arrPro, "FormID", "D02F6004")
        SetProperties(arrPro, "ReportID", txtReportCode.Text)
        SetProperties(arrPro, "Title", txtReportName1.Text)
        SetProperties(arrPro, "Mode", "")
        CallFormShow(Me, "D89D0240", "D89F9300", arrPro)
    End Sub
End Class