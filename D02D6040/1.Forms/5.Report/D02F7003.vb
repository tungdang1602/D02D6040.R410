'#-------------------------------------------------------------------------------------
'# Created Date: 02/01/2009 10:27:39 AM
'# Created User: Thiên Huỳnh
'# Modify Date: 02/01/2009 10:27:39 AM
'# Modify User: Thiên Huỳnh
'#-------------------------------------------------------------------------------------
Public Class D02F7003
	Dim dtCaptionCols As DataTable

#Region "Const of tdbg"
    Private Const COL_ReportCode As Integer = 0       ' Mã báo cáo
    Private Const COL_ReportName1 As Integer = 1      ' Tên báo cáo
    Private Const COL_ReportID As Integer = 2         ' Dạng báo cáo
    Private Const COL_Disabled As Integer = 3         ' Không sử dụng
    Private Const COL_CreateUserID As Integer = 4     ' CreateUserID
    Private Const COL_CreateDate As Integer = 5       ' CreateDate
    Private Const COL_LastModifyUserID As Integer = 6 ' LastModifyUserID
    Private Const COL_LastModifyDate As Integer = 7   ' LastModifyDate
#End Region

    Private dtGrid As DataTable
    Dim bRefreshFilter As Boolean
    Dim sFilter As New System.Text.StringBuilder()

#Region "Form Load"

    Private Sub D02F5003_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	LoadInfoGeneral()
        Loadlanguage()
        ResetColorGrid(tdbg, 0, tdbg.Splits.ColCount - 1)
        LoadTDBGrid()
        SetShortcutPopupMenu(Me, tbrTableToolStrip, ContextMenuStrip1)
        SetResolutionForm(Me, ContextMenuStrip1)
    End Sub

    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        Dim sSQL As String = ""
        sSQL = "Select Distinct ReportCode, ReportName1" & UnicodeJoin(gbUnicode) & " as ReportName1,ReportID, Disabled, CreateUserID, CreateDate, LastModifyUserID, LastModifyDate From D02T3110  WITH (NOLOCK) Order By ReportCode"
        dtGrid = ReturnDataTable(sSQL)
        'Cách mới theo chuẩn: Tìm kiếm và Liệt kê tất cả luôn luôn sáng Khi(dt.Rows.Count > 0)
        gbEnabledUseFind = dtGrid.Rows.Count > 0
        If FlagAdd Then
            ' Thêm mới thì gán sFind ="" và gán FilterText =’’
            ResetFilter(tdbg, sFilter, bRefreshFilter)
            sFind = ""
        End If
        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()
        If sKey <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select("ReportCode = " & SQLString(sKey), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt1.Rows.IndexOf(dr(0)) 'dùng tdbg.Bookmark có thể không đúng
            If Not tdbg.Focused Then tdbg.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
        End If
    End Sub

    Private Sub ReLoadTDBGrid()
        Dim strFind As String = sFind
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString

        If Not chkShowDisabled.Checked Then
            If strFind <> "" Then strFind &= " And "
            strFind &= "Disabled =0"
        End If
        dtGrid.DefaultView.RowFilter = strFind
        '  LoadGridFind(tdbg, dtGrid, strFind)'gây lỗi không nhập được ký tự thứ 2 trên filter
        ' Nếu lưới có Group thì bổ sung thêm 2 đoạn lệnh sau:
        tdbg.WrapCellPointer = tdbg.RowCount > 0
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        CheckMenu(Me.Name, tbrTableToolStrip, tdbg.RowCount, gbEnabledUseFind, False, ContextMenuStrip1)
        FooterTotalGrid(tdbg, COL_ReportCode)
        mnsExportDataScript.Enabled = tdbg.RowCount > 0
        tsmExportDataScript.Enabled = mnsExportDataScript.Enabled
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_sach_bao_cao_phan_tich_tai_san") & " - D02F7003" & UnicodeCaption(gbUnicode) 'Danh sÀch bÀo cÀo ph¡n tÛch tªi s¶n - D02F7003
        '================================================================ 
        tdbg.Columns("ReportCode").Caption = rl3("_Ma_bao_cao") 'Mã báo cáo
        tdbg.Columns("ReportName1").Caption = rl3("Ten_bao_cao") 'Tên báo cáo
        tdbg.Columns("ReportID").Caption = rl3("Dang_bao_cao") 'Dạng báo cáo
        tdbg.Columns("Disabled").Caption = rl3("KSD") 'KSD
        '================================================================ 
        chkShowDisabled.Text = rl3("Hien_thi_danh_sach_khong_su_dung") 'Hiển thị danh sÀch không sử dụng
        ''Them ngay 20/2/2013 theo ID 54356 của Bảo Trân bởi Văn Vinh
        'tsmExportDataScript.Text = rl3("MSG000051") 'Xuất dữ liệu thiết &lập
        'mnsExportDataScript.Text = tsmExportDataScript.Text
    End Sub

#End Region

#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
	Dim gbEnabledUseFind As Boolean = False
    'Cần sửa Tìm kiếm như sau:
	'Bỏ sự kiện Finder_FindClick.
	'Sửa tham số Me.Name -> Me
	'Phải tạo biến properties có tên chính xác strNewFind và strNewServer
	'Sửa gdtCaptionExcel thành dtCaptionCols: biến toàn cục trong form
	'Nếu có F12 dùng D09U1111 thì Sửa dtCaptionCols thành ResetTableByGrid(usrOption, dtCaptionCols.DefaultView.ToTable)
    Private sFind As String = ""
	Public WriteOnly Property strNewFind() As String
		Set(ByVal Value As String)
			sFind = Value
			ReLoadTDBGrid()'Làm giống sự kiện Finder_FindClick. Ví dụ đối với form Báo cáo thường gọi btnPrint_Click(Nothing, Nothing): sFind = "
		End Set
	End Property

    'Dim dtCaptionCols As DataTable

    Private Sub tsbFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbFind.Click, tsmFind.Click, mnsFind.Click
        'Dim sSQL As String = ""
        'gbEnabledUseFind = True
        'sSQL = "Select * From D02V1234 "
        'sSQL &= "Where FormID = " & SQLString(Me.Name) & "And Language = " & SQLString(gsLanguage)
        'ShowFindDialogClient(Finder, sSQL)
        gbEnabledUseFind = True
        '*****************************************
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        'Những cột bắt buộc nhập
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, , , , gbUnicode)
        'Tạo tableCaption: đưa tất cả các cột trên lưới có Visible = True vào table 
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If

        ShowFindDialogClient(Finder, dtCaptionCols, Me, "0", gbUnicode)
        '*****************************************

    End Sub

    'Private Sub Finder_FindClick(ByVal ResultWhereClause As Object) Handles Finder.FindClick
    '    If ResultWhereClause Is Nothing Then Exit Sub
    '    sFind = ResultWhereClause.ToString()
    '    ReLoadTDBGrid()
    'End Sub

    Private Sub tsbListAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, tsmListAll.Click, mnsListAll.Click
        sFind = ""
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        ReLoadTDBGrid()
    End Sub

#End Region

#Region "C1Context Menu"

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        If D99C0008.MsgAskDelete = Windows.Forms.DialogResult.No Then Exit Sub
        'If Not AllowDelete() Then Exit Sub
        Dim sSQL As String
        sSQL = "Delete D02T3110 Where ReportCode = " & SQLString(tdbg.Columns(COL_ReportCode).Text)
        Dim bRunSQL As Boolean = ExecuteSQL(sSQL)
        If bRunSQL Then
            DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
            ResetGrid()
            DeleteOK()
        Else
            DeleteNotOK()
        End If
    End Sub

    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, tsmAdd.Click, mnsAdd.Click
        Dim f As New D02F7004
        With f
            .m_ReportCode = ""
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            .Dispose()
            If .bSaved Then LoadTDBGrid(True, .m_ReportCode)
        End With
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, tsmEdit.Click, mnsEdit.Click
        If tdbg.RowCount <= 0 Then Exit Sub
        Dim f As New D02F7004
        With f
            .m_ReportCode = tdbg.Columns(COL_ReportCode).Text
            .FormState = EnumFormState.FormEdit
            .ShowDialog()
            .Dispose()
            If .bSaved Then LoadTDBGrid(False, .m_ReportCode)
        End With
    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        If tdbg.RowCount <= 0 Then Exit Sub
        Dim f As New D02F7004
        With f
            .m_ReportCode = tdbg.Columns(COL_ReportCode).Text
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, tsmSysInfo.Click, mnsSysInfo.Click
        ShowSysInfoDialog(Me,tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub

#End Region

    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            If bRefreshFilter Then Exit Sub 'set FilterText ="" thì thoát
            'Filter the data 
            FilterChangeGrid(tdbg, sFilter)
            ReLoadTDBGrid()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try

    End Sub

    Private Sub tdbg_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.DoubleClick
        If tdbg.FilterActive Then Exit Sub
        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.KeyCode = Keys.Enter Then tdbg_DoubleClick(Nothing, Nothing)
        HotKeyCtrlVOnGrid(tdbg, e)
    End Sub

    Private Sub chkShowDisabled_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowDisabled.CheckedChanged
        If dtGrid Is Nothing Then Exit Sub
        ReLoadTDBGrid()
    End Sub

    Private Sub tsbClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    Private Sub tsbInherit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbInherit.Click, mnsInherit.Click, tsmInherit.Click
        If tdbg.RowCount <= 0 Then Exit Sub
        Dim f As New D02F7004
        With f
            .m_ReportCode = tdbg.Columns(COL_ReportCode).Text
            .FormState = EnumFormState.FormCopy
            .ShowDialog()
            .Dispose()
            If .bSaved Then LoadTDBGrid(True, .m_ReportCode)
        End With
    End Sub

    Private Sub mnsExportDataSetup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnsExportDataScript.Click, tsmExportDataScript.Click
        'Dim frm As New D80F2095
        'frm.FormName = "D02F7003"
        'frm.ModuleID = "02"
        'frm.Str01 = tdbg.Columns(COL_ReportCode).Text
        'frm.ShowDialog()
        'frm.Dispose()

        Dim arrPro() As StructureProperties = Nothing
        SetProperties(arrPro, "sFormName", "D02F7003") ' Tài liệu phân tích
        SetProperties(arrPro, "ModuleID", "02")
        SetProperties(arrPro, "sStr01", tdbg.Columns(COL_ReportCode).Text) ' Tài liệu phân tích
        SetProperties(arrPro, "sStr02", "") ' Tài liệu phân tích
        SetProperties(arrPro, "sStr03", "") ' Tài liệu phân tích
        SetProperties(arrPro, "sStr04", "") ' Tài liệu phân tích
        SetProperties(arrPro, "sStr05", "") ' Tài liệu phân tích
        SetProperties(arrPro, "sStr06", "") ' Tài liệu phân tích
        SetProperties(arrPro, "sStr07", "") ' Tài liệu phân tích
        SetProperties(arrPro, "sStr08", "") ' Tài liệu phân tích
        SetProperties(arrPro, "sStr09", "") ' Tài liệu phân tích
        SetProperties(arrPro, "sStr10", "") ' Tài liệu phân tích
        CallFormShowDialog("D80D0040", "D80F2095", arrPro)
    End Sub

End Class