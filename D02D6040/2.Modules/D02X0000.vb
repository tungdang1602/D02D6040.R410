''' <summary>
''' D02E6040: Chứa các màn hình Báo cáo
''' Sub Main và các vấn đề liên quan đến việc khởi động exe con
''' </summary>
Module D02X0000

    Public Sub Main()

        SetSysDateTime()
#If DEBUG Then 'Nếu đang ở trạng thái DEBUG thì ...
        'CheckDLL() 'Kiểm tra các DLL tương thích và các file Module hợp lệ
        MakeVirtualConnection() 'tạo kết nối ảo
        SaveParameter() 'Gán giá trị các thông số vào Registry
#Else 'Đang trong trạng thái thực thi exe
        If PrevInstance() Then End 'Kiểm tra nếu chương trình đã chạy rồi thì END
        ReadLanguage() 'Đọc biến ngôn ngữ ở đây nhằm mục đích để báo lỗi theo ngôn ngữ cho những phần sau
        If Not CheckSecurity() Then End 'Kiểm tra an toàn cho chương trình, nếu không an toàn thì END
#End If
        GetAllParameter() 'Đọc các giá trị từ Registry lưu vào biến toàn cục

        If Not CheckConnection() Then End 'Kiểm tra nối không kết nối được với Server thì END
        'Update 19/11/2010: Kiểm tra đồng bộ exe và fix 
        If Not CheckExeFixSynchronous(My.Application.Info.AssemblyName) Then End

        'If Not CheckOther() Then End 'Vì lý do gì đó, có thể kiểm tra một điều kiện không hợp lệ và có thể kết thúc chương trình
        'Tới đây quá trình kiểm tra cho modlue đã hoàn thành, không còn lệnh END để kết thúc chương trình nữa
        LoadSystemInfo()
        LoadOptions() 'Load các thông số cho phần tùy chọn
        LoadOthers() 'Các lập trình viên có thể load những thứ khác ở đây

        'Xóa Registry
#If DEBUG Then
        PARA_FormID = "D02F7003"
        gbUnicode = True
#Else
        D99C0007.RegDeleteExe(EXECHILD)
#End If

        'Hiển thị form tương ứng
        'PARA_FormID = "D02F3040"
Select Case PARA_FormID
'Gọi form nhận tham số
            Case Else

                Try
                    'Gọi form không nhận tham số. Default 
                    Dim frm As New Form
                    Dim frmName As String = PARA_FormID
                    frmName = System.Reflection.Assembly.GetEntryAssembly.GetName.Name & "." & frmName
                    frm = DirectCast(System.Reflection.Assembly.GetEntryAssembly.CreateInstance(frmName), Form)
                    frm.ShowInTaskbar = True
                    frm.ShowDialog()
                    frm.Dispose()
                Catch ex As Exception
                    D99C0008.MsgL3(ex.Message)
                End Try
        End Select
        KillChildProcess(MODULED02)
    End Sub

    Private Function CheckOther() As Boolean
        Return True
    End Function

    Private Sub LoadOthers()
        GetModuleAdmin(D02)
        GeneralItems()
    End Sub

    Private Function PrevInstance() As Boolean
        If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
            Return True
        End If
        Return False
    End Function

    Private Sub ReadLanguage()
        Dim sLanguage As String = GetSetting("Lemon3 System Module", "Caption Setting", "Language", "0")
        If sLanguage = "0" Then
            geLanguage = EnumLanguage.Vietnamese
            gsLanguage = "84"
        Else
            geLanguage = EnumLanguage.English
            gsLanguage = "01"
        End If
        D99C0008.Language = geLanguage
        MsgAnnouncement = IIf(geLanguage = EnumLanguage.Vietnamese, "Th¤ng bÀo", "Announcement").ToString

    End Sub

    Private Function CheckSecurity() As Boolean
        Dim D00_CompanyName As String
        Dim D00_LegalCopyright As String
        Dim CompanyName As String
        Dim LegalCopyright As String

        If Not System.IO.File.Exists(Application.StartupPath & "\D00E0030.EXE") Then
            If gsLanguage = "84" Then
                MessageBox.Show("Thï tóc gãi nèi bè bÊt híp lÖ! (10)", "Th¤ng bÀo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Else
                MessageBox.Show("Invalid internal system call! (10)", "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End If
            Return False
        Else
            Dim D00_FiVerInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.StartupPath & "\D00E0030.EXE")
            Dim FiVerInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.StartupPath & "\" & MODULED02 & ".EXE")
            D00_CompanyName = D00_FiVerInfo.CompanyName
            D00_LegalCopyright = D00_FiVerInfo.LegalCopyright
            CompanyName = FiVerInfo.CompanyName
            LegalCopyright = FiVerInfo.LegalCopyright
            If (D00_CompanyName <> CompanyName) OrElse (D00_LegalCopyright <> LegalCopyright) Then
                If gsLanguage = "84" Then
                    MessageBox.Show("Thï tóc gãi nèi bè bÊt híp lÖ! (10)", "Th¤ng bÀo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Else
                    MessageBox.Show("Invalid internal system call! (10)", "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End If
                Return False
            End If
        End If

        Dim CommandArgs() As String = Environment.GetCommandLineArgs()

        If CommandArgs.Length <> 3 OrElse CommandArgs(1) <> "/DigiNet" OrElse CommandArgs(2) <> "Corporation" Then
            If gsLanguage = "84" Then
                MessageBox.Show("Thï tóc gãi nèi bè bÊt híp lÖ! (12)", "Th¤ng bÀo", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Else
                MessageBox.Show("Invalid internal system call! (12)", "Announcement", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End If
            Return False
        End If
        Return True
    End Function

    Private Sub MakeVirtualConnection()
        gsServer = "drd14"
        gsCompanyID = "drd02"
        gsConnectionUser = "sa"
        gsPassword = ""
        gsUserID = "LEMONADMIN"

        gsServer = "drd81"
        gsCompanyID = "na"
        gsPassword = "234"
    End Sub

    Private Sub GetAllParameter()
        PARA_Server = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "ServerName", "", CodeOption.lmCode)
        PARA_Database = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "DBName", "", CodeOption.lmCode)
        PARA_ConnectionUser = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "ConnectionUserID", "", CodeOption.lmCode)
        PARA_UserID = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "UserID", "", CodeOption.lmCode)
        PARA_Password = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Password", "", CodeOption.lmCode)
        PARA_DivisionID = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "DivisionID", "HANOI")
        PARA_TranMonth = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "TranMonth", "01")
        PARA_TranYear = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "TranYear", "2007")
        PARA_Language = CType(D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Language", "84"), EnumLanguage)
        PARA_FormID = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Ctrl01", "")
        PARA_FormIDPermission = D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Ctrl03", "")
        '-----------------------------------------------------------------------
        gbUnicode = CType(D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "CodeTable", "False"), Boolean)
        AssignToPublicVariable()
     
    End Sub

    Private Sub AssignToPublicVariable()
        gsServer = PARA_Server
        gsCompanyID = PARA_Database
        gsConnectionUser = PARA_ConnectionUser
        gsUserID = PARA_UserID
        gsPassword = PARA_Password
        gsDivisionID = PARA_DivisionID
        giTranMonth = CInt(PARA_TranMonth)
        giTranYear = CInt(PARA_TranYear)
        geLanguage = PARA_Language
        gsLanguage = IIf(geLanguage = EnumLanguage.Vietnamese, "84", "01").ToString
        D99C0008.Language = geLanguage
        PARA_FormID = PARA_FormID
        PARA_FormIDPermission = PARA_FormIDPermission
        '-----------------------------------------------------------------------        
    End Sub

    Private Sub SaveParameter()
        Dim sFormID As String = "D02F7003"
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ServerName", gsServer, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DBName", gsCompanyID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ConnectionUserID", gsConnectionUser, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "UserID", gsUserID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Password", gsPassword, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DivisionID", "HANOI")
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranMonth", "5")
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranYear", "2008")
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Language", "0")
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl01", sFormID) 'PARA_FormID
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl03", sFormID) 'PARA_FormIDPermission
    End Sub

#Region "Hàm dùng cho WEBSERVICE"
    Private Sub RemoteConnection()
        GetInfoWebService()
        If CheckRemoteConnection(gsAppServer) = False Then
            D99C0008.MsgInvalidConnection()
            End
        End If
    End Sub
    ''' <summary>
    ''' Kiểm tra kết nối của Webservice
    ''' </summary>
    ''' <param name="sHttp"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckRemoteConnection(ByVal sHttp As String) As Boolean
        Try
            CallWebService.Url = sHttp & "D91W0000.asmx"
            'CallWebService.Timeout = 10000000
            CallWebService.Timeout = nWSTimeOut
            CallWebService.UserExists("LEMONADMIN", gsWSSPara01, gsWSSPara02, gsWSSPara03, gsWSSPara04, gsWSSPara05)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Lấy thông tin của Webservice
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GetInfoWebService()
        gsAppServer = D99C0007.GetOthersSetting(D02, MODULED02, "AppServer")
        gsWSSPara01 = D99C0007.GetOthersSetting(D02, MODULED02, "WSSPara01")
        gsWSSPara02 = D99C0007.GetOthersSetting(D02, MODULED02, "WSSPara02")
        gsWSSPara03 = D99C0007.GetOthersSetting(D02, MODULED02, "WSSPara03")
        gsWSSPara04 = D99C0007.GetOthersSetting(D02, MODULED02, "WSSPara04")
        gsWSSPara05 = D99C0007.GetOthersSetting(D02, MODULED02, "WSSPara05")
    End Sub
    ''' <summary>
    ''' Ghi các thông tin về Webservice để test
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetInfoWS()
        gsAppServer = "http://webservice/ws2005/"
        D99C0007.SaveOthersSetting(D02, MODULED02, "AppMode", "1")
        D99C0007.SaveOthersSetting(D02, MODULED02, "AppServer", gsAppServer)
        D99C0007.SaveOthersSetting(D02, MODULED02, "WSSPara01", gsWSSPara01)
        D99C0007.SaveOthersSetting(D02, MODULED02, "WSSPara02", gsWSSPara02)
        D99C0007.SaveOthersSetting(D02, MODULED02, "WSSPara03", gsWSSPara03)
        D99C0007.SaveOthersSetting(D02, MODULED02, "WSSPara04", gsWSSPara04)
        D99C0007.SaveOthersSetting(D02, MODULED02, "WSSPara05", gsWSSPara05)
    End Sub
#End Region

End Module
