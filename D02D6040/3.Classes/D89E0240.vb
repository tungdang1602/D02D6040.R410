''' <summary>
''' Các màn hình của exe con D89E0240: Mẫu báo cáo
''' </summary>
Public Enum D89E0240Form
    ''' <summary>
    ''' D89F9101: Danh sách mẫu báo cáo
    ''' </summary>
    D89F9101 = 0
    ''' <summary>
    ''' D89F9300: Định nghĩa dữ liệu báo cáo
    ''' </summary>
    ''' <remarks></remarks>
    D89F9300 = 1
End Enum

'Cách gọi:
'Dim exe As New D89E0240(gsServer, gsCompanyID, gsConnectionUser, gsPassword, gsUserID, IIf(geLanguage = EnumLanguage.Vietnamese, "0", "10000").ToString, gsDivisionID, giTranMonth, giTranYear)
'        With exe
'            .FormActive = D89E0240Form.D89F9101
'            .FormPermission = "D22F9101"
'            .Run()
'        End With

Public Class D89E0240
    Private Const EXEMODULE As String = "D89"
    Private Const EXECHILD As String = "D89E0240"
    Private sLanguage As String

    ''' <summary>
    ''' Khởi tạo exe con D89E0240
    ''' </summary>
    ''' <param name="Server">Server kết nối đến hệ thống</param>
    ''' <param name="Database">Database kết nối đến hệ thống</param>
    ''' <param name="UserDatabaseID">User Database kết nối đến hệ thống</param>
    ''' <param name="Password">Password kết nối đến hệ thống</param>
    ''' <param name="UserID">User Lemon3 kết nối đến hệ thống</param>
    ''' <param name="Language">Ngôn ngữ sử dụng</param>
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UserDatabaseID As String, ByVal Password As String, ByVal UserID As String, ByVal Language As String)
        sLanguage = Language
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ServerName", Server, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DBName", Database, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ConnectionUserID", UserDatabaseID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Password", Password, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "UserID", UserID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Language", Language)
    End Sub

    ''' <summary>
    ''' Khởi tạo exe con D89E0240
    ''' </summary>
    ''' <param name="Server">Server kết nối đến hệ thống</param>
    ''' <param name="Database">Database kết nối đến hệ thống</param>
    ''' <param name="UserDatabaseID">User Database kết nối đến hệ thống</param>
    ''' <param name="Password">Password kết nối đến hệ thống</param>
    ''' <param name="UserID">User Lemon3 kết nối đến hệ thống</param>
    ''' <param name="Language">Ngôn ngữ sử dụng</param>
    ''' <param name="DivisionID">Đơn vị hiện tại</param>
    ''' <param name="TranMonth">Tháng kế toán hiện tại</param>
    ''' <param name="TranYear">Năm kế toán hiện tại</param>
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UserDatabaseID As String, ByVal Password As String, ByVal UserID As String, ByVal Language As String, ByVal DivisionID As String, ByVal TranMonth As Integer, ByVal TranYear As Integer)
        Me.New(Server, Database, UserDatabaseID, Password, UserID, Language)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DivisionID", DivisionID)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranMonth", TranMonth.ToString)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranYear", TranYear.ToString)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "CodeTable", gbUnicode.ToString)

    End Sub

    ''' <summary>
    ''' Màn hình cần hiển thị cho exe con
    ''' </summary>
    Public WriteOnly Property FormActive() As D89E0240Form
        Set(ByVal Value As D89E0240Form)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl01", [Enum].GetName(GetType(D89E0240Form), Value))
        End Set
    End Property

    ''' <summary>
    ''' Màn hình phân quyền cho exe con
    ''' </summary>
    Public WriteOnly Property FormPermission() As String
        Set(ByVal Value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl03", Value)
        End Set
    End Property

    ''' <summary>
    ''' Màn hình phân quyền cho exe con
    ''' </summary>
    Public WriteOnly Property FormStatus() As EnumFormState
        Set(ByVal Value As EnumFormState)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl02", CByte(Value).ToString)
        End Set
    End Property

    ''' <summary>
    ''' Truyền vào Key01ID: FormID
    ''' </summary>
    Public WriteOnly Property Key01ID() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID01", value)
        End Set
    End Property

    ''' <summary>
    ''' Truyền vào Key02ID: ReportID
    ''' </summary>
    Public WriteOnly Property Key02ID() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID02", value)
        End Set
    End Property

    ''' <summary>
    ''' Truyền vào Key03ID: Tiêu đề
    ''' </summary>
    Public WriteOnly Property Key03ID() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID03", value)
        End Set
    End Property

    ''' <summary>
    ''' Truyền vào Key04ID
    ''' </summary>
    Public WriteOnly Property Key04ID() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID04", value)
        End Set
    End Property

    ''' <summary>
    ''' Truyền vào Key05ID
    ''' </summary>
    Public WriteOnly Property Key05ID() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID05", value)
        End Set
    End Property

    ''' <summary>
    ''' Truyền vào ModuleID
    ''' </summary>
    Public WriteOnly Property ModuleID() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ModuleID", value)
        End Set
    End Property

    ''' <summary>
    ''' Thực thi exe con
    ''' </summary>
    Public Sub Run()
        If Not ExistFile(My.Application.Info.DirectoryPath & "\" & EXECHILD & ".exe") Then Exit Sub
        Dim pInfo As New System.Diagnostics.ProcessStartInfo(My.Application.Info.DirectoryPath & "\" & EXECHILD & ".exe")
        pInfo.Arguments = "/DigiNet Corporation"
        pInfo.WindowStyle = ProcessWindowStyle.Normal
        Process.Start(pInfo)
    End Sub

    ''' <summary>
    ''' Kiểm tra tồn tại exe con không ?
    ''' </summary>
    Private Function ExistFile(ByVal Path As String) As Boolean
        If System.IO.File.Exists(Path) Then Return True
        If geLanguage = EnumLanguage.Vietnamese Then
            D99C0008.MsgL3("Không tồn tại file " & EXECHILD & ".exe")
        Else
            D99C0008.MsgL3("Not exist file " & EXECHILD & ".exe")
        End If
        Return False
    End Function
End Class
