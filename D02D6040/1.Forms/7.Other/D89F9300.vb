Public Class D89F9300
    Private WithEvents backgroundWorker1 As System.ComponentModel.BackgroundWorker
    Private Const EXECHILD As String = "D89E0240"

    Private exe As D89E0240

    Private _FormName As String
    Public WriteOnly Property FormName() As String
        Set(ByVal Value As String)
            _FormName = Value
        End Set
    End Property

    Private _FormPermission As String
    Public WriteOnly Property FormPermission() As String
        Set(ByVal Value As String)
            _FormPermission = Value
        End Set
    End Property

    Private _FormStatus As EnumFormState
    Public WriteOnly Property FormStatus() As EnumFormState
        Set(ByVal Value As EnumFormState)
            _FormStatus = Value
        End Set
    End Property

    Private _key01ID As String
    Public WriteOnly Property Key01ID() As String
        Set(ByVal Value As String)
            _key01ID = Value
        End Set
    End Property

    Private _key02ID As String
    Public WriteOnly Property Key02ID() As String
        Set(ByVal Value As String)
            _key02ID = Value
        End Set
    End Property

    Private _key03ID As String
    Public WriteOnly Property Key03ID() As String
        Set(ByVal Value As String)
            _key03ID = Value
        End Set
    End Property

    Private _key04ID As String
    Public WriteOnly Property Key04ID() As String
        Set(ByVal Value As String)
            _key04ID = Value
        End Set
    End Property

    Private _key05ID As String
    Public WriteOnly Property Key05ID() As String
        Set(ByVal Value As String)
            _key05ID = Value
        End Set
    End Property

    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundWorker1.DoWork
        'Tạo một process gắn với exe con, process này sẽ quan sát exe con.
        Dim p As System.Diagnostics.Process
        Try
            p = Process.GetProcessesByName(EXECHILD)(0)

            If p Is Nothing Then
                Exit Sub
            End If

            'Chờ đợi exe con tắt tiến trình 
            p.EnableRaisingEvents = True
            p.WaitForExit()
        Catch ex As Exception
        End Try
    End Sub

    Public Sub FormLock_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Ẩn form trung gian
        Me.Size = New Size(0, 0)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None

        '----Truyền tham số exe con------
        exe = New D89E0240(gsServer, gsCompanyID, gsConnectionUser, gsPassword, gsUserID, IIf(geLanguage = EnumLanguage.Vietnamese, "0", "10000").ToString(), gsDivisionID, giTranMonth, giTranYear)


        exe.FormPermission = _FormPermission
        If _formName = "D89F9300" Then
            exe.FormActive = D89E0240Form.D89F9300
        End If

        exe.FormStatus = _FormStatus

        exe.Key01ID = _key01ID
        exe.Key02ID = _key02ID
        exe.Key03ID = _key03ID
        exe.Key04ID = _key04ID
        exe.Key05ID = _key05ID

        exe.Run()
        '------------------------------------

        'Bắt đầu chạy cơ chế background
        backgroundWorker1 = New System.ComponentModel.BackgroundWorker
        backgroundWorker1.RunWorkerAsync()
    End Sub

    'sự kiện hoàn thành và dừng của Background
    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundWorker1.RunWorkerCompleted

        Me.Close()
    End Sub
End Class