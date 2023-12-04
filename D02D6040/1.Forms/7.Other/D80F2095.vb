'#-------------------------------------------------------------------------------------
'# Created User: HOANGNHAN
'# Created Date: 01/03/2013 2:25:08 PM
'# Modify User: HOANGNHAN
'# Modify Date: 01/03/2013 2:25:08 PM
'# Description: Xuất dữ liệu ra File SQL
'#-------------------------------------------------------------------------------------

Imports System
Imports System.Text

Public Class D80F2095
    Private WithEvents backgroundWorker1 As System.ComponentModel.BackgroundWorker
    Private ChildName As String = "D80E0440"
    Dim exe As D80E0440
    Dim p As System.Diagnostics.Process

#Region "Property"

    Private _formActive As String = ""
    Public WriteOnly Property FormActive() As String
        Set(ByVal Value As String)
            _formActive = Value
        End Set
    End Property

    Private _formState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal Value As EnumFormState)
            _formState = Value
        End Set
    End Property

    Private _formPermission As String = ""
    Public WriteOnly Property FormPermission() As String
        Set(ByVal Value As String)
            _formPermission = Value
        End Set
    End Property

    Private _moduleID As String = ""
    Public WriteOnly Property ModuleID() As String
        Set(ByVal Value As String)
            _moduleID = Value
        End Set
    End Property

    Private _formName As String = ""
    Public WriteOnly Property FormName() As String 
        Set(ByVal Value As String )
            _formName = Value
        End Set
    End Property

    Private _mode As String= ""
    Public WriteOnly Property Mode() As String ' Tương ứng với biến Type
        Set(ByVal Value As String)
            _mode = Value
        End Set
    End Property

    Private _str01 As String = ""
    Public WriteOnly Property Str01() As String
        Set(ByVal Value As String)
            _str01 = Value
        End Set
    End Property

    Private _str02 As String = ""
    Public WriteOnly Property Str02() As String
        Set(ByVal Value As String)
            _str02 = Value
        End Set
    End Property

    Private _str03 As String = ""
    Public WriteOnly Property Str03() As String
        Set(ByVal Value As String)
            _str03 = Value
        End Set
    End Property

    Private _str04 As String = ""
    Public WriteOnly Property Str04() As String
        Set(ByVal Value As String)
            _str04 = Value
        End Set
    End Property

    Private _str05 As String = ""
    Public WriteOnly Property Str05() As String
        Set(ByVal Value As String)
            _str05 = Value
        End Set
    End Property

    Private _str06 As String = ""
    Public WriteOnly Property Str06() As String
        Set(ByVal Value As String)
            _str06 = Value
        End Set
    End Property

    Private _str07 As String = ""
    Public WriteOnly Property Str07() As String
        Set(ByVal Value As String)
            _str07 = Value
        End Set
    End Property

    Private _str08 As String = ""
    Public WriteOnly Property Str08() As String
        Set(ByVal Value As String)
            _str08 = Value
        End Set
    End Property

    Private _str09 As String = ""
    Public WriteOnly Property Str09() As String
        Set(ByVal Value As String)
            _str09 = Value
        End Set
    End Property

    Private _str10 As String = ""
    Public WriteOnly Property Str10() As String
        Set(ByVal Value As String)
            _str10 = Value
        End Set
    End Property

    Private _outPut01 As Boolean ' Kết quả trả về
    Public ReadOnly Property OutPut01() As Boolean
        Get
            Return _outPut01
        End Get
    End Property
#End Region

    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundWorker1.DoWork
        'Tạo một process gắn với exe con, process này sẽ quan sát exe con.
        Dim p As System.Diagnostics.Process
        Try
            p = Process.GetProcessesByName(ChildName)(0)
            If p Is Nothing Then
                Exit Sub
            End If
            p.EnableRaisingEvents = True
            p.WaitForExit()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub FormLock_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Ẩn form trung gian
        Me.Size = New Size(0, 0)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None

        '----Truyền tham số exe con------
        exe = New D80E0440(gsServer, gsCompanyID, gsConnectionUser, gsPassword, gsUserID, IIf(geLanguage = EnumLanguage.Vietnamese, "0", "10000").ToString, gsDivisionID, giTranMonth, giTranYear)
        With exe
            .FormActive = D80E0440Form.D80F2095
            .FormPermission = _formPermission
            .ModuleID = _moduleID
            .ID01 = _formName
            .ID02 = _mode
            .ID03 = _str01
            .ID04 = _str02
            .ID05 = _str03
            .ID06 = _str04
            .ID07 = _str05
            .ID08 = _str06
            .ID09 = _str07
            .ID10 = _str08
            .ID11 = _str09
            .ID12 = _str10
            .Run()
        End With

        'Bắt đầu chạy cơ chế background
        backgroundWorker1 = New System.ComponentModel.BackgroundWorker
        backgroundWorker1.RunWorkerAsync()
    End Sub

    'sự kiện hoàn thành và dừng của Background
    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundWorker1.RunWorkerCompleted
        _outPut01 = exe.Output01
        Me.Close()
    End Sub
End Class