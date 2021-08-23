'
' this form will merely show a window with cumulating PBX events 
' its purpose is to demonstrate the use of async events in VB.Net
'

Public Class Form1

    ' the forms PBX link
    Dim pbx As myPBX

    ' PBX access data
    ' assuming the controling user is "_TAPI_" which has a user-password "access"
    Const httpUser As String = "_TAPI_"
    Const httpPw As String = "access"
    Const pbxUser As String = "_TAPI_"
    Const pbxMonitor As String = "ckl-2"
    Const pbxUrl As String = "http://172.16.10.5/PBX0/user.soap"

    ' PBX runtime data
    Public pbxKey As Integer
    Public pbxSession As Integer
    Public pbxUserId As Integer

    ' initalize pbx link on load
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ' create the link
            pbx = New myPBX(Me)
            ' set the URL
            pbx.Url = pbxUrl
            ' set the HTTP level credentials
            pbx.Credentials = New Net.NetworkCredential(httpUser, httpPw, "")
            ' initialize the session, remember the session-id and -key
            pbxSession = pbx.Initialize(pbxUser, "VBNet SOAP Test", True, True, pbxKey)
            ' monitor a user
            pbxUserId = pbx.UserInitialize(pbxSession, pbxMonitor, False)
            ' start async retrieval of events from pbx
            pbx.startPolling()
        Catch err As Exception
            Dim msg1 As String
            Dim msg2 As String
            msg1 = err.Message
            If Not err.InnerException Is Nothing Then
                msg2 = err.InnerException.Message
            Else
                msg2 = ""
            End If
            MsgBox("pbx link initialization failed: " + msg1 + " / " + msg2)
            Exit Sub
        End Try

    End Sub
End Class


' the PBX link
' we need a derived class to be able to handle the async events
Public Class myPBX

    ' derive from the auto-generated soap class
    Inherits pbx_wsdl.pbx
    ' link to form
    Dim form As Form1

    ' constructor to save the form link
    Public Sub New(ByRef form As Form1)
        Me.form = form
    End Sub

    ' setup an async Poll
    Public Sub startPolling()
        Try
            Me.PollAsync(form.pbxSession)
        Catch e As Exception
            MsgBox("start poll failed: " + e.Message + " / " + e.InnerException.ToString)
        End Try
    End Sub

    Public Sub addEvent(ByVal ev As String)
        Me.form.events.Items.Add(ev)
    End Sub

    Delegate Sub addEventDelegate(ByVal ev As String)

    ' handle an async Poll result
    Private Sub pollCB(ByVal sender As Object, ByVal e As pbx_wsdl.PollCompletedEventArgs) _
        Handles MyBase.PollCompleted
        Dim p As addEventDelegate = AddressOf addEvent
        ' scan user and call events
        For Each ui As pbx_wsdl.UserInfo In e.Result.user
            ' Me.form.events.Items.Add("user " + ui.cn) -- dangerous
            Me.form.Invoke(p, New Object() {"user " + ui.cn})
        Next
        For Each ci As pbx_wsdl.CallInfo In e.Result.call
            ' Me.form.events.Items.Add("call -> " + ci.msg) -- dangerous
            Me.form.Invoke(p, New Object() {"call -> " + ci.msg})
        Next
        ' schedule next async Poll
        Me.startPolling()
    End Sub

End Class