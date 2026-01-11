Imports System.Drawing
Imports System.Net.NetworkInformation
Imports System.Windows.Forms

Public Class NetworkSpeedForm
    Inherits Form

    Private WithEvents Timer1 As Timer
    Private prevBytesReceived As Long = 0
    Private prevBytesSent As Long = 0

    Private drag As Boolean = False
    Private offset As Point

    Public Sub New()
        Me.FormBorderStyle = FormBorderStyle.None
        Me.TopMost = True
        Me.Width = 220
        Me.Height = 50
        Me.BackColor = Color.Black
        Me.ShowInTaskbar = False

        ' Position near bottom-right
        Dim bounds = Screen.PrimaryScreen.WorkingArea
        Me.Location = New Point(bounds.Right - Me.Width - 10, bounds.Bottom - Me.Height - 10)

        ' Draggable
        AddHandler Me.MouseDown, Sub(s, e)
                                     If e.Button = MouseButtons.Left Then
                                         drag = True
                                         offset = e.Location
                                     End If
                                 End Sub
        AddHandler Me.MouseMove, Sub(s, e)
                                     If drag Then
                                         Me.Location = New Point(Me.Left + e.X - offset.X, Me.Top + e.Y - offset.Y)
                                     End If
                                 End Sub
        AddHandler Me.MouseUp, Sub(s, e)
                                   drag = False
                               End Sub

        ' Initialize Timer
        Timer1 = New Timer()
        Timer1.Interval = 1000
        AddHandler Timer1.Tick, AddressOf Timer1_Tick
        Timer1.Start()

        ' Initialize previous bytes
        Dim nic = NetworkInterface.GetAllNetworkInterfaces() _
            .FirstOrDefault(Function(n) n.OperationalStatus = OperationalStatus.Up AndAlso n.NetworkInterfaceType <> NetworkInterfaceType.Loopback)
        If nic IsNot Nothing Then
            Dim stats = nic.GetIPv4Statistics()
            prevBytesReceived = stats.BytesReceived
            prevBytesSent = stats.BytesSent
        End If

        ' Force initial draw
        Me.Tag = New Tuple(Of Double, Double)(0, 0)
        Me.Invalidate()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        Dim nic = NetworkInterface.GetAllNetworkInterfaces() _
            .FirstOrDefault(Function(n) n.OperationalStatus = OperationalStatus.Up AndAlso n.NetworkInterfaceType <> NetworkInterfaceType.Loopback)
        If nic Is Nothing Then Return

        Dim stats = nic.GetIPv4Statistics()
        Dim downSpeed = stats.BytesReceived - prevBytesReceived
        Dim upSpeed = stats.BytesSent - prevBytesSent

        prevBytesReceived = stats.BytesReceived
        prevBytesSent = stats.BytesSent

        Me.Tag = New Tuple(Of Double, Double)(downSpeed, upSpeed)
        Me.Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim g = e.Graphics
        g.Clear(Me.BackColor)
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

        Dim speeds = TryCast(Me.Tag, Tuple(Of Double, Double))
        If speeds Is Nothing Then speeds = New Tuple(Of Double, Double)(0, 0)

        Using font As New Font("Consolas", 14, FontStyle.Bold)
            g.DrawString($"↓ {FormatSpeed(speeds.Item1)}", font, Brushes.Lime, 10, 5)
            g.DrawString($"↑ {FormatSpeed(speeds.Item2)}", font, Brushes.Cyan, 10, 25)
        End Using
    End Sub

    Private Function FormatSpeed(bytesPerSec As Double) As String
        Dim speed = bytesPerSec
        If speed < 1024 Then Return $"{speed:0} B/s"
        speed /= 1024
        If speed < 1024 Then Return $"{speed:0.0} KB/s"
        speed /= 1024
        Return $"{speed:0.00} MB/s"
    End Function
End Class


