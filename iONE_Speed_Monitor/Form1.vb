Imports System.Drawing
Imports System.Drawing.Text
Imports System.Windows.Forms
Imports System.Net.NetworkInformation

Public Class Form1
    Inherits Form

    Private WithEvents Timer1 As Timer
    Private NotifyIcon1 As NotifyIcon

    ' Previous bytes for calculating speed
    Private prevBytesReceived As Long = 0
    Private prevBytesSent As Long = 0

    Public Sub New()
        InitializeComponent()

        ' ---------------------------
        ' Hide form completely
        ' ---------------------------
        Me.WindowState = FormWindowState.Minimized
        Me.ShowInTaskbar = False
        Me.Visible = False

        ' ---------------------------
        ' Create NotifyIcon
        ' ---------------------------
        NotifyIcon1 = New NotifyIcon()
        NotifyIcon1.Icon = SystemIcons.Information
        NotifyIcon1.Visible = True
        NotifyIcon1.Text = "Network Speed Monitor"

        ' Tray context menu
        Dim trayMenu As New ContextMenuStrip()
        Dim exitItem As New ToolStripMenuItem("Exit")
        AddHandler exitItem.Click, AddressOf OnExit
        trayMenu.Items.Add(exitItem)
        NotifyIcon1.ContextMenuStrip = trayMenu

        ' ---------------------------
        ' Create Timer
        ' ---------------------------
        Timer1 = New Timer()
        Timer1.Interval = 1000 ' update every second
        Timer1.Start()
    End Sub

    ' ---------------------------
    ' Timer Tick: update tray icon
    ' ---------------------------
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Get first active network interface
        Dim nic As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces() _
            .Where(Function(n) n.OperationalStatus = OperationalStatus.Up AndAlso n.NetworkInterfaceType <> NetworkInterfaceType.Loopback) _
            .FirstOrDefault()

        If nic Is Nothing Then Return

        Dim stats As IPv4InterfaceStatistics = nic.GetIPv4Statistics()
        Dim bytesReceived As Long = stats.BytesReceived
        Dim bytesSent As Long = stats.BytesSent

        ' Calculate speed in bytes/sec
        Dim downBytes As Double = bytesReceived - prevBytesReceived
        Dim upBytes As Double = bytesSent - prevBytesSent

        prevBytesReceived = bytesReceived
        prevBytesSent = bytesSent

        ' Create dynamic icon showing both ↓ and ↑ speeds
        Dim icon As Icon = CreateIconFromText(downBytes, upBytes)
        Dim oldIcon = NotifyIcon1.Icon
        NotifyIcon1.Icon = icon
        If oldIcon IsNot Nothing Then oldIcon.Dispose()

        ' Update tooltip with full speeds
        NotifyIcon1.Text = $"↓ {FormatSpeed(downBytes)}{vbCrLf}↑ {FormatSpeed(upBytes)}"
    End Sub

    ' ---------------------------
    ' Format speed for tray icon (short)
    ' ---------------------------
    Private Function FormatSpeedForIcon(bytesPerSec As Double) As String
        If bytesPerSec >= 1024 * 1024 Then
            Return (bytesPerSec / 1024 / 1024).ToString("0") & "M"
        ElseIf bytesPerSec >= 1024 Then
            Return (bytesPerSec / 1024).ToString("0") & "K"
        Else
            Return bytesPerSec.ToString("0")
        End If
    End Function

    ' ---------------------------
    ' Format speed for tooltip (full)
    ' ---------------------------
    Private Function FormatSpeed(bytesPerSec As Double) As String
        Dim speed As Double = bytesPerSec

        If speed < 1024 Then
            Return $"{speed:0} B/s"
        End If

        speed /= 1024
        If speed < 1024 Then
            Return $"{speed:0.0} KB/s"
        End If

        speed /= 1024
        If speed < 1024 Then
            Return $"{speed:0.00} MB/s"
        End If

        speed /= 1024
        Return $"{speed:0.00} GB/s"
    End Function

    ' ---------------------------
    ' Create a 32x32 icon with ↓ and ↑ stacked
    ' ---------------------------
    Private Function CreateIconFromText(downSpeed As Double, upSpeed As Double) As Icon
        Dim bmp As New Bitmap(64, 64)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.Clear(Color.Transparent)
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit
            Using brush As New SolidBrush(Color.White)
                Using font As New Font("Arial", 10, FontStyle.Bold)
                    g.DrawString($"↓{FormatSpeedForIcon(downSpeed)}", font, brush, 0, 0)
                    g.DrawString($"↑{FormatSpeedForIcon(upSpeed)}", font, brush, 0, 16)
                End Using
            End Using
        End Using

        Dim hIcon As IntPtr = bmp.GetHicon()
        Dim icon As Icon = Icon.FromHandle(hIcon)
        Return icon
    End Function

    ' ---------------------------
    ' Exit application from tray
    ' ---------------------------
    Private Sub OnExit(sender As Object, e As EventArgs)
        NotifyIcon1.Visible = False
        Application.Exit()
    End Sub

    ' ---------------------------
    ' Hide form immediately after shown
    ' ---------------------------
    Protected Overrides Sub OnShown(e As EventArgs)
        Me.Hide()
        MyBase.OnShown(e)
    End Sub

    ' ---------------------------
    ' Prevent app from exiting when form is closed
    ' ---------------------------
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        e.Cancel = True
        Me.Hide()
        MyBase.OnFormClosing(e)
    End Sub
End Class
