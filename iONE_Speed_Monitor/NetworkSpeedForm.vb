Imports System.Drawing
Imports System.Net.NetworkInformation
Imports System.Windows.Forms

Public Class NetworkSpeedForm
    Inherits Form

    Private WithEvents Timer1 As Timer
    Private prevBytesReceived As Long
    Private prevBytesSent As Long

    ' Drag support
    Private dragging As Boolean = False
    Private dragOffset As Point

    ' Context menu
    Private contextMenu As ContextMenuStrip

    Private Sub ShowAbout(sender As Object, e As EventArgs)
        MessageBox.Show(
            "iONE Speed Monitor" & vbCrLf & vbCrLf &
            "Version 1.0" & vbCrLf &
            "Author: Iraz" & vbCrLf &
            "© 2026 iONE" & vbCrLf & vbCrLf &
            "Lightweight network speed monitor for Windows.",
            "About",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
    End Sub

    Public Sub New()
        ' ---------------- Form setup ----------------
        Me.FormBorderStyle = FormBorderStyle.None
        Me.TopMost = True
        Me.ShowInTaskbar = False
        Me.Width = 155
        Me.Height = 48
        Me.BackColor = Color.Black
        Me.DoubleBuffered = True

        Me.StartPosition = FormStartPosition.Manual
        ' Position window at bottom-right (above taskbar)
        Dim bounds As Rectangle = Screen.PrimaryScreen.WorkingArea
        Me.Location = New Point(
            bounds.Right - Me.Width - 0,
            bounds.Bottom - Me.Height - 0
        )

        contextMenu = New ContextMenuStrip()

        Dim aboutItem As New ToolStripMenuItem("About")
        AddHandler aboutItem.Click, AddressOf ShowAbout

        Dim exitItem As New ToolStripMenuItem("Exit")
        AddHandler exitItem.Click, AddressOf ExitApp

        contextMenu.Items.Add(aboutItem)
        contextMenu.Items.Add(New ToolStripSeparator())
        contextMenu.Items.Add(exitItem)

        Me.ContextMenuStrip = contextMenu


        ' ---------------- Drag handlers ----------------
        AddHandler Me.MouseDown, AddressOf Form_MouseDown
        AddHandler Me.MouseMove, AddressOf Form_MouseMove
        AddHandler Me.MouseUp, AddressOf Form_MouseUp

        ' ---------------- Timer ----------------
        Timer1 = New Timer()
        Timer1.Interval = 1000
        AddHandler Timer1.Tick, AddressOf Timer1_Tick
        Timer1.Start()

        ' ---------------- Init counters ----------------
        Dim nic = GetActiveNIC()
        If nic IsNot Nothing Then
            Dim stats = nic.GetIPv4Statistics()
            prevBytesReceived = stats.BytesReceived
            prevBytesSent = stats.BytesSent
        End If

        Me.Tag = New Tuple(Of Double, Double)(0, 0)
    End Sub

    ' ---------------- Drag logic ----------------
    Private Sub Form_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            dragging = True
            dragOffset = e.Location
        End If
    End Sub

    Private Sub Form_MouseMove(sender As Object, e As MouseEventArgs)
        If dragging Then
            Me.Location = New Point(Me.Left + e.X - dragOffset.X, Me.Top + e.Y - dragOffset.Y)
        End If
    End Sub

    Private Sub Form_MouseUp(sender As Object, e As MouseEventArgs)
        dragging = False
    End Sub

    ' ---------------- Timer update ----------------
    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        Dim nic = GetActiveNIC()
        If nic Is Nothing Then Return

        Dim stats = nic.GetIPv4Statistics()
        Dim downSpeed = stats.BytesReceived - prevBytesReceived
        Dim upSpeed = stats.BytesSent - prevBytesSent

        prevBytesReceived = stats.BytesReceived
        prevBytesSent = stats.BytesSent

        Me.Tag = New Tuple(Of Double, Double)(downSpeed, upSpeed)
        Me.Invalidate()
    End Sub

    ' ---------------- Paint ----------------
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim g = e.Graphics
        g.Clear(Color.Black)
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

        Dim speeds = CType(Me.Tag, Tuple(Of Double, Double))

        Using font As New Font("Consolas", 14, FontStyle.Bold)
            g.DrawString($"↓ {FormatSpeed(speeds.Item1)}", font, Brushes.Lime, 10, 5)
            g.DrawString($"↑ {FormatSpeed(speeds.Item2)}", font, Brushes.Cyan, 10, 30)
        End Using
    End Sub

    ' ---------------- Helpers ----------------
    Private Function FormatSpeed(bytesPerSec As Double) As String
        If bytesPerSec < 1024 Then Return $"{bytesPerSec:0} B/s"
        bytesPerSec /= 1024
        If bytesPerSec < 1024 Then Return $"{bytesPerSec:0.0} KB/s"
        bytesPerSec /= 1024
        Return $"{bytesPerSec:0.00} MB/s"
    End Function

    Private Function GetActiveNIC() As NetworkInterface
        Return NetworkInterface.GetAllNetworkInterfaces().
            FirstOrDefault(Function(n) n.OperationalStatus = OperationalStatus.Up AndAlso
                                       n.NetworkInterfaceType <> NetworkInterfaceType.Loopback)
    End Function

    ' ---------------- Exit ----------------
    Private Sub ExitApp(sender As Object, e As EventArgs)
        Timer1.Stop()
        Application.Exit()
    End Sub
End Class
