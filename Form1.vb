
Public Class Form1
    ReadOnly dataset(3, 10000) As Double
    ReadOnly gridcolor As Color = Color.FromArgb(30, 30, 30)
    ReadOnly plotcolor(3) As Color
    Private Xscale As Double = 1 / 500
    Private IsOffsetChanged As Boolean = False
    Private IsMagnificationChanged As Boolean = False
    Private IsMarkerMoved As Boolean = False
    Private MouseDownX As Double
    Private MouseDownY As Double
    Private MouseDownXo As Double
    Private magx_s As Double
    Private magx As Double = 1
    Private magy As Double = 1
    Private offx As Integer = 0
    Private offy As Integer = 0
    Private bmpwidth As Integer = 1
    Private bmpheight As Integer = 1
    Private d As Double
    Private dd As Double
    Private tmp As Integer = -1
    Private tmp2 As Integer
    Private tmpval1 As Double
    Private tmpval2 As Double
    Private ps As Integer = 0
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i = 0 To 10000
            dataset(0, i) = Math.Sin(i / 60) * 100 + 20
            dataset(1, i) = Math.Sin(i / 30) * 50
            dataset(2, i) = Math.Sin(i / 80) * 70
            dataset(3, i) = Math.Sin(i / 100) * 20
        Next
        plotcolor(0) = Color.Lime
        plotcolor(1) = Color.Yellow
        plotcolor(2) = Color.Blue
        plotcolor(3) = Color.Magenta
        RadioButton1.ForeColor = plotcolor(0)
        RadioButton2.ForeColor = plotcolor(1)
        RadioButton3.ForeColor = plotcolor(2)
        RadioButton4.ForeColor = plotcolor(3)
        AutoZoom(3)
    End Sub
    Private Sub PanelDB1_Paint(sender As Object, e As PaintEventArgs) Handles PanelDB1.Paint
        Dim bmp As New Bitmap(PanelDB1.Width, PanelDB1.Height)
        bmpwidth = bmp.Width
        bmpheight = bmp.Height

        For i = 0 To bmpwidth - 1
            bmp.SetPixel(i, bmpheight / 2, gridcolor)
            bmp.SetPixel(i, bmpheight / 4, gridcolor)
            bmp.SetPixel(i, bmpheight * 3 / 4, gridcolor)
        Next
        For i = 0 To bmpheight - 1
            bmp.SetPixel(bmpwidth / 2, i, gridcolor)
            bmp.SetPixel(bmpwidth / 4, i, gridcolor)
            bmp.SetPixel(bmpwidth * 3 / 4, i, gridcolor)
        Next

        Label5.Text = Format(-offy / magy, "0.0000")
        Label6.Text = Format((bmpheight / 2 - offy) / magy, "0.0000")
        Label7.Text = Format((bmpheight / 4 - offy) / magy, "0.0000")
        Label8.Text = Format((-bmpheight / 4 - offy) / magy, "0.0000")
        Label9.Text = Format((-bmpheight / 2 - offy) / magy, "0.0000")

        Label10.Text = Format(offx / magx, "0.0000")
        Label11.Text = Format((bmpwidth / 4 + offx) / magx * Xscale, "0.0000")
        Label12.Text = Format((bmpwidth / 2 + offx) / magx * Xscale, "0.0000")
        Label13.Text = Format((bmpwidth * 3 / 4 + offx) / magx * Xscale, "0.0000")
        Label14.Text = Format((bmpwidth + offx) / magx * Xscale, "0.0000")

        For count = 0 To 3
            If count = ps Then tmp = -1

            For i = 0 To dataset.Length / 4 - 1
                d = (i * magx) - offx
                If d >= 0 Then
                    dd = dataset(count, i) * -magy - offy + (bmpheight - 1) / 2
                End If
                If d >= 0 AndAlso d < bmpwidth - 1 AndAlso dd >= 0 AndAlso dd < bmpheight - 1 Then
                    bmp.SetPixel(d, dd, plotcolor(count))
                    If MouseDownX = d \ 1 AndAlso IsMarkerMoved AndAlso count = ps Then
                        tmp = d
                        tmp2 = dd
                        tmpval1 = i
                        tmpval2 = dataset(count, i)

                    End If
                End If
            Next
        Next

        Label15.Visible = False
        If tmp >= 0 Then
            Label15.Visible = True
            Label15.Text = Format(tmpval1 * Xscale, "0.0000") & "; " & Format(tmpval2, "0.0000")
            Label15.Location = New Point(tmp + 5, tmp2 + 5)
            For j = 0 To bmpwidth - 1
                bmp.SetPixel(j, tmp2, Color.Red)
            Next
            For j = 0 To bmpheight - 1
                bmp.SetPixel(tmp, j, Color.Red)
            Next
        End If

        e.Graphics.DrawImage(bmp, 0, 0)
    End Sub
    Private Sub AutoZoom(xy As Integer)
        If xy = 1 Or xy = 3 Then
            OffxWrite(0)
            MagxWrite((PanelDB1.Width / (dataset.Length / 4)))
        End If
        If xy = 2 Or xy = 3 Then
            Dim min As Double
            Dim max As Double
            For i = 0 To dataset.Length / 4 - 1
                If min > dataset(ps, i) Then min = dataset(ps, i)
                If max < dataset(ps, i) Then max = dataset(ps, i)
            Next

            MagyWrite(PanelDB1.Height / (max - min + 1))
            OffyWrite(-(max + min) / 2 * magy)
        End If
        Refresh()
    End Sub
    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PanelDB1.MouseDown
        If e.Button = MouseButtons.Left Then
            IsOffsetChanged = True
            MouseDownX = e.X + offx
            MouseDownY = e.Y + offy
        End If
        If e.Button = MouseButtons.Right Then
            IsMagnificationChanged = True
            MouseDownX = -e.X + magx * 1000
            MouseDownY = e.Y + magy * 1000
            MouseDownXo = e.X
        End If
    End Sub
    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PanelDB1.MouseUp
        If e.Button = MouseButtons.Left Then
            IsOffsetChanged = False
        End If
        If e.Button = MouseButtons.Right Then
            IsMagnificationChanged = False
        End If
    End Sub
    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PanelDB1.MouseMove
        If IsOffsetChanged Then
            IsMarkerMoved = False
            OffxWrite(-e.X + MouseDownX)
            OffyWrite(-e.Y + MouseDownY)
            Refresh()

        ElseIf IsMagnificationChanged Then
            IsMarkerMoved = False
            magx_s = magx
            MagxWrite((e.X + MouseDownX) / 1000)
            'MagyWrite((-e.Y + MouseDownY) / 1000)
            magx_s = magx / magx_s
            OffxWrite(((offx + MouseDownXo) * magx_s) - MouseDownXo)
            Refresh()
        Else
            IsMarkerMoved = True
            MouseDownX = e.X
            MouseDownY = e.Y
            Refresh()
        End If

    End Sub
    Private Sub MagxWrite(value As Double)
        If value > 1 Then value = 1
        If value <= 0 Then value = 0.001
        magx = Math.Round(value, 3)
    End Sub
    Private Sub MagyWrite(value As Double)
        If value < 0 Then value = 0
        magy = Math.Round(value, 3)
    End Sub
    Private Sub OffxWrite(value As Double)
        offx = value
    End Sub
    Private Sub OffyWrite(value As Double)
        offy = value
    End Sub
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        PanelDB1.Width = Width - 112
        PanelDB1.Height = Height - 176 - 32
        Label5.Location = (New Point(16, (PanelDB1.Bottom - PanelDB1.Top + 20) / 2))
        Label6.Location = (New Point(16, (PanelDB1.Top) - 6))
        Label9.Location = (New Point(16, (PanelDB1.Bottom - 7)))
        Label7.Location = (New Point(16, (Label5.Location.Y + Label6.Location.Y) / 2))
        Label8.Location = (New Point(16, (Label9.Location.Y + Label5.Location.Y) / 2))

        Label10.Location = (New Point(PanelDB1.Left, PanelDB1.Bottom + 2))
        Label12.Location = (New Point((PanelDB1.Right + PanelDB1.Left) / 2 - Label11.Width / 2, PanelDB1.Bottom + 2))
        Label14.Location = (New Point(PanelDB1.Right - Label14.Width, PanelDB1.Bottom + 2))
        Label11.Location = (New Point((Label10.Location.X + Label12.Location.X) / 2, PanelDB1.Bottom + 2))
        Label13.Location = (New Point((Label14.Location.X + Label12.Location.X) / 2, PanelDB1.Bottom + 2))

        GroupBox1.Location = New Point(80, Label13.Bottom + 5)
        PanelDB1.Refresh()
    End Sub
    Private Sub PanelDB1_MouseLeave(sender As Object, e As EventArgs) Handles PanelDB1.MouseLeave
        IsMarkerMoved = False
        Refresh()
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        ps = 0
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        ps = 1
    End Sub
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        ps = 2
    End Sub
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        ps = 3
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        AutoZoom(3)

    End Sub

    Private Sub PanelDB1_MouseWheel(sender As Object, e As MouseEventArgs) Handles PanelDB1.MouseWheel
        MagyWrite(magy * 1.05 ^ (e.Delta / SystemInformation.MouseWheelScrollDelta))
        Refresh()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        AutoZoom(1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        AutoZoom(2)
    End Sub

End Class
