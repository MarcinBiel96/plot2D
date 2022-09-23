
Public Class Form1
    Private dataset(,) As Double
    ReadOnly gridcolor As Color = Color.FromArgb(30, 30, 30)
    ReadOnly plotcolor(3) As Color
    ReadOnly PlotEnabled(3) As Boolean
    Private DANE(,) As Short
    Private Xscale As Double = 1
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
    Private linksamples As Boolean
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        plotcolor(0) = Color.Lime
        plotcolor(1) = Color.Yellow
        plotcolor(2) = Color.Blue
        plotcolor(3) = Color.Magenta
        RadioButton1.ForeColor = plotcolor(0)
        RadioButton2.ForeColor = plotcolor(1)
        RadioButton3.ForeColor = plotcolor(2)
        RadioButton4.ForeColor = plotcolor(3)
        ComboBox1.ForeColor = plotcolor(0)
        ComboBox2.ForeColor = plotcolor(1)
        ComboBox3.ForeColor = plotcolor(2)
        ComboBox4.ForeColor = plotcolor(3)
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        ComboBox4.SelectedIndex = 0
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

        Label10.Text = Format(offx / magx * Xscale, "0.0000")
        Label11.Text = Format((bmpwidth / 4 + offx) / magx * Xscale, "0.0000")
        Label12.Text = Format((bmpwidth / 2 + offx) / magx * Xscale, "0.0000")
        Label13.Text = Format((bmpwidth * 3 / 4 + offx) / magx * Xscale, "0.0000")
        Label14.Text = Format((bmpwidth + offx) / magx * Xscale, "0.0000")

        Dim ff As Double
        Dim f As Double
        For count = 0 To 3
            If PlotEnabled(count) Then
                If count = ps Then tmp = -1

                For i = 0 To dataset.Length / 4 - 1
                    d = (i * magx) - offx
                    If d >= 0 Then
                        dd = dataset(count, i) * -magy - offy + (bmpheight - 1) / 2
                    End If
                    If d >= 0 AndAlso d < bmpwidth - 1 AndAlso dd >= 0 AndAlso dd < bmpheight - 1 Then
                        bmp.SetPixel(d, dd, plotcolor(count))
                        If linksamples Then
                            If dd > ff Then
                                For z = ff To dd
                                    If z < bmpheight - 1 AndAlso z > 0 AndAlso i >= 1 Then
                                        bmp.SetPixel(d, z, plotcolor(count))
                                    End If
                                Next
                            ElseIf dd < ff Then
                                For z = dd To ff
                                    If z < bmpheight - 1 AndAlso z > 0 AndAlso i >= 1 Then
                                        bmp.SetPixel(d, z, plotcolor(count))
                                    End If
                                Next
                            End If
                        End If
                        If MouseDownX = d \ 1 AndAlso IsMarkerMoved AndAlso count = ps Then
                                tmp = d
                                tmp2 = dd
                                tmpval1 = i
                                tmpval2 = dataset(count, i)

                            End If
                        End If
                        ff = dd
                    f = d
                Next
            End If
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
        PanelDB1.Height = Height - 176 - 32 - 16
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
    Sub Mbj(filename As String)
        Dim Sign(31) As Byte
        Dim Format As Short
        Dim RozmiarNagl As Short
        Dim Ext(3) As Byte
        Dim tihour As Byte
        Dim timin As Byte
        Dim tisec As Byte
        Dim tihund As Short
        Dim dayear As Short
        Dim daday As Byte
        Dim damon As Byte
        Dim CardSampleFreq As Double
        Dim kanMax As Byte
        Dim ACBits As Short

        Dim WspX() As Double
        Dim WspY() As Double
        Dim WspZ() As Double
        Dim TypPrzetw() As Byte
        Dim Skladowa() As Byte
        Dim WzmocnienieToru() As Double
        Dim CzulSejKsd() As Double
        Dim TlumienieD() As Double
        Dim EngAK() As Double

        Dim EngK() As Double
        Dim EngX() As Double
        Dim EngA() As Double
        Dim EngB() As Double

        Dim rezerwa1() As Double
        Dim rezerwa2() As Double
        Dim rezerwa3() As Double
        Dim rezerwa4() As Double
        Dim rezerwa5(67) As Byte

        Dim kanal_export() As Short

        Dim Dlugosc As Long

        FileOpen(1, filename, OpenMode.Binary, OpenAccess.Read)
        Dlugosc = LOF(1)

        FileGet(1, Sign)
        FileGet(1, Format)
        FileGet(1, RozmiarNagl)
        FileGet(1, Ext)
        FileGet(1, tihour)
        FileGet(1, timin)
        FileGet(1, tisec)
        FileGet(1, tihund)
        FileGet(1, dayear)
        FileGet(1, daday)
        FileGet(1, damon)
        FileGet(1, CardSampleFreq)
        FileGet(1, kanMax)
        FileGet(1, ACBits)

        ReDim WspX(kanMax - 1)
        ReDim WspY(kanMax - 1)
        ReDim WspZ(kanMax - 1)
        ReDim TypPrzetw(kanMax - 1)
        ReDim Skladowa(kanMax - 1)
        ReDim WzmocnienieToru(kanMax - 1)
        ReDim CzulSejKsd(kanMax - 1)
        ReDim TlumienieD(kanMax - 1)
        ReDim EngAK(kanMax - 1)
        ReDim EngK(kanMax - 1)
        ReDim EngX(kanMax - 1)
        ReDim EngA(kanMax - 1)
        ReDim EngB(kanMax - 1)
        ReDim rezerwa1(kanMax - 1)
        ReDim rezerwa2(kanMax - 1)
        ReDim rezerwa3(kanMax - 1)
        ReDim rezerwa4(kanMax - 1)

        FileGet(1, WspX)
        FileGet(1, WspY)
        FileGet(1, WspZ)
        FileGet(1, TypPrzetw)
        FileGet(1, Skladowa)
        FileGet(1, WzmocnienieToru)
        FileGet(1, CzulSejKsd)
        FileGet(1, TlumienieD)
        FileGet(1, EngAK)
        FileGet(1, EngK)
        FileGet(1, EngX)
        FileGet(1, EngA)
        FileGet(1, EngB)
        FileGet(1, rezerwa1)
        FileGet(1, rezerwa2)
        FileGet(1, rezerwa3)
        FileGet(1, rezerwa4)

        Xscale = 1 / CardSampleFreq
        Dlugosc = (Dlugosc - RozmiarNagl) / 128
        ReDim DANE(kanMax - 1, Dlugosc - 1)

        Seek(1, RozmiarNagl + 1)
        FileGet(1, DANE)
        FileClose(1)

        ReDim kanal_export(Dlugosc - 1)
        For i = 0 To Dlugosc - 1
            kanal_export(i) = DANE(32 - 1, i)
        Next i
    End Sub
    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Mbj(OpenFileDialog1.FileNames(0))
        ReDim dataset(3, DANE.Length / 64 - 1)
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        ComboBox4.SelectedIndex = 0
        ComboBox1.SelectedIndex = 1
        AutoZoom(3)
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        OpenFileDialog1.ShowDialog()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            PlotEnabled(0) = False
            RadioButton1.Enabled = False
        Else
            PlotEnabled(0) = True
            RadioButton1.Enabled = True
            For i = 0 To DANE.Length / 64 - 1
                dataset(0, i) = DANE(ComboBox1.SelectedIndex - 1, i) / 65536 * 20
            Next i
        End If
        Refresh()
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex = 0 Then
            PlotEnabled(1) = False
            RadioButton2.Enabled = False
        Else
            PlotEnabled(1) = True
            RadioButton2.Enabled = True
            For i = 0 To DANE.Length / 64 - 1
                dataset(1, i) = DANE(ComboBox2.SelectedIndex - 1, i) / 65536 * 20
            Next i
        End If
        Refresh()
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.SelectedIndex = 0 Then
            PlotEnabled(2) = False
            RadioButton3.Enabled = False
        Else
            PlotEnabled(2) = True
            RadioButton3.Enabled = True
            For i = 0 To DANE.Length / 64 - 1
                dataset(2, i) = DANE(ComboBox3.SelectedIndex - 1, i) / 65536 * 20
            Next i
        End If
        Refresh()
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox4.SelectedIndex = 0 Then
            PlotEnabled(3) = False
            RadioButton4.Enabled = False
        Else
            PlotEnabled(3) = True
            RadioButton4.Enabled = True
            For i = 0 To DANE.Length / 64 - 1
                dataset(3, i) = DANE(ComboBox4.SelectedIndex - 1, i) / 65536 * 20
            Next i
        End If
        Refresh()
    End Sub
    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        If files.Length = 1 Then
            If files(0).Contains(".mbj") Then
                Mbj(files(0))
                ReDim dataset(3, DANE.Length / 64 - 1)
                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                ComboBox3.Enabled = True
                ComboBox4.Enabled = True
                ComboBox1.SelectedIndex = 0
                ComboBox2.SelectedIndex = 0
                ComboBox3.SelectedIndex = 0
                ComboBox4.SelectedIndex = 0
                ComboBox1.SelectedIndex = 1
                AutoZoom(3)
            End If
        End If

    End Sub
    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        linksamples = CheckBox1.Checked
        Refresh()
    End Sub
End Class
