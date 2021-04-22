Imports System.Drawing.Graphics
Imports System.Drawing.Pen
Imports System.Drawing.Color
Imports System.Drawing.Brush
Imports System.Drawing.Point
Imports System.Drawing.Bitmap
Public Class Form1
    Protected m_timer As Timer
    Protected m_vertices(8) As Point3D
    Protected m_faces(6, 4) As Integer
    Protected m_colors(6) As Color
    Protected m_brushes(6) As Brush
    Protected m_angle As Integer
    Dim alpha, red, blue, green As Integer

    Private Sub Form1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        End
    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.WindowState = FormWindowState.Maximized

        Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)

        InitCube()

        m_timer = New Timer()

        m_timer.Interval = 25

        AddHandler m_timer.Tick, AddressOf AnimationLoop

        m_timer.Start()

        Dim ctimer As New Timer
        ctimer.Interval = 1000
        AddHandler ctimer.Tick, AddressOf ChangeColor
        ctimer.Start()
    End Sub

    Sub ChangeColor()
        InitCube()
    End Sub

    Private Sub InitCube()
        m_vertices = New Point3D() {
                     New Point3D(-1.2, 1.2, -1.2),
                     New Point3D(1.2, 1.2, -1.2),
                     New Point3D(1.2, -1.2, -1.2),
                     New Point3D(-1.2, -1.2, -1.2),
                     New Point3D(-1.2, 1.2, 1.2),
                     New Point3D(1.2, 1.2, 1.2),
                     New Point3D(1.2, -1.2, 1.2),
                     New Point3D(-1.2, -1.2, 1.2)}

        m_faces = New Integer(,) {{0, 1, 2, 3}, {1, 5, 6, 2}, {5, 4, 7, 6}, {4, 0, 3, 7}, {0, 4, 5, 1}, {3, 2, 6, 7}}

        'm_colors = New Color() {FromArgb(200, 0, 0, 255), FromArgb(200, 255, 0, 0), FromArgb(200, 0, 255, 0), FromArgb(200, 255, 0, 255), FromArgb(200, 0, 255, 255), FromArgb(200, 255, 255, 0)}

        Dim x As New Random()
        For i = 0 To 5
            alpha = x.Next(150, 200)
            red = x.Next(0, 255)
            green = x.Next(0, 255)
            blue = x.Next(0, 255)
            m_brushes(i) = New SolidBrush(FromArgb(alpha, red, green, blue))
        Next
    End Sub

    Private Sub AnimationLoop()
        Me.Invalidate()

        m_angle += 1
    End Sub

    Private Sub Form1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseEnter
        Me.Opacity = 0.1
    End Sub

    Private Sub Form1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseLeave
        Me.Opacity = 0.8
    End Sub

    Private Sub Main_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Dim t(8) As Point3D
        Dim f(4) As Integer
        Dim v As Point3D
        Dim avgZ(6) As Double
        Dim order(6) As Integer
        Dim tmp As Double
        Dim iMax As Integer

        e.Graphics.Clear(Color.Black)

        For i = 0 To 7
            Dim b As Brush = New SolidBrush(Color.White)
            v = m_vertices(i)
            t(i) = v.RotateX(m_angle).RotateY(m_angle).RotateZ(Me.m_angle)
            t(i) = t(i).Project(Me.ClientSize.Width, Me.ClientSize.Height, 256, 4)
        Next

        For i = 0 To 5
            avgZ(i) = (t(m_faces(i, 0)).Z + t(m_faces(i, 1)).Z + t(m_faces(i, 2)).Z + t(m_faces(i, 3)).Z) / 4.0
            order(i) = i
        Next

        For i = 0 To 4
            iMax = i
            For j = i + 1 To 5
                If avgZ(j) > avgZ(iMax) Then
                    iMax = j
                End If
            Next
            If iMax <> i Then
                tmp = avgZ(i)
                avgZ(i) = avgZ(iMax)
                avgZ(iMax) = tmp

                tmp = order(i)
                order(i) = order(iMax)
                order(iMax) = tmp
            End If
        Next

        For i = 0 To 5
            Dim points() As Point
            Dim index As Integer = order(i)
            points = New Point() {
                New Point(CInt(t(m_faces(index, 0)).X), CInt(t(m_faces(index, 0)).Y)),
                New Point(CInt(t(m_faces(index, 1)).X), CInt(t(m_faces(index, 1)).Y)),
                New Point(CInt(t(m_faces(index, 2)).X), CInt(t(m_faces(index, 2)).Y)),
                New Point(CInt(t(m_faces(index, 3)).X), CInt(t(m_faces(index, 3)).Y))
            }
            e.Graphics.FillPolygon(m_brushes(index), points)
        Next
    End Sub
End Class
