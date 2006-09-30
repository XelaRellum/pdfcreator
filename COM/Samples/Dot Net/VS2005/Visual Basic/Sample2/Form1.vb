Imports System.Drawing
Imports System.Drawing.Printing

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Vom Windows Form Designer generierter Code "

    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private components As System.ComponentModel.IContainer

    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Button1 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(24, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(152, 40)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "&Start"
        '
        'Button2
        '
        Me.Button2.Enabled = False
        Me.Button2.Location = New System.Drawing.Point(288, 8)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(152, 40)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "&Preview"
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(192, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(0, 56)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBox1.Size = New System.Drawing.Size(464, 80)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "TextBox1"
        Me.TextBox1.WordWrap = False
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(466, 142)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.TextBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.Text = "Sample2 - PDFCreator COM interface"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private WithEvents _PDFCreator As PDFCreator.clsPDFCreator
    Private pErr As PDFCreator.clsPDFCreatorError

    Private pd As PrintDocument

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim parameters As String
        AddStatus("Status: Program is started.", True)

        pErr = New PDFCreator.clsPDFCreatorError
        _PDFCreator = New PDFCreator.clsPDFCreator

        parameters = "/NoProcessingAtStartup"

        If _PDFCreator.cStart(parameters) = True Then
            _PDFCreator.cClearCache()
            _PDFCreator.cOption("UseAutosave") = 0
            Button1.Enabled = True
            Button2.Enabled = True
            _PDFCreator.cPrinterStop = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        pd = New PrintDocument
        AddHandler pd.PrintPage, AddressOf pd_PrintPage
        pd.PrinterSettings.PrinterName = "PDFCreator"
        pd.DocumentName = "PDFCreator Dot Net - Sample2"
        pd.Print()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        pd = New PrintDocument
        AddHandler pd.PrintPage, AddressOf pd_PrintPage
        pd.PrinterSettings.PrinterName = "PDFCreator"
        pd.DocumentName = "PDFCreator Dot Net - Sample2"
        Dim ppdlg As New System.Windows.Forms.PrintPreviewDialog
        With ppdlg
            .Document = pd
            .WindowState = FormWindowState.Maximized
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    Private Sub pd_PrintPage(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        Dim x As Single, y As Single, r As Single
        With pd.PrinterSettings.DefaultPageSettings.PaperSize
            x = .Width / 2
            y = .Height / 2
            r = .Width / 4
        End With
        DrawCircles(x - r / 2, y - r / 2, r, 5, ev)
    End Sub

    Private Sub DrawCircles(ByVal x As Single, ByVal y As Single, ByVal r As Single, ByVal rec As Long, ByVal ev As PrintPageEventArgs)
        If rec <> 0 Then
            Dim p As New Pen(Color.Red)
            With ev.Graphics
                .DrawString("PDFCreator", New Font("Arial", 16), Brushes.Black, 100, 100, New StringFormat)
                .DrawEllipse(p, x - r, y, r, r)
                .DrawEllipse(p, x + r, y, r, r)
                .DrawEllipse(p, x, y - r, r, r)
                .DrawEllipse(p, x, y + r, r, r)
                .DrawEllipse(p, x, y, r, r)
            End With
            DrawCircles(x - r / 2, y - r / 2, r / 2, rec - 1, ev)
            DrawCircles(x - r / 2, y + r, r / 2, rec - 1, ev)
            DrawCircles(x + r, y - r / 2, r / 2, rec - 1, ev)
            DrawCircles(x + r, y + r, r / 2, rec - 1, ev)
        End If
    End Sub

    Private Sub _PDFCreator_eReady() Handles _PDFCreator.eReady
        AddStatus("Status: """ & _PDFCreator.cOutputFilename & """ was created!")
        _PDFCreator.cPrinterStop = True
    End Sub

    Private Sub _PDFCreator_eError() Handles _PDFCreator.eError
        pErr = _PDFCreator.cError
        AddStatus("Status: Error[" & pErr.Number & "]: " & pErr.Description)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
    End Sub

    Private Sub AddStatus(ByVal Str1 As String, Optional ByVal ClearStatus As Boolean = False)
        With TextBox1
            If ClearStatus = True Then
                .Text = Str1
                .SelectionStart = 0
            Else
                If .Text.Length = 0 Then
                    .Text = Str1
                    .SelectionStart = 0
                Else
                    .Text = .Text & vbCrLf & Str1
                End If
            End If
        End With
    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        _PDFCreator.cClose()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(_PDFCreator)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(pErr)
        pErr = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub
End Class