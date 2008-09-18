Imports System.IO

Public Class Form1
    Inherits System.Windows.Forms.Form

    Private Const maxTime As Long = 20

    Private WithEvents _PDFCreator As PDFCreator.clsPDFCreator
    Private pErr As PDFCreator.clsPDFCreatorError

    Private ReadyState As Boolean

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

    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SuspendLayout()
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 118)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Size = New System.Drawing.Size(536, 16)
        Me.StatusBar1.TabIndex = 2
        Me.StatusBar1.Text = "Status:"
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(8, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(152, 40)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Show options dialog"
        '
        'Button2
        '
        Me.Button2.Enabled = False
        Me.Button2.Location = New System.Drawing.Point(8, 56)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(152, 40)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Show logfile dialog"
        '
        'Button3
        '
        Me.Button3.Enabled = False
        Me.Button3.Location = New System.Drawing.Point(192, 8)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(152, 40)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Print printer testpage"
        '
        'Button4
        '
        Me.Button4.Enabled = False
        Me.Button4.Location = New System.Drawing.Point(192, 56)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(152, 40)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "Print PDFCreator testpage"
        '
        'Button5
        '
        Me.Button5.Enabled = False
        Me.Button5.Location = New System.Drawing.Point(376, 8)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(152, 40)
        Me.Button5.TabIndex = 5
        Me.Button5.Text = "Convert to PDF"
        '
        'Button6
        '
        Me.Button6.Enabled = False
        Me.Button6.Location = New System.Drawing.Point(376, 56)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(152, 40)
        Me.Button6.TabIndex = 6
        Me.Button6.Text = "Convert to TIFF"
        '
        'Timer1
        '
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(536, 134)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.StatusBar1)
        Me.Name = "Form1"
        Me.Text = "Sample1 - PDFCreator COM interface"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim parameters As String
        StatusBar1.Text = "Status: Program is started."

        pErr = New PDFCreator.clsPDFCreatorError
        _PDFCreator = New PDFCreator.clsPDFCreator

        parameters = "/NoProcessingAtStartup"

        If _PDFCreator.cStart(parameters) = False Then
            StatusBar1.Text = "Status: Error[" & pErr.Number & "]: " & pErr.Description
        Else
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
        End If
    End Sub

    Private Sub PDFCreator_Ready() Handles _PDFCreator.eReady
        StatusBar1.Text = "Status: """ & _PDFCreator.cOutputFilename & """ was created!"
        _PDFCreator.cPrinterStop = True
        ReadyState = True
    End Sub

    Private Sub _PDFCreator_eError() Handles _PDFCreator.eError
        pErr = _PDFCreator.cError
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        _PDFCreator.cShowOptionsDialog(True)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        _PDFCreator.cShowLogfileDialog(True)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        _PDFCreator.cDefaultPrinter = "PDFCreator"
        ' Wait 1 second
        Timer1.Interval = 1000
        Timer1.Enabled = True
        Do While Timer1.Enabled
            Application.DoEvents()
        Loop
        _PDFCreator.cPrinterStop = False
        _PDFCreator.cPrintPrinterTestpage()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        _PDFCreator.cPrintPDFCreatorTestpage()
        _PDFCreator.cPrinterStop = False
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        PrintIt(0)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        PrintIt(5)
    End Sub

    Private Sub PrintIt(ByVal Filetyp As Long)
        Dim fname As String, fi As FileInfo, DefaultPrinter As String
        Dim opt As PDFCreator.clsPDFCreatorOptions
        With OpenFileDialog1
            .Multiselect = False
            .CheckFileExists = True
            .CheckPathExists = True
            .Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        End With
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            fi = New FileInfo(OpenFileDialog1.FileName)
            If fi.Name.Length > 0 Then
                If InStr(fi.Name, ".", CompareMethod.Text) > 1 Then
                    fname = Mid(fi.Name, 1, InStr(fi.Name, ".", CompareMethod.Text) - 1)
                Else
                    fname = fi.Name
                End If
            End If
            If Not _PDFCreator.cIsPrintable(fi.FullName) Then
                MsgBox("File '" & fi.FullName & "' is not printable!", MsgBoxStyle.Exclamation, Me.Text)
                Exit Sub
            End If
            opt = _PDFCreator.cOptions
            With opt
                .UseAutosave = 1
                .UseAutosaveDirectory = 1
                .AutosaveDirectory = fi.DirectoryName
                .AutosaveFormat = Filetyp
                If Filetyp = 5 Then
                    .BitmapResolution = 72
                End If
                opt.AutosaveFilename = fname
            End With
            With _PDFCreator
                .cOptions = opt
                .cClearCache()
                DefaultPrinter = .cDefaultPrinter
                .cDefaultPrinter = "PDFCreator"
                .cPrintFile(fi.FullName)
                ReadyState = False
                .cPrinterStop = False
            End With

            With Timer1
                .Interval = maxTime * 1000
                .Enabled = True
                Do While Not ReadyState And .Enabled
                    Application.DoEvents()
                Loop
                .Enabled = False
            End With
            If Not ReadyState Then
                MsgBox("Creating printer test page as pdf." & vbCrLf & vbCrLf & _
                 "An error is occured: Time is up!", MsgBoxStyle.Exclamation, Me.Text)
            End If
            _PDFCreator.cPrinterStop = True
            _PDFCreator.cDefaultPrinter = DefaultPrinter
        End If
        opt = Nothing
    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        _PDFCreator.cClose()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(_PDFCreator)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(pErr)
        pErr = Nothing
        _PDFCreator = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub
End Class