Imports System
Imports System.Net.NetworkInformation
Imports System.Windows.Forms
Public Class Form1
    Inherits System.Windows.Forms.Form
    Private Const URL_MESSAGE As String = "https://github.com/azul100/faucet"
    Private Const DIR_MESSAGE As String = "Enter directory to download to here"

    'DECLARE THIS WITHEVENTS SO WE GET EVENTS ABOUT DOWNLOAD PROGRESS
    Private WithEvents _Downloader As WebFileDownloader

#Region " Windows Form Designer generated code "

   

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer


    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txtURL As System.Windows.Forms.TextBox
    Friend WithEvents ProgBar As System.Windows.Forms.ProgressBar
    Friend WithEvents cmdDownload As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdGetFolder As System.Windows.Forms.Button
    Friend WithEvents txtDownloadTo As System.Windows.Forms.TextBox
    Friend WithEvents dlgFolderBrowse As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents lnkForums As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.txtURL = New System.Windows.Forms.TextBox()
        Me.ProgBar = New System.Windows.Forms.ProgressBar()
        Me.cmdDownload = New System.Windows.Forms.Button()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.cmdGetFolder = New System.Windows.Forms.Button()
        Me.txtDownloadTo = New System.Windows.Forms.TextBox()
        Me.dlgFolderBrowse = New System.Windows.Forms.FolderBrowserDialog()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.lnkForums = New System.Windows.Forms.LinkLabel()
        Me.lblUpload = New System.Windows.Forms.Label()
        Me.lblDownload = New System.Windows.Forms.Label()
        Me.label6 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.lblInterfaceType = New System.Windows.Forms.Label()
        Me.label5 = New System.Windows.Forms.Label()
        Me.lblBytesReceived = New System.Windows.Forms.Label()
        Me.lblBytesSent = New System.Windows.Forms.Label()
        Me.lblSpeed = New System.Windows.Forms.Label()
        Me.label4 = New System.Windows.Forms.Label()
        Me.label3 = New System.Windows.Forms.Label()
        Me.label2 = New System.Windows.Forms.Label()
        Me.cmbInterface = New System.Windows.Forms.ComboBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtURL
        '
        Me.txtURL.Location = New System.Drawing.Point(10, 93)
        Me.txtURL.Name = "txtURL"
        Me.txtURL.Size = New System.Drawing.Size(507, 20)
        Me.txtURL.TabIndex = 0
        Me.txtURL.Text = "https://github.com/azul100/faucet"
        '
        'ProgBar
        '
        Me.ProgBar.Location = New System.Drawing.Point(10, 157)
        Me.ProgBar.Name = "ProgBar"
        Me.ProgBar.Size = New System.Drawing.Size(507, 16)
        Me.ProgBar.TabIndex = 1
        '
        'cmdDownload
        '
        Me.cmdDownload.BackColor = System.Drawing.Color.White
        Me.cmdDownload.BackgroundImage = CType(resources.GetObject("cmdDownload.BackgroundImage"), System.Drawing.Image)
        Me.cmdDownload.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.cmdDownload.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdDownload.Image = CType(resources.GetObject("cmdDownload.Image"), System.Drawing.Image)
        Me.cmdDownload.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDownload.Location = New System.Drawing.Point(390, 331)
        Me.cmdDownload.Name = "cmdDownload"
        Me.cmdDownload.Size = New System.Drawing.Size(100, 26)
        Me.cmdDownload.TabIndex = 3
        Me.cmdDownload.Text = "Download"
        Me.cmdDownload.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdClose.Location = New System.Drawing.Point(485, 1)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(40, 24)
        Me.cmdClose.TabIndex = 4
        Me.cmdClose.Text = "Close"
        '
        'cmdGetFolder
        '
        Me.cmdGetFolder.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdGetFolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetFolder.Location = New System.Drawing.Point(485, 125)
        Me.cmdGetFolder.Name = "cmdGetFolder"
        Me.cmdGetFolder.Size = New System.Drawing.Size(32, 20)
        Me.cmdGetFolder.TabIndex = 2
        Me.cmdGetFolder.Text = "..."
        Me.cmdGetFolder.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'txtDownloadTo
        '
        Me.txtDownloadTo.Location = New System.Drawing.Point(10, 125)
        Me.txtDownloadTo.Name = "txtDownloadTo"
        Me.txtDownloadTo.Size = New System.Drawing.Size(469, 20)
        Me.txtDownloadTo.TabIndex = 1
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(10, 176)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(65, 21)
        Me.lblProgress.TabIndex = 8
        Me.lblProgress.Text = "#Progress"
        '
        'lnkForums
        '
        Me.lnkForums.LinkColor = System.Drawing.Color.Yellow
        Me.lnkForums.Location = New System.Drawing.Point(210, 58)
        Me.lnkForums.Name = "lnkForums"
        Me.lnkForums.Size = New System.Drawing.Size(203, 16)
        Me.lnkForums.TabIndex = 10
        Me.lnkForums.TabStop = True
        Me.lnkForums.Text = "https://github.com/azul100/faucet"
        '
        'lblUpload
        '
        Me.lblUpload.AutoSize = True
        Me.lblUpload.Location = New System.Drawing.Point(106, 344)
        Me.lblUpload.Name = "lblUpload"
        Me.lblUpload.Size = New System.Drawing.Size(13, 13)
        Me.lblUpload.TabIndex = 27
        Me.lblUpload.Text = "0"
        '
        'lblDownload
        '
        Me.lblDownload.AutoSize = True
        Me.lblDownload.Location = New System.Drawing.Point(106, 318)
        Me.lblDownload.Name = "lblDownload"
        Me.lblDownload.Size = New System.Drawing.Size(13, 13)
        Me.lblDownload.TabIndex = 26
        Me.lblDownload.Text = "0"
        '
        'label6
        '
        Me.label6.AutoSize = True
        Me.label6.Location = New System.Drawing.Point(12, 344)
        Me.label6.Name = "label6"
        Me.label6.Size = New System.Drawing.Size(41, 13)
        Me.label6.TabIndex = 25
        Me.label6.Text = "Upload"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(12, 318)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(58, 13)
        Me.label1.TabIndex = 24
        Me.label1.Text = "Download:"
        '
        'lblInterfaceType
        '
        Me.lblInterfaceType.AutoSize = True
        Me.lblInterfaceType.Location = New System.Drawing.Point(106, 214)
        Me.lblInterfaceType.Name = "lblInterfaceType"
        Me.lblInterfaceType.Size = New System.Drawing.Size(13, 13)
        Me.lblInterfaceType.TabIndex = 23
        Me.lblInterfaceType.Text = "0"
        '
        'label5
        '
        Me.label5.AutoSize = True
        Me.label5.Location = New System.Drawing.Point(12, 214)
        Me.label5.Name = "label5"
        Me.label5.Size = New System.Drawing.Size(79, 13)
        Me.label5.TabIndex = 22
        Me.label5.Text = "Interface Type:"
        '
        'lblBytesReceived
        '
        Me.lblBytesReceived.AutoSize = True
        Me.lblBytesReceived.Location = New System.Drawing.Point(106, 292)
        Me.lblBytesReceived.Name = "lblBytesReceived"
        Me.lblBytesReceived.Size = New System.Drawing.Size(13, 13)
        Me.lblBytesReceived.TabIndex = 21
        Me.lblBytesReceived.Text = "0"
        '
        'lblBytesSent
        '
        Me.lblBytesSent.AutoSize = True
        Me.lblBytesSent.Location = New System.Drawing.Point(106, 266)
        Me.lblBytesSent.Name = "lblBytesSent"
        Me.lblBytesSent.Size = New System.Drawing.Size(13, 13)
        Me.lblBytesSent.TabIndex = 20
        Me.lblBytesSent.Text = "0"
        '
        'lblSpeed
        '
        Me.lblSpeed.AutoSize = True
        Me.lblSpeed.Location = New System.Drawing.Point(106, 240)
        Me.lblSpeed.Name = "lblSpeed"
        Me.lblSpeed.Size = New System.Drawing.Size(13, 13)
        Me.lblSpeed.TabIndex = 19
        Me.lblSpeed.Text = "0"
        '
        'label4
        '
        Me.label4.AutoSize = True
        Me.label4.Location = New System.Drawing.Point(12, 292)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(85, 13)
        Me.label4.TabIndex = 18
        Me.label4.Text = "Bytes Received:"
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.Location = New System.Drawing.Point(12, 266)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(61, 13)
        Me.label3.TabIndex = 17
        Me.label3.Text = "Bytes Sent:"
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(12, 240)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(41, 13)
        Me.label2.TabIndex = 16
        Me.label2.Text = "Speed:"
        '
        'cmbInterface
        '
        Me.cmbInterface.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbInterface.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmbInterface.FormattingEnabled = True
        Me.cmbInterface.Location = New System.Drawing.Point(10, 12)
        Me.cmbInterface.Name = "cmbInterface"
        Me.cmbInterface.Size = New System.Drawing.Size(449, 21)
        Me.cmbInterface.TabIndex = 28
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.MidnightBlue
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(363, 310)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(154, 53)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 29
        Me.PictureBox1.TabStop = False
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Navy
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(529, 375)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmbInterface)
        Me.Controls.Add(Me.lblUpload)
        Me.Controls.Add(Me.lblDownload)
        Me.Controls.Add(Me.label6)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.lblInterfaceType)
        Me.Controls.Add(Me.label5)
        Me.Controls.Add(Me.lblBytesReceived)
        Me.Controls.Add(Me.lblBytesSent)
        Me.Controls.Add(Me.lblSpeed)
        Me.Controls.Add(Me.label4)
        Me.Controls.Add(Me.label3)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.lnkForums)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.cmdGetFolder)
        Me.Controls.Add(Me.txtDownloadTo)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdDownload)
        Me.Controls.Add(Me.ProgBar)
        Me.Controls.Add(Me.txtURL)
        Me.ForeColor = System.Drawing.Color.Yellow
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Download File "
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    'SUB MAIN WHERE WE ENABLE VISUAL STYLES, AND RUN MAIN FORM
    Public Shared Sub Main()
        Application.EnableVisualStyles()
        Application.DoEvents()
        Application.Run(New Form1)
        Application.Exit()
    End Sub

    'THESE ARE JUST FOR GUI TO MAKE IT LOOK A LITTLE MORE NICE
    Private Sub txtURL_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtURL.GotFocus
        If txtURL.Text = URL_MESSAGE Then txtURL.Text = String.Empty
    End Sub
    Private Sub txtURL_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtURL.LostFocus
        If txtURL.Text = String.Empty Then txtURL.Text = URL_MESSAGE
    End Sub
    Private Sub txtDownloadTo_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDownloadTo.GotFocus
        If txtDownloadTo.Text = DIR_MESSAGE Then txtDownloadTo.Text = String.Empty
    End Sub
    Private Sub txtDownloadTo_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDownloadTo.LostFocus
        If txtDownloadTo.Text = String.Empty Then txtDownloadTo.Text = DIR_MESSAGE
    End Sub

    'CLOSE PROGRAM
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub cmdDownload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDownload.Click

      

    End Sub

    'returns all text from last "/" in URL, but puts a "\" in front of the file name..
    Private Function GetFileNameFromURL(ByVal URL As String) As String
        If URL.IndexOf("/"c) = -1 Then Return String.Empty

        Return "\" & URL.Substring(URL.LastIndexOf("/"c) + 1)
    End Function

    'GET A FOLDER TO DOWNLOAD THE FILE TO
    Private Sub cmdGetFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetFolder.Click
        If dlgFolderBrowse.ShowDialog(Me) <> DialogResult.Cancel Then
            txtDownloadTo.Text = dlgFolderBrowse.SelectedPath
        End If
    End Sub

    'FIRES WHEN WE HAVE GOTTEN THE DOWNLOAD SIZE, SO WE KNOW WHAT BOUNDS TO GIVE THE PROGRESS BAR
    Private Sub _Downloader_FileDownloadSizeObtained(ByVal iFileSize As Long) Handles _Downloader.FileDownloadSizeObtained
        ProgBar.Value = 0
        ProgBar.Maximum = Convert.ToInt32(iFileSize)
    End Sub

    'FIRES WHEN DOWNLOAD IS COMPLETE
    Private Sub _Downloader_FileDownloadComplete() Handles _Downloader.FileDownloadComplete
        ProgBar.Value = ProgBar.Maximum
        MessageBox.Show("File Download Complete")
    End Sub

    'FIRES WHEN DOWNLOAD FAILES. PASSES IN EXCEPTION INFO
    Private Sub _Downloader_FileDownloadFailed(ByVal ex As System.Exception) Handles _Downloader.FileDownloadFailed
        MessageBox.Show("An error has occured during download: " & ex.Message)
    End Sub

    'FIRES WHEN MORE OF THE FILE HAS BEEN DOWNLOADED
    Private Sub _Downloader_AmountDownloadedChanged(ByVal iNewProgress As Long) Handles _Downloader.AmountDownloadedChanged
        ProgBar.Value = Convert.ToInt32(iNewProgress)
        lblProgress.Text = WebFileDownloader.FormatFileSize(iNewProgress) & " of " & WebFileDownloader.FormatFileSize(ProgBar.Maximum) & " downloaded"
        Application.DoEvents()
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblProgress.Text = String.Empty
        txtURL.Text = URL_MESSAGE
        txtDownloadTo.Text = DIR_MESSAGE
        Show()
        cmdGetFolder.Focus()
    End Sub
    Private Const timerUpdate As Double = 1000
    Private nicArr As NetworkInterface()
    Private timer As Timer
    Public Sub New()
        InitializeComponent()
        InitializeNetworkInterface()
        InitializeTimer()
    End Sub
    Private Sub InitializeNetworkInterface()
        nicArr = NetworkInterface.GetAllNetworkInterfaces()
        Dim i As Integer = 0
        While i < nicArr.Length
            cmbInterface.Items.Add(nicArr(i).Name)
            System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
        End While
        cmbInterface.SelectedIndex = 0
    End Sub
    Private Sub InitializeTimer()
        timer = New Timer()
        timer.Interval = DirectCast(timerUpdate.GetTypeCode, Integer)

        timer.Start()
    End Sub
    Private Sub UpdateNetworkInterface()
        Dim nic As NetworkInterface = nicArr(cmbInterface.SelectedIndex)
        Dim interfaceStats As IPv4InterfaceStatistics = nic.GetIPv4Statistics()
        Dim bytesSentSpeed As Integer = DirectCast((Uri.UriSchemeHttps.Clone(interfaceStats.BytesSent - Double.Parse(lblBytesSent.Text))), Integer) / 1024
        Dim bytesReceivedSpeed As Integer = DirectCast((Uri.UriSchemeHttps.Clone(interfaceStats.BytesReceived - Double.Parse(lblBytesReceived.Text))), Integer) / 1024
        lblSpeed.Text = nic.Speed.ToString()
        lblInterfaceType.Text = nic.NetworkInterfaceType.ToString()
        lblSpeed.Text = nic.Speed.ToString()
        lblBytesReceived.Text = interfaceStats.BytesReceived.ToString()
        lblBytesSent.Text = interfaceStats.BytesSent.ToString()
        lblUpload.Text = bytesSentSpeed.ToString() + " KB/s"
        lblDownload.Text = bytesReceivedSpeed.ToString() + " KB/s"
    End Sub
    Sub timer_Tick(sender As Object, e As EventArgs)
        UpdateNetworkInterface()
    End Sub

   

    Private Sub lnkForums_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkForums.LinkClicked
        Process.Start("https://github.com/azul100/faucet")
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        'VERIFY A DIRECTORY WAS PICKED AND THAT IT EXISTS
        If Not IO.Directory.Exists(txtDownloadTo.Text) Then
            MessageBox.Show("Not a valid directory to download to, please pick a valid directory", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        'DO THE DOWNLOAD
        Try
            _Downloader = New WebFileDownloader
            _Downloader.DownloadFileWithProgress(txtURL.Text, txtDownloadTo.Text.TrimEnd("\"c) & GetFileNameFromURL(txtURL.Text))
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub
End Class
