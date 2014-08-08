
Imports System.IO

Public Class Main_Screen
    Dim savefilescheck As Boolean = True

    Private AutoUpdate As Boolean = False

    Private FAQsCount As Integer = 0
    Private CoverArtCount As Integer = 0
    Private GamesCount As Integer = 0

    Private GameFAQs As Integer = 0
    Private GameImages As Integer = 0
    Private GameSecrets As Integer = 0
    Private GameReviews As Integer = 0
    Private GameFAQsS As String = ""
    Private GameImagesS As String = ""
    Private GameSecretsS As String = ""
    Private GameReviewsS As String = ""



    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\Sounds\UHOH.WAV").Replace("\\", "\")) = True Then
                My.Computer.Audio.Play((Application.StartupPath & "\Sounds\UHOH.WAV").Replace("\\", "\"), AudioPlayMode.Background)
            End If
            Dim Display_Message1 As New Display_Message()
            Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.Message.ToString
            Display_Message1.Timer1.Interval = 1000
            Display_Message1.ShowDialog()
            Display_Message1.Dispose()
            Display_Message1 = Nothing
            If My.Computer.FileSystem.DirectoryExists((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs") = False Then
                My.Computer.FileSystem.CreateDirectory((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
            End If
            Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ": " & ex.ToString)
            filewriter.Flush()
            filewriter.Close()
            filewriter = Nothing
            ex = Nothing
            identifier_msg = Nothing
            If WebBrowser1.Url.Equals((Application.StartupPath & "\Loading.html").Replace("\\", "\")) Then
                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\LoadFailed.html").Replace("\\", "\")) Then
                    WebBrowser1.Navigate((Application.StartupPath & "\LoadFailed.html").Replace("\\", "\"))
                End If

            End If


        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub RunWorker()
        Try
            Controls_Enabler(False)
            BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            Error_Handler(ex, "Run Worker")
        End Try
    End Sub

    Private Sub LoadFAQs_List()
        Try
            ListBox1.Items.Clear()
            ListBox1.Text = ""
            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\No Cover Art.jpg").Replace("\\", "\")) = True Then
                PictureBox2.Image = Image.FromFile((Application.StartupPath & "\No Cover Art.jpg").Replace("\\", "\"))
                PictureBox3.Image = Image.FromFile((Application.StartupPath & "\No Cover Art.jpg").Replace("\\", "\"))
            Else
                PictureBox2.Image = ImageList1.Images.Item(0)
                PictureBox3.Image = ImageList1.Images.Item(0)
            End If
            Dim dinfo As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\FAQs").Replace("\\", "\"))
            If dinfo.Exists Then
                For Each subdir As DirectoryInfo In dinfo.GetDirectories
                    If Not subdir.Name = "Error Logs" Then
                        Label_Info.Text = FAQsCount & " FAQs, " & CoverArtCount & " Cover Images Located"
                        GamesCount = GamesCount + 1
                        If GamesCount <> 1 Then
                            Label_GameCount.Text = GamesCount & " Games Located"
                        Else
                            Label_GameCount.Text = GamesCount & " Game Located"
                        End If
                        ListBox1.Items.Add(subdir.Name.Replace(" -", ":"))
                        '----------------------------
                        Dim name As String = subdir.Name.Replace(":", " -")
                        If My.Computer.FileSystem.FileExists((subdir.FullName & "\" & name & ".jpg").Replace("\\", "\")) Then
                            CoverArtCount = CoverArtCount + 1
                        End If
                        For Each finfo As FileInfo In subdir.GetFiles("*.txt", SearchOption.TopDirectoryOnly)
                            FAQsCount = FAQsCount + 1
                            finfo = Nothing
                        Next
                        '----------------------------
                        Label_Info.Text = FAQsCount & " FAQs, " & CoverArtCount & " Cover Images Located"
                    End If
                    subdir = Nothing
                Next
            Else
                My.Computer.FileSystem.CreateDirectory((Application.StartupPath & "\FAQs").Replace("\\", "\"))
            End If
            dinfo = Nothing
            If ListBox1.Items.Count > 0 Then
                CoverPanel1.Visible = False
                CoverPanel2.Visible = False
                CoverPanel3.Visible = False
                ListBox1.SelectedIndex = 0
            Else
                CoverPanel1.Visible = True
                CoverPanel2.Visible = True
                CoverPanel3.Visible = True
            End If
            
            ListBox1.Select()
            ListBox1.Focus()
        Catch ex As Exception
            Error_Handler(ex, "Load FAQ Directories")
        End Try
    End Sub


    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.Text = My.Application.Info.ProductName & " " & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ""
            Label2.Text = "Build " & My.Application.Info.Version.Major & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & My.Application.Info.Version.Revision
            Timer1.Start()
        Catch ex As Exception
            Error_Handler(ex, "Application Load")
        End Try
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Try

            If ListBox1.SelectedIndex <> -1 Then
                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\No Cover Art.jpg").Replace("\\", "\")) = True Then
                    PictureBox2.Image = Image.FromFile((Application.StartupPath & "\No Cover Art.jpg").Replace("\\", "\"))
                    PictureBox3.Image = Image.FromFile((Application.StartupPath & "\No Cover Art.jpg").Replace("\\", "\"))
                Else
                    PictureBox2.Image = ImageList1.Images.Item(0)
                    PictureBox3.Image = ImageList1.Images.Item(0)
                End If


                GameFAQs = 0
                GameImages = 0
                GameSecrets = 0
                GameReviews = 0
                GameFAQsS = ""
                GameImagesS = ""
                GameSecretsS = ""
                GameReviewsS = ""
                If GameFAQs = 1 Then
                    GameFAQsS = ""
                Else
                    GameFAQsS = "s"
                End If
                If GameImages = 1 Then
                    GameImagesS = ""
                Else
                    GameImagesS = "s"
                End If
                If GameSecrets = 1 Then
                    GameSecretsS = ""
                Else
                    GameSecretsS = "s"
                End If
                If GameReviews = 1 Then
                    GameReviewsS = ""
                Else
                    GameReviewsS = "s"
                End If
                Label_GameInfo.Text = GameFAQs & " FAQ" & GameFAQsS & " | " & GameReviews & " Review" & GameReviewsS & " | " & GameSecrets & " Secret" & GameSecretsS & " | " & GameImages & " Image" & GameImagesS


                ComboBox1.Items.Clear()
                ComboBox1.Text = ""
                WebBrowser1.Navigate((Application.StartupPath & "\No-FAQs.html").Replace("\\", "\"))
                Label1.Text = ""

                Dim name As String = ListBox1.Items.Item(ListBox1.SelectedIndex)
                name = name.Replace(":", " -")
                Label1.Text = name.Replace(" -", ":")
                LinkLabel1.Tag = Label1.Text.Replace("-", " ").Replace(";", " ").Replace("  ", " ")
                LinkLabel2.Tag = Label1.Text.Replace("-", " ").Replace(";", " ").Replace("  ", " ")
                Dim dinfo As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\FAQs\" & name).Replace("\\", "\"))
                If dinfo.Exists Then
                    If My.Computer.FileSystem.FileExists((dinfo.FullName & "\" & name & ".jpg").Replace("\\", "\")) Then
                        PictureBox2.Image = Image.FromFile((dinfo.FullName & "\" & name & ".jpg").Replace("\\", "\"))
                        PictureBox3.Image = Image.FromFile((dinfo.FullName & "\" & name & ".jpg").Replace("\\", "\"))
                    End If
                    ComboBox1.Items.Add("-- Select a Document to Display --")
                    For Each finfo As FileInfo In dinfo.GetFiles("*.txt", SearchOption.TopDirectoryOnly)
                        GameFAQs = GameFAQs + 1
                        If GameFAQs = 1 Then
                            GameFAQsS = ""
                        Else
                            GameFAQsS = "s"
                        End If
                        If GameImages = 1 Then
                            GameImagesS = ""
                        Else
                            GameImagesS = "s"
                        End If
                        If GameSecrets = 1 Then
                            GameSecretsS = ""
                        Else
                            GameSecretsS = "s"
                        End If
                        If GameReviews = 1 Then
                            GameReviewsS = ""
                        Else
                            GameReviewsS = "s"
                        End If
                        Label_GameInfo.Text = GameFAQs & " FAQ" & GameFAQsS & " | " & GameReviews & " Review" & GameReviewsS & " | " & GameSecrets & " Secret" & GameSecretsS & " | " & GameImages & " Image" & GameImagesS
                        ComboBox1.Items.Add(finfo.Name)
                        finfo = Nothing
                    Next
                    For Each finfo As FileInfo In dinfo.GetFiles("*.mht", SearchOption.TopDirectoryOnly)
                        ComboBox1.Items.Add(finfo.Name)
                        finfo = Nothing
                    Next
                    For Each finfo As FileInfo In dinfo.GetFiles("*.htm", SearchOption.TopDirectoryOnly)
                        ComboBox1.Items.Add(finfo.Name)
                        finfo = Nothing
                    Next
                    For Each finfo As FileInfo In dinfo.GetFiles()
                        If finfo.Name.ToLower.IndexOf("secrets") <> -1 And Not finfo.Extension.ToLower = ".txt" Then
                            GameSecrets = GameSecrets + 1
                            If GameFAQs = 1 Then
                                GameFAQsS = ""
                            Else
                                GameFAQsS = "s"
                            End If
                            If GameImages = 1 Then
                                GameImagesS = ""
                            Else
                                GameImagesS = "s"
                            End If
                            If GameSecrets = 1 Then
                                GameSecretsS = ""
                            Else
                                GameSecretsS = "s"
                            End If
                            If GameReviews = 1 Then
                                GameReviewsS = ""
                            Else
                                GameReviewsS = "s"
                            End If
                            Label_GameInfo.Text = GameFAQs & " FAQ" & GameFAQsS & " | " & GameReviews & " Review" & GameReviewsS & " | " & GameSecrets & " Secret" & GameSecretsS & " | " & GameImages & " Image" & GameImagesS
                        End If
                        If finfo.Name.ToLower.IndexOf("review") <> -1 And Not finfo.Extension.ToLower = ".txt" Then
                            GameReviews = GameReviews + 1
                            If GameFAQs = 1 Then
                                GameFAQsS = ""
                            Else
                                GameFAQsS = "s"
                            End If
                            If GameImages = 1 Then
                                GameImagesS = ""
                            Else
                                GameImagesS = "s"
                            End If
                            If GameSecrets = 1 Then
                                GameSecretsS = ""
                            Else
                                GameSecretsS = "s"
                            End If
                            If GameReviews = 1 Then
                                GameReviewsS = ""
                            Else
                                GameReviewsS = "s"
                            End If
                            Label_GameInfo.Text = GameFAQs & " FAQ" & GameFAQsS & " | " & GameReviews & " Review" & GameReviewsS & " | " & GameSecrets & " Secret" & GameSecretsS & " | " & GameImages & " Image" & GameImagesS
                        End If
                        If ComboBox1.Items.IndexOf(finfo.Name) = -1 Then
                            If Not finfo.Name.ToLower = "thumbs.db" And Not finfo.Name.ToLower = "desktop.ini" Then
                                If finfo.Extension.ToLower = ".jpg" Or finfo.Extension.ToLower = ".bmp" Or finfo.Extension.ToLower = ".png" Or finfo.Extension.ToLower = ".gif" Then
                                    GameImages = GameImages + 1
                                    If GameFAQs = 1 Then
                                        GameFAQsS = ""
                                    Else
                                        GameFAQsS = "s"
                                    End If
                                    If GameImages = 1 Then
                                        GameImagesS = ""
                                    Else
                                        GameImagesS = "s"
                                    End If
                                    If GameSecrets = 1 Then
                                        GameSecretsS = ""
                                    Else
                                        GameSecretsS = "s"
                                    End If
                                    If GameReviews = 1 Then
                                        GameReviewsS = ""
                                    Else
                                        GameReviewsS = "s"
                                    End If
                                    Label_GameInfo.Text = GameFAQs & " FAQ" & GameFAQsS & " | " & GameReviews & " Review" & GameReviewsS & " | " & GameSecrets & " Secret" & GameSecretsS & " | " & GameImages & " Image" & GameImagesS
                                End If
                                ComboBox1.Items.Add(finfo.Name)
                            End If
                        End If
                        finfo = Nothing
                    Next
                End If
                If ComboBox1.Items.Count > 0 Then
                    ComboBox1.SelectedIndex = 0
                    ComboBox1.Select(0, 0)
                End If
                dinfo = Nothing
            End If
        Catch ex As Exception
            Error_Handler(ex, "ListBox1 SelectedIndexChanged")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If ComboBox1.SelectedIndex <> -1 Then
                If ComboBox1.SelectedIndex <> 0 Then
                    Panel2.Visible = False
                    Dim dname As String = ListBox1.Items.Item(ListBox1.SelectedIndex).replace(":", " -")
                    Dim name As String = ComboBox1.Items.Item(ComboBox1.SelectedIndex).replace(":", " -")
                    If My.Computer.FileSystem.FileExists((Application.StartupPath & "\FAQs\" & dname & "\" & name).Replace("\\", "\")) Then
                        WebBrowser1.Navigate((Application.StartupPath & "\FAQs\" & dname & "\" & name).Replace("\\", "\"))
                    End If
                Else
                    Panel2.Visible = True
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Load FAQ")
        End Try
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        Try
            HelpBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        Try
            AboutBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            LoadFAQs_List()
        Catch ex As Exception
            Error_Handler(ex, "Background Worker Do Work")
        End Try
    End Sub

    Private Sub Controls_Enabler(ByVal enable As Boolean)
        If enable = True Then
            Me.ControlBox = True
            ComboBox1.Enabled = True
            ListBox1.Enabled = True
            Label7.Enabled = True
            Label5.Enabled = True
            Label4.Enabled = True
            Panel3.Visible = False
        Else
            Me.ControlBox = False
            ComboBox1.Enabled = False
            ListBox1.Enabled = False
            Label7.Enabled = False
            Label5.Enabled = False
            Label4.Enabled = False
            Panel3.Visible = True
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Controls_Enabler(True)
            If ComboBox1.Items.Count > 0 Then
                ComboBox1.SelectedIndex = 0
                ComboBox1.Select(0, 0)
            End If
        Catch ex As Exception
            Error_Handler(ex, "Background Worker Completed")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        RunWorker()
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            Process.Start("http://www.gamespot.com/search.html?type=11&stype=all&tag=search%3Bbutton&qs=" & LinkLabel2.Tag.ToString.Replace(" ", "+") & "&x=36&y=5")
        Catch ex As Exception
            Error_Handler(ex, "GameSpot Link Clicked")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Process.Start("http://www.gamefaqs.com/search/index.html?game=" & LinkLabel1.Tag.ToString.Replace(" ", "+") & "&platform=7")
        Catch ex As Exception
            Error_Handler(ex, "GameFAQs Link Clicked")
        End Try
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        Try
            AutoUpdate = True
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex, "AutoUpdate")
        End Try
    End Sub

    Private Sub Main_Screen_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            If AutoUpdate = True Then
                If My.Computer.FileSystem.FileExists((Application.StartupPath & "\AutoUpdate.exe").Replace("\\", "\")) = True Then
                    Dim startinfo As ProcessStartInfo = New ProcessStartInfo
                    startinfo.FileName = (Application.StartupPath & "\AutoUpdate.exe").Replace("\\", "\")
                    startinfo.Arguments = "force"
                    startinfo.CreateNoWindow = False
                    Process.Start(startinfo)
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex, "Application Closed")
        End Try
    End Sub
End Class
