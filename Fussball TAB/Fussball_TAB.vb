Option Explicit On
Imports System.Xml
Imports System.Net.Sockets
Imports System.IO
Imports System.IO.Ports
Imports Casparobjects

Public Class Fussball
    Dim Client As New TcpClient 'Client
    Dim stream As NetworkStream
    Dim streamw As StreamWriter  'Zum senden
    Dim streamr As StreamReader  'Zum empfangen

    Public watchfolder As FileSystemWatcher

    Public home_active As Integer = 0
    Public away_active As Integer = 0

    Public schiripanel_visible As Boolean


    Public channel(10) As String
    Public textfield(100) As String
    Public numfield(100) As Single
    Public checked(100) As Boolean

    'variable for filename of the initial settings (ip, color, channel)
    Public Shared filename As String = "C:\CG_Sports\Fussball\Fussball_Data.xml"
    Public directory As String = "C:\CG_Sports\Fussball\"
    'Dim directory As String = "C:\CG_Sports\Fussball"

    'variables for caspar ip and channels
    Public ip As String = "localhost"
    Dim Channel1 As String = "1"
    Dim Channel2 As String = "2"

    Public Connected As Integer = 0
    Dim kanal As Integer
    Dim ri As ReturnInfo
    'for Didi Kunz library
    Private _Caspar As CasparCG = New CasparCG()
    '---------------------------------

    'actual game teams
    Dim actual_game_homelong As String
    Dim actual_game_homeshort As String
    Dim actual_game_awaylong As String
    Dim actual_game_awayshort As String

    'variables for color handling of buttons and rgb values of the flash templates
    Dim h1rgb As String = "0xffffff"
    Dim h1r As Int32
    Dim h1g As Int32
    Dim h1b As Int32
    Dim h2rgb As String = "0xffffff"
    Dim h2r As Int32
    Dim h2g As Int32
    Dim h2b As Int32
    Dim a1rgb As String = "0x000000"
    Dim a1r As Int32
    Dim a1g As Int32
    Dim a1b As Int32
    Dim a2rgb As String = "0x000000"
    Dim a2r As Int32
    Dim a2g As Int32
    Dim a2b As Int32

    'clock variables
    Dim MyTime, MyDate, MyStr
    Dim StopWatch1 As New Diagnostics.Stopwatch
    Dim StopWatch2 As New Diagnostics.Stopwatch
    Dim StopWatch3 As New Diagnostics.Stopwatch
    Dim time_is_over As Boolean = False
    Dim zeit As Integer
    Dim uhrlaeuft As Integer
    Dim time As DateTime = DateTime.Now
    Dim format As String = "HH:mm:ss"
    ' This integer variables keeps track of the remaining time.
    Dim elapsed1 As Integer
    Dim timeLeft As Integer
    Dim elapsed2 As Integer
    Dim timeLeft2 As Integer
    'elapsed seconds
    Dim vergangene_sekunden As Integer
    Dim vergangene_sekunden2 As Integer
    ' displayed seconds 
    Dim angezeigte_sekunden As Integer
    Dim angezeigte_sekunden2 As Integer
    ' correct time +-
    Dim correct_time As Integer

    ' Dim MsgBoxResult As IButtonControl

    'variables for the advertising logos
    Dim werbunglogo1 As String
    Dim werbunglogo2 As String
    Dim werbunglogo3 As String
    Dim werbunglogo4 As String

    'variables if onscreen or not
    Dim logo1_on As Boolean = False
    Dim logo2_on As Boolean = False
    Dim logo3_on As Boolean = False
    Dim logo4_on As Boolean = False

    'rasen and Fussball are two statusbar buttons for CG-background on/off on layer 5
    Dim Fussball_Rasen_on As Boolean = False
    Dim Fussball_Image_on As Boolean = False
    Dim Fussball_Video_on As Boolean = False

    Dim filenameland As String

    'holds all home and guest players, 5' players à 9 fields
    Dim hmatrix(50, 9) As String
    Dim amatrix(50, 9) As String

    'variables for sorting parts of hmatrix and amatrix
    Private homearray(0 To 30) As String
    Private homearraynr(0 To 30) As String
    Private awayarray(0 To 30) As String
    Private awayarraynr(0 To 30) As String
    Private homeascendingsort(0 To 30) As Integer
    Private awayascendingsort(0 To 30) As Integer
    Public homestartlineup(30) As String
    Public homesubstitutes(30) As String
    Public awaystartlineup(30) As String
    Public awaysubstitutes(30) As String
    'Dim bh(0 To 30) As String
    'Dim ba(0 To 30) As String

    'flags / logos of clubs
    Dim hflagge As String
    Dim aflagge As String

    ' colors for casper scorebug
    Dim h1color As String
    Dim a1color As String
    Dim h2color As String
    Dim a2color As String

    'array holds all colors from xml file
    Dim colorhome(8) As String
    Dim coloraway(8) As String

    'arrays for home trainers
    Dim ht(6) As String ' trainer
    Dim hc(6) As String ' cotrainer
    Dim hcc(6) As String 'cocotrainer

    'arrays for away trainers
    Dim at(6) As String
    Dim ac(6) As String
    Dim acc(6) As String

    'arrays for titles
    Dim title(50) As String

    'arrays for settings
    Dim settings(50) As String

    'score
    Dim score_home As Integer
    Dim score_away As Integer

    'variables to fill all 9 template-values per player for CG
    Dim CG_nr As String
    Dim CG_name As String
    Dim CG_firstname As String
    Dim CG_pos As String
    Dim CG_Data1 As String
    Dim CG_Data2 As String
    Dim CG_Data3 As String
    Dim CG_Data4 As String
    Dim CG_Data5 As String

    'referees and names
    Dim refname(50) As String
    Dim reffunction(50) As String
    Dim freenamefirstrow(50) As String
    Dim freenamesecondrow(50) As String

    'how many timeouts, period text
    Dim timeout_h As Integer = 0
    Dim timeout_a As Integer = 0
    Dim period_text As String


    Dim yellowcard As String
    Dim redcard As String
    Dim yellowredcard As String
    Dim emptycard As String
    Dim cardon As Boolean = False

    Dim Werbelogo As String

    Dim cg_on As Boolean = False
    Dim scorebug_on As Boolean = False
    Dim lineup1_on As Integer = False
    Dim lineup2_on As Integer = False
    Dim singlename_on As Boolean = False
    Dim freetext As Boolean = False


    Dim result_big_state As Boolean = False
    Dim Titel_state As Boolean = False
    Dim Titel2_state As Boolean = False
    Dim tabellen_state As Boolean = False
    Dim schiri1_state As Boolean = False
    Dim schiri2_state As Boolean = False
    Dim schiri3_state As Boolean = False
    Dim schiri4_state As Boolean = False

    Dim button_state As Boolean = False

    Dim COMPort As New SerialPort
    Dim gpi1_pressed As Boolean = False
    Dim gpi2_pressed As Boolean = False
    Dim gpi3_pressed As Boolean = False
    Dim gpi4_pressed As Boolean = False
    Dim COMPort_NR As String

    Dim button_scoreplushometext As String
    Dim button_scoreplusawaytext As String
    Dim button_scoreminushometext As String
    Dim button_scoreminusawaytext As String

    Dim freetext_active As Boolean = False
    Dim button_name As String
    Dim anzahl_buttons As Integer
    Private burstbox As New List(Of Button)

    Public Config_On As Boolean = 0
    Public filechanged As Boolean = False

    Sub Watch_data()
        'this is the path we want to monitor
        watchfolder = New System.IO.FileSystemWatcher With {.Path = directory}
        watchfolder.NotifyFilter = watchfolder.NotifyFilter Or IO.NotifyFilters.Attributes
        ' add the handler to each event
        AddHandler watchfolder.Changed, AddressOf Logchange
        'Set this property to true to start watching
        watchfolder.EnableRaisingEvents = True
    End Sub

    Private Sub Logchange(ByVal source As Object, ByVal e As System.IO.FileSystemEventArgs)
        If e.ChangeType = IO.WatcherChangeTypes.Changed Then
            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel1.Text &= "File " & e.FullPath & vbCrLf & "has been modified"
            filechanged = True
        End If
    End Sub

    Sub Connect_to_caspar()
        If Connected = 0 Then
            Try
                _Caspar.Connect(ip)
                Status_CGConnected.BackColor = Color.LightGreen
                Status_CGConnected.Text = ip
                Connected = 1
            Catch ex As Exception
                MsgBox("Caspar starten")
                Status_CGConnected.Text = "not connected"
                'Button_Einstellungen.PerformClick()
                'Me.Close()
                Exit Sub
            End Try
        End If
    End Sub

    Private Sub Fussball_TAB_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
        Read_Init()
        'read data from team away xml file
        Read_teamhome()
        'read data from team away xml file
        Read_teamaway()
        ' labels buttons 

        ' sets clock to default time and 0
        Timer1.Enabled = True
        Timer1.Stop()
        uhrlaeuft = 0
        clock.Text = vbNullString

        ' calls sub to set radiobuttons, clock-text and period text for CG
        Check_radiobuttons()

        Reset_clock()
        Format_seconds()

        RadioButton1.Checked = True

        If CheckBox1.Checked = True Then Channel2 = 2
        If CheckBox1.Checked = False Then Channel2 = 1

        'needed for keypress sub 
        Me.KeyPreview = True

        ' open com port for GPI Interface
        Open_COMPort()

        Me.Text = My.Application.Info.Title & " " & My.Application.Info.Version.ToString '+ " - " & My.Application.Info.Description
        Watch_data()
        Connect_to_caspar()

    End Sub

    Sub Read_Init()
        filename = "C:\CG_Sports\Fussball\Fussball_Data.xml"
        directory = "C:\CG_Sports\Fussball\"
        ' checks directory
        If (Not System.IO.Directory.Exists(directory)) Then
            System.IO.Directory.CreateDirectory(directory)
        End If

        'if file not exits, create empty file
        If (Not System.IO.File.Exists(filename)) Then
            'no initial file found, exiting
            MsgBox(" no initial Data found, please use Datacenter first")
            End
        End If

        'fills all variables textfield out of Datafile
        Try
            Dim Contactlist As XDocument = XDocument.Load(filename)

            For Each contact As XElement In Contactlist...<Fussball>
                ip = contact.Element("IP")
                For i = 1 To 10
                    channel(i) = contact.Element("channel" + Trim(i))
                Next
                For i = 1 To 50
                    refname(i) = contact.Element("refname" + Trim(i))
                    reffunction(i) = contact.Element("reffunction" + Trim(i))
                    freenamefirstrow(i) = contact.Element("freenamefirstrow" + Trim(i))
                    freenamesecondrow(i) = contact.Element("freenamesecondrow" + Trim(i))
                    settings(i) = contact.Element("settings" + Trim(i))
                    title(i) = contact.Element("title" + Trim(i))
                Next

                actual_game_homelong = contact.Element("actual_game_homelong")
                actual_game_homeshort = contact.Element("actual_game_homeshort")
                actual_game_awaylong = contact.Element("actual_game_awaylong")
                actual_game_awayshort = contact.Element("actual_game_awayshort")

                For i = 1 To 8
                    colorhome(i) = contact.Element("colorhome" + Trim(i))
                    coloraway(i) = contact.Element("coloraway" + Trim(i))
                Next
                For i = 1 To 50
                    refname(i) = contact.Element("refname" + Trim(i))
                    reffunction(i) = contact.Element("reffunction" + Trim(i))
                Next

                For i = 1 To 50
                    freenamefirstrow(i) = contact.Element("freenamefirstrow" + Trim(i))
                    freenamesecondrow(i) = contact.Element("freenamesecondrow" + Trim(i))
                Next

                For i = 1 To 100
                    textfield(i) = (contact.Element("textfield" + Trim(i)))
                    numfield(i) = contact.Element("numfield" + Trim(i))
                    checked(i) = contact.Element("checked" + Trim(i))
                Next
            Next
        Catch ex As System.IO.IOException
            MessageBox.Show("File nicht vorhanden")
        Catch ex As NullReferenceException
            MessageBox.Show("NullReferenceException: " & ex.Message)
            MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)
        Catch ex As Exception
        End Try

        For i = 1 To 100
            If textfield(i) = "not used" Then textfield(i) = ""
        Next

        Set_Variables()
        Set_labels()
    End Sub

    Sub Set_Variables()
        h1rgb = colorhome(1)
        h1r = colorhome(2)
        h1g = colorhome(3)
        h1b = colorhome(4)
        h2rgb = colorhome(5)
        h2r = colorhome(6)
        h2g = colorhome(7)
        h2b = colorhome(8)
        a1rgb = coloraway(1)
        a1r = coloraway(2)
        a1g = coloraway(3)
        a1b = coloraway(4)
        a2rgb = coloraway(5)
        a2r = coloraway(6)
        a2g = coloraway(7)
        a2b = coloraway(8)

        h1color = h1rgb
        a1color = a1rgb
        h2color = h2rgb
        a2color = a2rgb

        'timeLeft = settings(1)
        timeLeft = 0

        COMPort_NR = Trim(settings(28))
        button_scoreplushometext = Trim(settings(24))
        button_scoreplusawaytext = Trim(settings(25))
        button_scoreminushometext = Trim(settings(26))
        button_scoreminusawaytext = Trim(settings(27))

        'replaces windows \ with / in filenames
        werbunglogo1 = Trim(settings(15).Replace("\", "/"))
        werbunglogo2 = Trim(settings(16).Replace("\", "/"))
        werbunglogo3 = Trim(settings(17).Replace("\", "/"))
        werbunglogo4 = Trim(settings(18).Replace("\", "/"))

        yellowcard = Trim(settings(19).Replace("\", "/"))
        redcard = Trim(settings(20).Replace("\", "/"))
        yellowredcard = Trim(settings(21).Replace("\", "/"))
        emptycard = Trim(settings(22).Replace("\", "/"))

    End Sub

    Sub Set_labels()
        'fills color buttons with the team colors
        colorpanelhome1.BackColor = Color.FromArgb(h1r, h1g, h1b)
        colorpanelhome2.BackColor = Color.FromArgb(h2r, h2g, h2b)
        colorpanelaway1.BackColor = Color.FromArgb(a1r, a1g, a1b)
        colorpanelaway2.BackColor = Color.FromArgb(a2r, a2g, a2b)

        'labels all buttons to needed values during runtime or startup
        'here you can change to your language
        Home.Text = actual_game_homelong
        homek.Text = actual_game_homeshort
        Away.Text = actual_game_awaylong
        awayk.Text = actual_game_awayshort


        'button_timeouth.Text = "timeout " + Chr(13) + LSet(home.Text, 8)
        'button_timeouta.Text = "timeout " + Chr(13) + LSet(away.Text, 8)

        buttonscorebug.Text = "Resultat klein scorebug" + Chr(13) + "F1"
        result_big.Text = "Resultat gross" + Chr(13) + "F2"

        'Tabelle_1.Text = "Tabelle 1" + Chr(13) + "F5"
        'Tabelle_2.Text = "Tabelle 2" + Chr(13) + "F6"
        'Tabelle_3.Text = "Tabelle 3" + Chr(13) + "F7"
        'Tabelle_4.Text = "Tabelle 4" + Chr(13) + "F8"

        clock_start.Text = "START" + Chr(13) + "Numpad 0"
        clock_stop.Text = "STOP" + Chr(13) + "Numpad ."
        minus.Text = "-" + Chr(13) + "Numpad -"
        plus.Text = "+" + Chr(13) + "Numpad +"
        'clear_screen.Text = "CLEAR SCREEN" ' + Chr(13) + "c" + Chr(13) + " shift-c clears Channel2 "

        'extracts picturename out of filename of the advertising logo buttons
        'Dim leftPart As String
        'Dim count As Integer
        'count = CountCharacter(werbunglogo1, "/")
        'leftPart = werbunglogo1.Split("/")(count)
        'Label11.Text = (leftPart)
        'count = CountCharacter(werbunglogo2, "/")
        'leftPart = werbunglogo2.Split("/")(count)
        'Label12.Text = (leftPart)
        'count = CountCharacter(werbunglogo3, "/")
        'leftPart = werbunglogo3.Split("/")(count)
        'Label13.Text = (leftPart)
        'count = CountCharacter(werbunglogo4, "/")
        'leftPart = werbunglogo4.Split("/")(count)
        'Label14.Text = (leftPart)

        Status_Version.Text = "Version " & My.Application.Info.Version.ToString
        Status_Copyright.Text = My.Application.Info.Copyright

        'labels buttons
        'button_foul1.Text = settings(8)
        'button_foul2.Text = settings(9)
        'button_foul3.Text = settings(10)
        'button_foul4.Text = settings(11)
        'button_foul5.Text = settings(12)
        'button_foul6.Text = settings(13)
        'button_foul7.Text = settings(14)

        RadioButton1.Text = settings(4)
        RadioButton2.Text = settings(5)
        RadioButton3.Text = settings(6)
        RadioButton4.Text = settings(7)

        torheimlabel.Text = score_home
        torgastlabel.Text = score_away
        'secondline_result.Text = Trim(Str(score_home)) + " : " + Trim(Str(score_away))

        'label referee buttons
        button_referee_1.Text = refname(1)
        button_referee_2.Text = refname(2)
        button_referee_3.Text = refname(3)
        button_referee_4.Text = refname(4)

        'torheimbutton.Text = "← Score + " + Chr(13) + home.Text
        'torgastbutton.Text = "+ Score → " + Chr(13) + away.Text
        'minustorheimbutton.Text = "- Score " + Chr(13) + home.Text
        'minustorgastbutton.Text = "- Score " + Chr(13) + away.Text

        'torheimbutton.Text = button_scoreplushometext + Chr(13) + "←"
        'torgastbutton.Text = button_scoreplusawaytext + Chr(13) + "→"
        'minustorheimbutton.Text = button_scoreminushometext
        'minustorgastbutton.Text = button_scoreminusawaytext
        ToolStripStatusLabel1.Text = "Data up to date"
    End Sub

    Sub Read_teamhome()
        ' reads all players home
        filenameland = "C:\CG_Sports\Fussball\" + Trim(actual_game_homelong) + ".xml"
        Dim Contactlist As XDocument = XDocument.Load(filenameland)
        For Each contact As XElement In Contactlist...<Fussballteam>
            hflagge = contact.Element("Flag")
            For i = 1 To 30
                hmatrix(i, 0) = contact.Element("Player" + Trim(i) + "1")
                If hmatrix(i, 0) = "not used" Then hmatrix(i, 0) = "0"

                hmatrix(i, 1) = contact.Element("Player" + Trim(i) + "2")
                hmatrix(i, 2) = contact.Element("Player" + Trim(i) + "3")
                hmatrix(i, 3) = contact.Element("Player" + Trim(i) + "4")
                hmatrix(i, 4) = contact.Element("Player" + Trim(i) + "5")
                hmatrix(i, 5) = contact.Element("Player" + Trim(i) + "6")
                hmatrix(i, 6) = contact.Element("Player" + Trim(i) + "7")
                hmatrix(i, 7) = contact.Element("Player" + Trim(i) + "8")
                hmatrix(i, 8) = contact.Element("Player" + Trim(i) + "9")
                homestartlineup(i) = contact.Element("Startlineup" + Trim(i))
                homesubstitutes(i) = contact.Element("Substitutes" + Trim(i))
            Next
            For i = 1 To 6
                ht(i) = contact.Element("coach" + Trim(i))
                hc(i) = contact.Element("first_assistant" + Trim(i))
                hcc(i) = contact.Element("second_assistant" + Trim(i))
            Next
        Next

        'Numerische Array sortierung für Buttons
        'füllt arraycopy() mit NR - Name - Vorname  

        For i = 1 To 30
            'namen ohne nummern bekommen die nummer 9999
            If Trim(hmatrix(i, 0)) = "" Then hmatrix(i, 0) = 9999
            'füllt ih mit allen vorhandenen Nummern zum sortieren
            If (hmatrix(i, 0)) > 0 Then homeascendingsort(i) = Int(hmatrix(i, 0))
            'If (hmatrix(i, 0)) < 0 Or (hmatrix(i, 0)) Mod 9999 Then homeascendingsort(i) = Int(hmatrix(i, 0))
            'namen ohne nummern bekommen die nummer 9999
            If hmatrix(i, 0) = 9999 Then hmatrix(i, 0) = ""
            homearraynr(i) = hmatrix(i, 0) + "  " + hmatrix(i, 1)
        Next i

        'sortiert Array
        Array.Sort(homeascendingsort, homearraynr)
        'For i = 1 To 30
        '    MsgBox(Trim(homearraynr(i)))
        'Next
        'Alphabetische Combobox hComboBox
        'füllt array mit Name -NR - Vorname
        For i = 1 To 30
            If Len(hmatrix(i, 0)) < 2 Then hmatrix(i, 0) = " " + hmatrix(i, 0)
            homearray(i) = hmatrix(i, 1) + "  " + hmatrix(i, 2) + "  " + hmatrix(i, 0)
        Next i
        'sortiert Array 
        Array.Sort(homearray)

        'löscht Combobox 2
        hComboBox.Items.Clear()

        '_________________________________________________________________________________________________
        ' füllt Combobox hComboBox , löscht dabei Nullwerte
        For i = 1 To 30
            If Trim(homearray(i)) > "" Then hComboBox.Items.Add(homearray(i))
            'If Trim(homeascendingsort(i)) IsNot "9999" Then bh(i) = (homeascendingsort(i))
        Next
        '_________________________________________________________________________________________________
        'defines flag actual_game_homelong on the casper server computer
        hflagge = "C:/CG_Sports/flags/" + Trim(hflagge) + ".png"
        '_________________________________________________________________________________________________
        'defines flag actual_game_homelong on the computer which runs this program,
        If My.Computer.FileSystem.FileExists(hflagge.Replace("/", "\")) Then PictureHome.Load(hflagge)
    End Sub

    Sub Read_teamaway()
        ' reads all players away
        filenameland = "C:\CG_Sports\Fussball\" + Trim(actual_game_awaylong) + ".xml"
        Dim Contactlist As XDocument = XDocument.Load(filenameland)
        For Each contact As XElement In Contactlist...<Fussballteam>
            aflagge = contact.Element("Flag")

            For i = 1 To 30
                amatrix(i, 0) = contact.Element("Player" + Trim(i) + "1")
                If amatrix(i, 0) = "not used" Then amatrix(i, 0) = "0"

                amatrix(i, 1) = contact.Element("Player" + Trim(i) + "2")
                amatrix(i, 2) = contact.Element("Player" + Trim(i) + "3")
                amatrix(i, 3) = contact.Element("Player" + Trim(i) + "4")
                amatrix(i, 4) = contact.Element("Player" + Trim(i) + "5")
                amatrix(i, 5) = contact.Element("Player" + Trim(i) + "6")
                amatrix(i, 6) = contact.Element("Player" + Trim(i) + "7")
                amatrix(i, 7) = contact.Element("Player" + Trim(i) + "8")
                amatrix(i, 8) = contact.Element("Player" + Trim(i) + "9")
                awaystartlineup(i) = contact.Element("Startlineup" + Trim(i))
                awaysubstitutes(i) = contact.Element("Substitutes" + Trim(i))

            Next
            For i = 1 To 6
                at(i) = contact.Element("coach" + Trim(i))
                ac(i) = contact.Element("first_assistant" + Trim(i))
                acc(i) = contact.Element("second_assistant" + Trim(i))
            Next
        Next
        'Numerische Array sortierung für Buttons
        'füllt arraycopy() mit NR - Name - Vorname  

        For i = 1 To 30
            'namen ohne nummern bekommen die nummer 9999
            If Trim(amatrix(i, 0)) = vbNullString Then amatrix(i, 0) = 9999
            'füllt ih mit allen vorhandenen Nummern zum sortieren
            If (amatrix(i, 0)) > 0 Then awayascendingsort(i) = Int(amatrix(i, 0))
            'namen ohne nummern bekommen die nummer 9999
            If amatrix(i, 0) = 9999 Then amatrix(i, 0) = vbNullString
            awayarraynr(i) = amatrix(i, 0) + "  " + amatrix(i, 1)
        Next i

        'sortiert Array
        Array.Sort(awayascendingsort, awayarraynr)

        'Alphabetische Combobox hComboBox
        'füllt array mit Name -NR - Vorname
        For i = 1 To 30
            If Len(amatrix(i, 0)) < 2 Then amatrix(i, 0) = vbNullString + amatrix(i, 0)
            awayarray(i) = amatrix(i, 1) + "  " + amatrix(i, 2) + "  " + amatrix(i, 0)
        Next i

        'sortiert Array 
        Array.Sort(awayarray)
        'löscht Combobox 2
        aComboBox.Items.Clear()

        '_________________________________________________________________________________________________
        ' füllt Combobox hComboBox , löscht dabei Nullwerte
        For i = 1 To 30
            If Trim(awayarray(i)) > vbNullString Then aComboBox.Items.Add(awayarray(i))
            'If Trim(awayascendingsort(i)) IsNot "9999" Then bh(i) = (awayascendingsort(i))
        Next
        '_________________________________________________________________________________________________
        'defines flag actual_game_awaylong on the casper server computer
        aflagge = "C:/CG_Sports/flags/" + Trim(aflagge) + ".png"
        '_________________________________________________________________________________________________
        'defines flag actual_game_awaylong on the computer which runs this program,
        If My.Computer.FileSystem.FileExists(aflagge.Replace("/", "\")) Then PictureAway.Load(aflagge)
    End Sub

    Sub Update_data(filename_update, node, updated_data)
        'update a specific XML node direct
        'sub, receives filename, number of dataset, value to change, saves a specific XML dataset back
        If Trim(updated_data) = vbNullString Then updated_data = "not used"
        Dim MyXML As New XmlDocument()
        Dim MyXMLNode As XmlNode
        MyXML.Load(filename_update)
        MyXMLNode = MyXML.SelectSingleNode(node)
        If MyXMLNode IsNot Nothing Then
            MyXMLNode.ChildNodes(0).InnerText = updated_data
            MyXML.Save(filename_update)
        Else
            MessageBox.Show("update Fehler")
        End If ' Save the Xml.
    End Sub

    Private Sub Clock_stop_Click(sender As Object, e As EventArgs) Handles clock_stop.Click
        Stoptime()
        If time_is_over = False Then
            clock_stop.BackColor = Color.FromArgb(224, 224, 224)
            clock_stop.Text = "STOP" + Chr(13) + "Numpad 0"
            If RadioButton3.Checked = True Then RadioButton4.Checked = True
            If RadioButton2.Checked = True Then RadioButton3.Checked = True
            If RadioButton1.Checked = True Then RadioButton2.Checked = True

            If RadioButton1.Checked = True Then timeLeft = 0
            If RadioButton2.Checked = True Then timeLeft = 45 * 60
            If RadioButton3.Checked = True Then timeLeft = 90 * 60
            If RadioButton4.Checked = True Then timeLeft = 105 * 60
            Format_seconds()
        End If
    End Sub

    Sub Stoptime()
        Me.StopWatch1.Stop()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim elapsed1 As TimeSpan = Me.StopWatch1.Elapsed
        vergangene_sekunden = String.Format("{2:00}", Math.Floor(elapsed1.TotalHours), elapsed1.Minutes, elapsed1.TotalSeconds)
        vergangene_sekunden = vergangene_sekunden - correct_time
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        angezeigte_sekunden = timeLeft + vergangene_sekunden
        'hier erfolgt die eigentliche anzeige
        clock.Text = Format_timestring(angezeigte_sekunden)
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        If filechanged = True Then
            Read_Init()
            ''read data from team away xml file
            Read_teamhome()
            ''read data from team away xml file
            Read_teamaway()
            ToolStripStatusLabel1.Text = "Data up to date"
        End If
    End Sub

    Function Format_timestring(ByVal intSeconds)
        ' formats the timestring to the needed format
        ' see all isolated times (m,s,) in the bottom-statusbar
        Dim minutes, seconds

        ' calculates the whole number of minutes in the remaining number of seconds
        minutes = intSeconds \ 60
        ' minutes = intSeconds
        Status_Minute.Text = minutes 'statusline minutes
        If minutes < 10 Then minutes = "0" & minutes

        ' calculates the remaining number of seconds after taking the number of minutes
        seconds = intSeconds Mod 60
        If seconds < 10 Then seconds = "0" & seconds
        Status_Sekunde.Text = seconds 'statusline seconds

        Status_IntSekunde.Text = intSeconds 'statusline total seconds from beginning

        ' returns as a string
        Format_timestring = minutes & ":" & seconds
    End Function

    Sub Format_seconds()
        Reset_clock()
        'formats seconds to needed time format
        clock.Text = Format_timestring(timeLeft)

        'sets the correction factor to 0
        correct_time = 0

        'here only for the bottom statusbar display of minutes
        If Str(timeLeft / 60) = 1 Then
            Status_Minute.Text = Str(timeLeft / 60) + " Minute"
        End If
        If Str(timeLeft / 60) > 1 Then
            Status_Minute.Text = Str(timeLeft / 60) + " Minuten"
        End If
    End Sub

    Private Sub Clock_doubleClick(sender As Object, e As EventArgs) Handles clock.DoubleClick
        'calls sub to set radiobuttons, clock-text and period text for CG
        Check_radiobuttons()
        timeLeft = 0
        RadioButton1.Checked = True
        Dim tmpl As Template = New Template
        _Caspar.CG_Update(Channel2, 21, tmpl)
        Reset_clock()
        Format_seconds()
        timeLeft = 0
        RadioButton1.Checked = True
    End Sub

    Private Sub Clock_TextChanged(sender As Object, e As EventArgs) Handles clock.TextChanged
        If scorebug_on = True Then
            Dim tmpl As Template = New Template
            tmpl.Fields.Add(New TemplateField("_time", clock.Text))
            _Caspar.CG_Update(Channel2, 21, tmpl)
        End If
    End Sub

    Sub Check_radiobuttons()
        If time_is_over = False Then
            'sub to set radiobuttons, clock-text and period text for CG
            If RadioButton1.Checked = True Then timeLeft = 0 : clock.Text = 0 : period_text = settings(4)
            If RadioButton2.Checked = True Then timeLeft = 45 * 60 : clock.Text = 45 * 60 : period_text = settings(5)
            If RadioButton3.Checked = True Then timeLeft = 90 * 60 : clock.Text = 90 * 60 : period_text = settings(6)
            If RadioButton4.Checked = True Then timeLeft = 105 * 60 : clock.Text = 105 * 60 : period_text = settings(7)
        End If
    End Sub
    Private Sub PictureHome_Click(sender As Object, e As EventArgs) Handles PictureHome.Click
        If home_active = 0 Then
            Panel_away.Visible = False
            Show_Torplusminus()
            home_active = 1
            Button_TorPlus.BackColor = Color.FromArgb(h1r, h1g, h1b)
        Else
            Reset_Main()
            Exit Sub
        End If

        'beschriftet Buttons Spieler Heim
        Dim ii = 1
        For i = 1 To 30
            If homestartlineup(i) = Trim("True") Then
                live.Controls.Item("Button_Spieler_" & ii).Visible = True
                live.Controls.Item("Button_Spieler_" & ii).Text = (hmatrix(i, 0)) + " " + (hmatrix(i, 1))
                'hmatrix(i, 0)
                ii = ii + 1
            End If
        Next i
        Show_Player()

        'beschriftet Buttons Ersatz-Spieler Heim
        ii = 1
        For i = 1 To 30
            If homesubstitutes(i) = Trim("True") Then
                live.Controls.Item("Button_ErsatzSpieler_" & ii).Visible = True
                live.Controls.Item("Button_ErsatzSpieler_" & ii).Text = (hmatrix(i, 0)) + " " + (hmatrix(i, 1))
                ii = ii + 1
            End If
        Next i

        ' Label trainerbuttons
        Button_Trainer.Text = ht(2) + " " + ht(1) + "  " + ht(3)
        Button_CoTrainer1.Text = hc(2) + " " + hc(1) + "  " + hc(3)
        Button_CoTrainer2.Text = hcc(2) + " " + hcc(1) + "  " + hcc(3)
        Show_Trainer()
        Button_TorPlus.Text = "Tor  + " + Chr(13) + actual_game_homelong
    End Sub
    Private Sub PictureAway_Click_1(sender As Object, e As EventArgs) Handles PictureAway.Click
        If away_active = 0 Then
            Panel_Home.Visible = False
            Show_Torplusminus()
            away_active = 1
            Button_TorPlus.BackColor = Color.FromArgb(a1r, a1g, a1b)
        Else
            Reset_Main()
            Exit Sub
        End If

        'beschriftet Buttons Spieler Gast
        Dim ii = 1
        For i = 1 To 30
            If awaystartlineup(i) = Trim("True") Then
                live.Controls.Item("Button_Spieler_" & ii).Visible = True
                live.Controls.Item("Button_Spieler_" & ii).Text = (amatrix(i, 0)) + " " + (amatrix(i, 1))
                ii = ii + 1
            End If
        Next i

        Show_Player()

        'beschriftet Etsatz-Buttons Spieler Gast
        ii = 1
        For i = 1 To 30
            If awaysubstitutes(i) = Trim("True") Then
                live.Controls.Item("Button_ErsatzSpieler_" & ii).Visible = True
                live.Controls.Item("Button_ErsatzSpieler_" & ii).Text = (amatrix(i, 0)) + " " + (amatrix(i, 1))
                ii = ii + 1
            End If
        Next i

        ' Label trainerbuttons
        Button_Trainer.Text = at(2) + " " + at(1) + "  " + at(3)
        Button_CoTrainer1.Text = ac(2) + " " + ac(1) + "  " + ac(3)
        Button_CoTrainer2.Text = acc(2) + " " + acc(1) + "  " + acc(3)
        Show_Trainer()
        Button_TorPlus.Text = "Tor  + " + Chr(13) + actual_game_awaylong
    End Sub

    Sub Show_Main()
        For i = 1 To 11
            live.Controls.Item("Button_Spieler_" & i).Visible = True
        Next

        If home_active = 1 Then
            hComboBox.Location = New Point(180, 826)
            hComboBox.Visible = True
            Panel_away.Visible = False
        End If
        If away_active = 1 Then
            aComboBox.Location = New Point(180, 826)
            aComboBox.Visible = True
            Panel_Home.Visible = False
        End If
    End Sub

    Sub Reset_Main()
        Hide_Player()
        Hide_Subst()
        Hide_Trainer()
        Hide_Cards()
        Hide_Torplusminus()
        Hide_Player_Second
        home_active = 0
        away_active = 0
        Panel_Home.Visible = True
        Panel_away.Visible = True

        Label_actual_name.Visible = False
        Label_actual_name.Text = ""
        Label_secondrow.Visible = False
        Label_secondrow.Text = ""
        hComboBox.Text = "home"
        aComboBox.Text = "away"
        'Clear_Player_Buttons_Color()
    End Sub

    Sub Hide_all_Panels()
        'Panel_Werbung.Visible = False
        Label_actual_name.Visible = False
        'Panel_Namen.Visible = False
        aComboBox.Visible = False
        hComboBox.Visible = False
    End Sub

    Sub Show_Player()
        For i = 1 To 11
            live.Controls.Item("Button_Spieler_" & i).Visible = True
        Next
    End Sub

    Sub Hide_Player()
        For i = 1 To 11
            live.Controls.Item("Button_Spieler_" & i).Visible = False
        Next
    End Sub

    Sub Hide_Player_Second()
        For i = 1 To 7
            live.Controls.Item("Button_second" & i).Visible = False
        Next
    End Sub


    Sub Show_Subst()
        For i = 1 To 22
            live.Controls.Item("Button_ErsatzSpieler_" & i).Visible = True
        Next
    End Sub

    Sub Hide_Subst()
        For i = 1 To 22
            live.Controls.Item("Button_ErsatzSpieler_" & i).Visible = False
            live.Controls.Item("Button_ErsatzSpieler_" & i).Text = ""
        Next
    End Sub

    Sub Show_Trainer()
        live.Controls.Item("Button_Trainer").Visible = True
        live.Controls.Item("Button_CoTrainer1").Visible = True
        live.Controls.Item("Button_CoTrainer2").Visible = True
    End Sub

    Sub Hide_Trainer()
        live.Controls.Item("Button_Trainer").Visible = False
        live.Controls.Item("Button_CoTrainer1").Visible = False
        live.Controls.Item("Button_CoTrainer2").Visible = False
        live.Controls.Item("Button_Trainer").Text = ""
        live.Controls.Item("Button_CoTrainer1").Text = ""
        live.Controls.Item("Button_CoTrainer2").Text = ""
    End Sub

    Sub Show_cards()
        yellow.Visible = True
        red.Visible = True
        yellowred.Visible = True
    End Sub

    Sub Hide_Cards()
        yellow.Visible = False
        red.Visible = False
        yellowred.Visible = False
    End Sub

    Sub Show_Torplusminus()
        Button_TorPlus.Visible = True
        Button_TorMinus.Visible = True
    End Sub

    Sub Hide_Torplusminus()
        Button_TorPlus.Visible = False
        Button_TorMinus.Visible = False
    End Sub

    Sub Clear_Player_Buttons_Color()
        Button_Spieler_1.BackColor = Color.Gainsboro
        Button_Spieler_2.BackColor = Color.Gainsboro
        Button_Spieler_3.BackColor = Color.Gainsboro
        Button_Spieler_4.BackColor = Color.Gainsboro
        Button_Spieler_5.BackColor = Color.Gainsboro
        Button_Spieler_6.BackColor = Color.Gainsboro
        Button_Spieler_7.BackColor = Color.Gainsboro
        Button_Spieler_8.BackColor = Color.Gainsboro
        Button_Spieler_9.BackColor = Color.Gainsboro
        Button_Spieler_10.BackColor = Color.Gainsboro
        Button_Spieler_11.BackColor = Color.Gainsboro
        Button_Ersatzspieler_1.BackColor = Color.Gainsboro
        Button_Ersatzspieler_2.BackColor = Color.Gainsboro
        Button_Ersatzspieler_3.BackColor = Color.Gainsboro
        Button_Ersatzspieler_4.BackColor = Color.Gainsboro
        Button_Ersatzspieler_5.BackColor = Color.Gainsboro
        Button_Ersatzspieler_6.BackColor = Color.Gainsboro
        Button_Ersatzspieler_7.BackColor = Color.Gainsboro
        Button_Ersatzspieler_8.BackColor = Color.Gainsboro
        Button_Ersatzspieler_9.BackColor = Color.Gainsboro
        Button_Ersatzspieler_10.BackColor = Color.Gainsboro
        Button_Ersatzspieler_11.BackColor = Color.Gainsboro
        Button_Ersatzspieler_12.BackColor = Color.Gainsboro
        Button_Ersatzspieler_13.BackColor = Color.Gainsboro
        Button_Ersatzspieler_14.BackColor = Color.Gainsboro
        Button_Ersatzspieler_15.BackColor = Color.Gainsboro
        Button_Ersatzspieler_16.BackColor = Color.Gainsboro
        Button_Ersatzspieler_17.BackColor = Color.Gainsboro
        Button_Ersatzspieler_18.BackColor = Color.Gainsboro
        Button_Ersatzspieler_19.BackColor = Color.Gainsboro
        Button_Ersatzspieler_20.BackColor = Color.Gainsboro
        Button_Ersatzspieler_21.BackColor = Color.Gainsboro
        Button_Ersatzspieler_22.BackColor = Color.Gainsboro
    End Sub

    Private Sub Update_Gamedata()
        'result big
        Dim tmpl As Template = New Template

        tmpl.Fields.Add(New TemplateField("_home", Home.Text))
        tmpl.Fields.Add(New TemplateField("_away", Away.Text))
        tmpl.Fields.Add(New TemplateField("_zeit", Trim(Str(score_home)) + " : " + Trim(Str(score_away))))
        tmpl.Fields.Add(New TemplateField("_hpoint", Trim(Str(score_home))))
        tmpl.Fields.Add(New TemplateField("_apoint", Trim(Str(score_away))))

        'all in here for flash template
        tmpl.Fields.Add(New TemplateField("hflagge", hflagge))
        tmpl.Fields.Add(New TemplateField("aflagge", aflagge))

        ' if HTML template from E2C librrary then the caspar needs it like this: aflagge = "img=c: /CG_Sports/flags/ch.png"
        tmpl.Fields.Add(New TemplateField("hflagge_html", "img=C:/CG_Sports/flags/" + Trim(hflagge) + ".png"))
        tmpl.Fields.Add(New TemplateField("aflagge_html", "img=C:/CG_Sports/flags/" + Trim(aflagge) + ".png"))

        'colors for teamidentifyer
        tmpl.Fields.Add(New TemplateField("_homecolor1", h1color))
        tmpl.Fields.Add(New TemplateField("_awaycolor1", a1color))
        tmpl.Fields.Add(New TemplateField("_homecolor2", h2color))
        tmpl.Fields.Add(New TemplateField("_awaycolor2", a2color))

        If RadioButton1.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(4)))
        If RadioButton2.Checked = True And clock_stop.Text = "RESET" Then
            tmpl.Fields.Add(New TemplateField("_period", settings(4)))
        Else
            If RadioButton2.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(5)))
        End If

        If RadioButton3.Checked = True And clock_stop.Text = "RESET" Then
            tmpl.Fields.Add(New TemplateField("_period", settings(5)))
        Else
            If RadioButton3.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(6)))
        End If

        If RadioButton4.Checked = True And clock_stop.Text = "RESET" Then
            tmpl.Fields.Add(New TemplateField("_period", settings(6)))
        Else
            If RadioButton4.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(7)))
        End If

        'Return_CG(_Caspar.CG_Add(Channel1, 20, "handball/resgross", tmpl, False))
        _Caspar.CG_Add(Channel1, 20, "handball/resgross", tmpl, False)
        '_______________________________________________________________________________________________________________________________________________________________________________________
        'all in here for HTML template
        'Next 2 lines For Adobe Edge plus Sippos E2C Library Files template (EXAMPLE: Template field hflagge, "img=c:/folder/subfolder/flag.png"            
        'tmpl.Fields.Add(New TemplateField("hflagge", "img=" + hflagge))
        'tmpl.Fields.Add(New TemplateField("aflagge", "img=" + aflagge))

        'example If you want to use a mask image for your logos with  (EXAMPLE: Template field hflagge, "img=c:/folder/subfolder/flag.png;mask=c:/folder/subfolder/mask.pngvbNullString        
        'tmpl.Fields.Add(New TemplateField("hflagge", "img=" + hflagge + ";mask=c:/CG_Sports/flags/mask2.png"))
        'tmpl.Fields.Add(New TemplateField("aflagge", "img=" + aflagge + ";mask=c/CG_Sports/flags/mask2.png"))

        'desperately :-) trying to change background of a color box in a html template
        ' tmpl.Fields.Add(New TemplateField("_homecolor1", "fill:('#ffffff');"))
        '_Caspar.Execute("call 1-20 invoke vbNullString        _homecolor1,fill:('#ffffff');vbNullString        ")
    End Sub

    Private Sub Result_big_Click(sender As Object, e As EventArgs) Handles result_big.Click
        If result_big_state = True Then
            'Fadeout()
            _Caspar.CG_Stop(Channel1, 20)
            result_big_state = False
            result_big.BackColor = Color.Gainsboro
        Else
            'Clear_all_on()
            ' One_after_the_other()
            result_big_state = True
            result_big.BackColor = Color.LightSalmon
            'result big
            Update_Gamedata()
            'Return_CG(_Caspar.CG_Add(Channel1, 20, "handball/resgross", tmpl, True))
            _Caspar.CG_Play(Channel1, 20)

            'this would be the HTML template
            'Return_CG(_Caspar.CG_Add(channel1, 20, "Fussball/resgross3", tmpl, True))
            cg_on = True
            Label_actual_name.Visible = True
            Label_actual_name.Text = Home.Text + " " + Trim(Str(score_home)) + " : " + Trim(Str(score_away)) + " " + Away.Text
        End If
    End Sub

    Function Reset_clock()
        'resets all
        Me.StopWatch1.Reset()
        Me.StopWatch2.Reset()
        correct_time = 0
        vergangene_sekunden = 0
        angezeigte_sekunden = 0

        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        RadioButton3.Enabled = True
        RadioButton4.Enabled = True
        time_is_over = False
        Return (0)
    End Function

    Sub Open_COMPort()
        ' Close the serial port before exiting
        If (COMPort.IsOpen) Then
            COMPort.Close()
        End If

        'gpi com port open
        COMPort.PortName = "COM" + COMPort_NR
        On Error Resume Next
        COMPort.Open()

        If Err.Number > 0 Then
            GPI1.BackColor = Color.Red
            GPI2.Text = vbNullString
            GPI1.Text = "Could not open COM Port " + COMPort_NR + " for GPI Interface"
            On Error GoTo 0
        End If
    End Sub

    Private Sub Button_TorPlus_Click(sender As Object, e As EventArgs) Handles Button_TorPlus.Click
        If home_active = 1 Then
            score_home = score_home + 1
        End If
        If away_active = 1 Then
            score_away = score_away + 1
        End If
        Set_labels()
        Update_Gamedata()
    End Sub

    Private Sub Button_TorMinus_Click(sender As Object, e As EventArgs) Handles Button_TorMinus.Click
        If home_active = 1 Then
            If score_home > 0 Then score_home = score_home - 1
        End If
        If away_active = 1 Then
            If score_away > 0 Then score_away = score_away - 1
        End If
        Set_labels()
        Update_Gamedata()
    End Sub


    Private Sub Button_Spieler_1_Click(sender As Object, e As EventArgs) Handles Button_Spieler_1.Click, Button_Spieler_2.Click, Button_Spieler_3.Click, Button_Spieler_4.Click, Button_Spieler_5.Click, Button_Spieler_6.Click, Button_Spieler_7.Click, Button_Spieler_8.Click, Button_Spieler_9.Click, Button_Spieler_10.Click, Button_Spieler_11.Click,
        Button_Ersatzspieler_1.Click, Button_Ersatzspieler_2.Click, Button_Ersatzspieler_3.Click, Button_Ersatzspieler_4.Click, Button_Ersatzspieler_5.Click, Button_Ersatzspieler_6.Click, Button_Ersatzspieler_7.Click, Button_Ersatzspieler_8.Click, Button_Ersatzspieler_9.Click, Button_Ersatzspieler_10.Click, Button_Ersatzspieler_11.Click,
        Button_Ersatzspieler_12.Click, Button_Ersatzspieler_13.Click, Button_Ersatzspieler_14.Click, Button_Ersatzspieler_15.Click, Button_Ersatzspieler_16.Click, Button_Ersatzspieler_17.Click, Button_Ersatzspieler_18.Click, Button_Ersatzspieler_19.Click, Button_Ersatzspieler_20.Click, Button_Ersatzspieler_21.Click, Button_Ersatzspieler_22.Click
        'collects all buttons for home players, keep the 3 rows up here together (button 1 - button 30)
        Dim b As Button = DirectCast(sender, Button)

        If singlename_on = False Then
            Label_secondrow.Text = ""
            Clear_Player_Buttons_Color()

            b.BackColor = Color.Salmon
            If home_active = 1 Then
                Search_home_button(b.Text)
                hComboBox.Text = ""
            End If
            If away_active = 1 Then
                Search_away_button(b.Text)
                aComboBox.Text = ""
            End If
            Show_cards()
        Else
            Clear_Player_Buttons_Color()
            Fadeout()
            Hide_Player_Second()
        End If

    End Sub

    Sub Search_home_button(search_string_home)
        'search name from buttons home
        Dim searchstring As String
        searchstring = search_string_home
        For i = 1 To 30
            If Trim(hmatrix(i, 1)) > vbNullString Then
                If searchstring.Contains(hmatrix(i, 0) + " " + hmatrix(i, 1)) Then
                    Clear_all_on()
                    'all data per player ready to send to caspar
                    CG_nr = hmatrix(i, 0)
                    CG_name = hmatrix(i, 1)
                    CG_firstname = hmatrix(i, 2)
                    CG_pos = hmatrix(i, 3)
                    CG_Data1 = hmatrix(i, 4)
                    CG_Data2 = hmatrix(i, 5)
                    CG_Data3 = hmatrix(i, 6)
                    CG_Data4 = hmatrix(i, 7)
                    CG_Data5 = hmatrix(i, 8)
                    Dim tmpl As Template = New Template
                    tmpl.Fields.Add(New TemplateField("_nr", CG_nr))
                    tmpl.Fields.Add(New TemplateField("_name", Trim(CG_firstname) + " " + Trim(CG_name)))
                    tmpl.Fields.Add(New TemplateField("logo1", hflagge))

                    If CheckBox2.Checked = True Then
                        tmpl.Fields.Add(New TemplateField("_zeile2", Trim(Str(score_home)) + " :  " + Trim(Str(score_away))))
                    Else
                        tmpl.Fields.Add(New TemplateField("_zeile2", vbNullString))
                    End If

                    '_________________________________________________________________________________________________
                    'NOT SOLVED, checks if a photo of the player exists, folder management of the players Is NOT SOLVED
                    'Dim pic = "C:/CG_Sports/pics/" + hr(6) + "/" + Trim(CG_firstname) + "-" + Trim(CG_name) + ".png"

                    'If My.Computer.FileSystem.FileExists(pic.Replace("/", "\")) Then
                    '    tmpl.Fields.Add(New TemplateField("pic", pic))
                    'Else
                    '    tmpl.Fields.Add(New TemplateField("pic", vbNullString))
                    'End If
                    '_________________________________________________________________________________________________

                    'calls template
                    '                    mixer_squeeze_ready()
                    '                    Return_CG(_Caspar.CG_Add(Channel1, 20, "Fussball/einzelname", tmpl, True))
                    _Caspar.CG_Add(Channel1, 20, "Fussball/einzelname", tmpl, True)

                    '                    mixer_squeeze_in()


                    ' somethin is on screen
                    cg_on = True
                    ' and its a singlename (to allow second line functions)
                    singlename_on = True
                End If
                Label_actual_name.Text = CG_nr + " " + CG_firstname + " " + CG_name : Label_actual_name.Visible = True
            End If

        Next
        ' shows or disables secondrow buttons
        Button_second1.Text = Trim(Str(score_home)) + " : " + Trim(Str(score_away)) : Button_second1.Visible = True
        If Trim(CG_pos) = "" Then Button_second2.Visible = False Else Button_second2.Text = CG_pos : Button_second2.Visible = True
        If Trim(CG_Data1) = "" Then Button_second3.Visible = False Else Button_second3.Text = CG_Data1 : Button_second3.Visible = True
        If Trim(CG_Data2) = "" Then Button_second4.Visible = False Else Button_second4.Text = CG_Data2 : Button_second4.Visible = True
        If Trim(CG_Data3) = "" Then Button_second5.Visible = False Else Button_second5.Text = CG_Data3 : Button_second5.Visible = True
        If Trim(CG_Data4) = "" Then Button_second6.Visible = False Else Button_second6.Text = CG_Data4 : Button_second6.Visible = True
        If Trim(CG_Data5) = "" Then Button_second7.Visible = False Else Button_second7.Text = CG_Data5 : Button_second7.Visible = True
    End Sub

    Sub Search_away_button(search_string_away)
        'search name from buttons away
        Dim searchstring As String
        searchstring = search_string_away

        For i = 1 To 30
            If Trim(amatrix(i, 1)) > vbNullString Then
                If searchstring.Contains(amatrix(i, 0) + " " + amatrix(i, 1)) Then
                    Clear_all_on()
                    CG_nr = amatrix(i, 0)
                    CG_name = amatrix(i, 1)
                    CG_firstname = amatrix(i, 2)
                    CG_pos = amatrix(i, 3)
                    CG_Data1 = amatrix(i, 4)
                    CG_Data2 = amatrix(i, 5)
                    CG_Data3 = amatrix(i, 6)
                    CG_Data4 = amatrix(i, 7)
                    CG_Data5 = amatrix(i, 8)
                    Dim tmpl As Template = New Template
                    tmpl.Fields.Add(New TemplateField("_nr", CG_nr))
                    tmpl.Fields.Add(New TemplateField("_name", Trim(CG_firstname) + " " + Trim(CG_name)))
                    tmpl.Fields.Add(New TemplateField("logo1", aflagge))

                    If CheckBox2.Checked = True Then
                        tmpl.Fields.Add(New TemplateField("_zeile2", Trim(Str(score_home)) + " :  " + Trim(Str(score_away))))
                    Else
                        tmpl.Fields.Add(New TemplateField("_zeile2", vbNullString))
                    End If

                    '_________________________________________________________________________________________________
                    'NOT SOLVED, checks if a photo of the player exists, folder management of the players Is NOT SOLVED
                    'Dim pic = "C:/CG_Sports/pics/" + hr(6) + "/" + Trim(CG_firstname) + "-" + Trim(CG_name) + ".png"

                    'If My.Computer.FileSystem.FileExists(pic.Replace("/", "\")) Then
                    '    tmpl.Fields.Add(New TemplateField("pic", pic))
                    'Else
                    '    tmpl.Fields.Add(New TemplateField("pic", vbNullString))
                    'End If
                    '_________________________________________________________________________________________________

                    'calls template
                    'Return_CG(_Caspar.CG_Add(Channel1, 20, "Fussball/einzelname", tmpl, True))
                    _Caspar.CG_Add(Channel1, 20, "Fussball/einzelname", tmpl, True)
                    ' somethin is on screen
                    cg_on = True
                    ' and its a singlename (to allow second line functions)
                    singlename_on = True
                End If
                Label_actual_name.Text = CG_nr + " " + CG_firstname + " " + CG_name : Label_actual_name.Visible = True
            End If
        Next
        'second line functions
        Button_second1.Text = Trim(Str(score_home)) + " : " + Trim(Str(score_away)) : Button_second1.Visible = True
        If Trim(CG_pos) = vbNullString Then Button_second2.Visible = False Else Button_second2.Text = CG_pos : Button_second2.Visible = True
        If Trim(CG_Data1) = vbNullString Then Button_second3.Visible = False Else Button_second3.Text = CG_Data1 : Button_second3.Visible = True
        If Trim(CG_Data2) = vbNullString Then Button_second4.Visible = False Else Button_second4.Text = CG_Data2 : Button_second4.Visible = True
        If Trim(CG_Data3) = vbNullString Then Button_second5.Visible = False Else Button_second5.Text = CG_Data3 : Button_second5.Visible = True
        If Trim(CG_Data4) = vbNullString Then Button_second6.Visible = False Else Button_second6.Text = CG_Data4 : Button_second6.Visible = True
        If Trim(CG_Data5) = vbNullString Then Button_second7.Visible = False Else Button_second7.Text = CG_Data5 : Button_second7.Visible = True
    End Sub

    Private Sub Yellow_Click(sender As Object, e As EventArgs) Handles yellow.Click
        If singlename_on = True Then
            If cardon = False Then
                Dim tmpl As Template = New Template
                tmpl.Fields.Add(New TemplateField("card", yellowcard))
                _Caspar.CG_Update(20, tmpl)
                yellow.BackColor = Color.OrangeRed
                cg_on = True
                cardon = True
            Else
                Clear_card()
            End If
        End If
    End Sub

    Private Sub Red_Click(sender As Object, e As EventArgs) Handles red.Click
        If singlename_on = True Then
            If cardon = False Then
                Dim tmpl As Template = New Template
                tmpl.Fields.Add(New TemplateField("card", redcard))
                _Caspar.CG_Update(20, tmpl)
                red.BackColor = Color.OrangeRed
                cg_on = True
                cardon = True
            Else
                Clear_card()
            End If
        End If
    End Sub

    Private Sub Yellowred_Click(sender As Object, e As EventArgs) Handles yellowred.Click
        If singlename_on = True Then
            If cardon = False Then
                Dim tmpl As Template = New Template
                tmpl.Fields.Add(New TemplateField("card", yellowredcard))
                _Caspar.CG_Update(20, tmpl)
                yellowred.BackColor = Color.OrangeRed
                cg_on = True
                cardon = True
            Else
                Clear_card()
            End If
        End If
    End Sub

    Sub Clear_card()
        'emptycard = " "
        Dim tmpl As Template = New Template
        tmpl.Fields.Add(New TemplateField("card", emptycard))
        _Caspar.CG_Update(20, tmpl)
        yellow.BackColor = Color.WhiteSmoke
        red.BackColor = Color.WhiteSmoke
        yellowred.BackColor = Color.WhiteSmoke
        cardon = False
    End Sub

    Private Sub Button_second1_Click(sender As Object, e As EventArgs) Handles Button_second1.Click, Button_second2.Click, Button_second3.Click, Button_second4.Click, Button_second5.Click, Button_second6.Click, Button_second7.Click
        ' loads result in second line and updates CG
        Dim tmpl As Template = New Template
        Label_secondrow.Text = ""
        Select Case True
            Case sender Is Button_second1
                ' loads result in second line and updates CG
                tmpl.Fields.Add(New TemplateField("_zeile2", Trim(Str(score_home)) + " : " + Trim(Str(score_away))))
                Label_secondrow.Text = Trim(Str(score_home)) + " : " + Trim(Str(score_away)) : Label_secondrow.Visible = True
            Case sender Is Button_second2
                tmpl.Fields.Add(New TemplateField("_zeile2", CG_pos))
                Label_secondrow.Text = CG_pos : Label_secondrow.Visible = True
            Case sender Is Button_second3
                tmpl.Fields.Add(New TemplateField("_zeile2", CG_Data1))
                Label_secondrow.Text = CG_Data1 : Label_secondrow.Visible = True
            Case sender Is Button_second4
                tmpl.Fields.Add(New TemplateField("_zeile2", CG_Data2))
                Label_secondrow.Text = CG_Data2 : Label_secondrow.Visible = True
            Case sender Is Button_second5
                tmpl.Fields.Add(New TemplateField("_zeile2", CG_Data3))
                Label_secondrow.Text = CG_Data3 : Label_secondrow.Visible = True
            Case sender Is Button_second6
                tmpl.Fields.Add(New TemplateField("_zeile2", CG_Data4))
                Label_secondrow.Text = CG_Data4 : Label_secondrow.Visible = True
            Case sender Is Button_second7
                tmpl.Fields.Add(New TemplateField("_zeile2", CG_Data5))
                Label_secondrow.Text = CG_Data5 : Label_secondrow.Visible = True
        End Select
        tmpl.Fields.Add(New TemplateField("hflagge", ""))
        tmpl.Fields.Add(New TemplateField("aflagge", ""))
        _Caspar.CG_Update(20, tmpl)
    End Sub

    Private Sub Button_referee_1_Click(sender As Object, e As EventArgs) Handles button_referee_1.Click, button_referee_2.Click, button_referee_3.Click, button_referee_4.Click
        Dim referee_image = "C:/CG_Sports/flags/referee.png"
        Dim tmpl As Template = New Template

        Select Case True
            Case sender Is button_referee_1
                If schiri1_state = True Then
                    Fadeout()
                    schiri1_state = False
                    Return
                Else
                    Clear_all_on()
                    'One_after_the_other()
                    schiri1_state = True
                    button_referee_1.BackColor = Color.LightSalmon
                    tmpl.Fields.Add(New TemplateField("_name", Trim(refname(1))))
                    tmpl.Fields.Add(New TemplateField("_zeile2", Trim(reffunction(1))))
                    Label_actual_name.Visible = True
                    Label_secondrow.Visible = True
                    Label_actual_name.Text = Trim(refname(1))
                    Label_secondrow.Text = Trim(reffunction(1))
                End If
            Case sender Is button_referee_2
                If schiri2_state = True Then
                    Fadeout()
                    schiri2_state = False
                    Return
                Else
                    Clear_all_on()
                    ' One_after_the_other()
                    schiri2_state = True
                    button_referee_2.BackColor = Color.LightSalmon
                    tmpl.Fields.Add(New TemplateField("_name", Trim(refname(2))))
                    tmpl.Fields.Add(New TemplateField("_zeile2", Trim(reffunction(2))))
                    Label_actual_name.Visible = True
                    Label_secondrow.Visible = True
                    Label_actual_name.Text = Trim(refname(2))
                    Label_secondrow.Text = Trim(reffunction(2))
                End If
            Case sender Is button_referee_3
                If schiri3_state = True Then
                    Fadeout()
                    schiri3_state = False
                    Return
                Else
                    Clear_all_on()
                    'One_after_the_other()
                    schiri3_state = True
                    button_referee_3.BackColor = Color.LightSalmon
                    tmpl.Fields.Add(New TemplateField("_name", Trim(refname(3))))
                    tmpl.Fields.Add(New TemplateField("_zeile2", Trim(reffunction(3))))
                    Label_actual_name.Visible = True
                    Label_secondrow.Visible = True
                    Label_actual_name.Text = Trim(refname(3))
                    Label_secondrow.Text = Trim(reffunction(3))
                End If
            Case sender Is button_referee_4
                If schiri4_state = True Then
                    Fadeout()
                    schiri4_state = False
                    Return
                Else
                    Clear_all_on()
                    ' One_after_the_other()
                    schiri4_state = True
                    button_referee_4.BackColor = Color.LightSalmon
                    tmpl.Fields.Add(New TemplateField("_name", Trim(refname(4))))
                    tmpl.Fields.Add(New TemplateField("_zeile2", Trim(reffunction(4))))
                    Label_actual_name.Visible = True
                    Label_secondrow.Visible = True
                    Label_actual_name.Text = Trim(refname(4))
                    Label_secondrow.Text = Trim(reffunction(4))
                End If
        End Select

        tmpl.Fields.Add(New TemplateField("logo1", referee_image))

        'calls template
        'Return_CG(_Caspar.CG_Add(Channel1, 20, "Fussball/einzelname", tmpl, True))
        _Caspar.CG_Add(Channel1, 20, "Fussball/einzelname", tmpl, True)

        ' somethin is on screen
        cg_on = True
        ' and its a singlename (to allow second line functions)
        singlename_on = True
    End Sub

    Private Sub Buttonscorebug_Click(sender As Object, e As EventArgs) Handles buttonscorebug.Click
        If scorebug_on = False Then
            'result small scorebug
            Dim tmpl As Template = New Template
            tmpl.Fields.Add(New TemplateField("_home", homek.Text))
            tmpl.Fields.Add(New TemplateField("_away", awayk.Text))
            tmpl.Fields.Add(New TemplateField("_time", clock.Text))
            tmpl.Fields.Add(New TemplateField("_hpoint", Trim(Str(score_home))))
            tmpl.Fields.Add(New TemplateField("_apoint", Trim(Str(score_away))))
            tmpl.Fields.Add(New TemplateField("hflagge", hflagge))
            tmpl.Fields.Add(New TemplateField("aflagge", aflagge))
            tmpl.Fields.Add(New TemplateField("_homecolor1", h1color))

            tmpl.Fields.Add(New TemplateField("_awaycolor1", a1color))
            tmpl.Fields.Add(New TemplateField("_homecolor2", h2color))
            tmpl.Fields.Add(New TemplateField("_awaycolor2", a2color))
            If RadioButton1.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(4)))
            If RadioButton2.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(5)))
            If RadioButton3.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(6)))
            If RadioButton4.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(7)))

            Return_CG(_Caspar.CG_Add(Channel2, 21, "Fussball/resklein", tmpl, True))
            scorebug_on = True
            buttonscorebug.BackColor = Color.LightSalmon

        Else
            If scorebug_on = True Then
                ' if template has no outro then fade out with a transparent image
                '_Caspar.Execute("PLAY 2-21 vbNullString        TRANSPARENTvbNullString         MIX 6")

                ' start outro
                _Caspar.CG_Stop(Channel2, 21)
            End If
            scorebug_on = False
            buttonscorebug.BackColor = Color.FromArgb(224, 224, 224)
        End If
    End Sub

    Sub Fadeout()
        'if something is on on Layer 20 then fade out, set button status back
        '_Caspar.Execute("PLAY 1-20 vbNullString        TRANSPARENTvbNullString         MIX 6")
        _Caspar.Execute("stop 1-20")
        '_Caspar.CG_Stop(channel1, 20)
        result_big.BackColor = Color.FromArgb(224, 224, 224)
        'Titel_1.BackColor = Color.FromArgb(224, 224, 224)
        'Titel_2.BackColor = Color.FromArgb(224, 224, 224)
        button_referee_1.BackColor = Color.FromArgb(224, 224, 224)
        button_referee_2.BackColor = Color.FromArgb(224, 224, 224)
        button_referee_3.BackColor = Color.FromArgb(224, 224, 224)
        button_referee_4.BackColor = Color.FromArgb(224, 224, 224)
        Clear_all_on()
    End Sub

    Private Sub Status_Rasen_Click(sender As Object, e As EventArgs) Handles Status_Rasen.Click
        'shows grass background for CG_Text
        If Fussball_Rasen_on = False Then
            _Caspar.Execute(String.Format("PLAY {0}-{1} ""{2}""", Channel1, 5, "RASEN"))
            Fussball_Rasen_on = True
            Fussball_Image_on = False
            Fussball_Video_on = False
            Status_Rasen.BackColor = Color.LightGreen
            Status_Fussball.BackColor = Color.FromArgb(224, 224, 224)
            Status_Video.BackColor = Color.FromArgb(224, 224, 224)
        Else
            _Caspar.CG_Clear(5)
            Fussball_Image_on = False
            Fussball_Rasen_on = False
            Fussball_Video_on = False
            Status_Rasen.BackColor = Color.FromArgb(224, 224, 224)
            Status_Fussball.BackColor = Color.FromArgb(224, 224, 224)
            Status_Video.BackColor = Color.FromArgb(224, 224, 224)
        End If
    End Sub

    Private Sub Status_Fussball_Click(sender As Object, e As EventArgs) Handles Status_Fussball.Click
        'shows handball stadium background for CG_Text
        If Fussball_Image_on = False Then
            _Caspar.Execute(String.Format("PLAY {0}-{1} ""{2}""", Channel1, 5, "fussball"))
            Fussball_Image_on = True
            Fussball_Rasen_on = False
            Fussball_Video_on = False
            Status_Fussball.BackColor = Color.LightGreen
            Status_Rasen.BackColor = Color.FromArgb(224, 224, 224)
            Status_Video.BackColor = Color.FromArgb(224, 224, 224)
        Else
            _Caspar.CG_Clear(5)
            Fussball_Image_on = False
            Fussball_Rasen_on = False
            Fussball_Video_on = False
            Status_Rasen.BackColor = Color.FromArgb(224, 224, 224)
            Status_Fussball.BackColor = Color.FromArgb(224, 224, 224)
            Status_Video.BackColor = Color.FromArgb(224, 224, 224)
        End If
    End Sub

    Private Sub Status_Video_Click(sender As Object, e As EventArgs) Handles Status_Video.Click
        'shows grass background for CG_Text
        If Fussball_Video_on = False Then
            _Caspar.Execute(String.Format("PLAY {0}-{1} ""{2}""  {3} ", 1, 5, "fb", "Loop"))
            Fussball_Video_on = True
            Fussball_Rasen_on = False
            Fussball_Image_on = False
            Status_Video.BackColor = Color.LightGreen
            Status_Rasen.BackColor = Color.FromArgb(224, 224, 224)
            Status_Fussball.BackColor = Color.FromArgb(224, 224, 224)
        Else
            _Caspar.CG_Clear(5)
            Fussball_Video_on = False
            Fussball_Rasen_on = False
            Fussball_Image_on = False
            Status_Video.BackColor = Color.FromArgb(224, 224, 224)
            Status_Rasen.BackColor = Color.FromArgb(224, 224, 224)
            Status_Fussball.BackColor = Color.FromArgb(224, 224, 224)
        End If
    End Sub

    Private Sub Clock_start_Click(sender As Object, e As EventArgs) Handles clock_start.Click
        'starts clock
        Starttime()
        If uhrlaeuft = 1 Then
            'disables period radiobuttons to avoid errors
            'RadioButton1.Enabled = False
            'RadioButton2.Enabled = False
            'RadioButton3.Enabled = False
            'RadioButton4.Enabled = False
        End If
    End Sub

    Sub Starttime()
        If time_is_over = False Then
            Timer1.Start()
            Timer2.Start()
            Me.StopWatch1.Start()
            uhrlaeuft = 1

            'sets CG period_text
            If RadioButton1.Checked = True Then period_text = settings(4)
            If RadioButton2.Checked = True Then period_text = settings(5)
            If RadioButton3.Checked = True Then period_text = settings(6)
            If RadioButton4.Checked = True Then period_text = settings(7) : time_is_over = True

            'if no checkbox is checked, sets first checkbox as true, errorhandling
            If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False Then
                RadioButton1.Checked = True
                period_text = settings(4)
            End If
        End If
    End Sub

    Private Sub Minus_Click(sender As Object, e As EventArgs) Handles minus.Click
        correct_time = correct_time + 1
    End Sub

    Private Sub Plus_Click(sender As Object, e As EventArgs) Handles plus.Click
        correct_time = correct_time - 1
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton3.CheckedChanged, RadioButton4.CheckedChanged
        Dim tmpl As Template = New Template
        If RadioButton1.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(4)))
        If RadioButton2.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(5)))
        If RadioButton3.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(6)))
        If RadioButton4.Checked = True Then tmpl.Fields.Add(New TemplateField("_period", settings(7)))
        _Caspar.CG_Update(21, tmpl)
    End Sub

    Sub Clear_all_on()
        cg_on = False
        singlename_on = False
        'Clear_card()
        result_big_state = False
        Titel_state = False
        Titel2_state = False
        tabellen_state = False

        lineup1_on = False
        lineup2_on = False
        ri_Data.Text = vbNullString
        ri_Number.Text = vbNullString
        ri_message.Text = vbNullString
        'Titel_1.BackColor = Color.FromArgb(224, 224, 224)
        'Titel_2.BackColor = Color.FromArgb(224, 224, 224)
        result_big.BackColor = Color.FromArgb(224, 224, 224)
        'Button_lineup_home.BackColor = Color.FromArgb(224, 224, 224)
        'Button_lineup_away.BackColor = Color.FromArgb(224, 224, 224)
        Label_actual_name.Text = vbNullString
        Label_secondrow.Text = vbNullString
    End Sub

    Sub Return_CG(Caspar_Command)
        ' can display a return information from Caspar

        ' first the variable ri as ReturnInfo has to be declared in the main declarations: Dim ri As ReturnInfo
        ' this sub,  Return_CG(Caspar_Command), takes the caspar string and returns errors as message box
        ' example: Return_CG(_Caspar.CG_Add(channel1, 20, "Fussball/hb_lineup", tmpl, True))

        ' if you dont want return infos, just use the caspar command allone
        ' example: _Caspar.CG_Add(channel1, 20, "Fussball/hb_lineup", tmpl, True)

        ri = Caspar_Command
        'ri returns ri.Message, ri.Data, ri.Nubmer

        'bottom statusbar display 
        ri_message.Text = ri.Message
        ri_Data.Text = ri.Data
        ri_Number.Text = ri.Number

        'errors will call up a messagebox
        Select Case ri.Number
            Case 0
                Status_CGConnected.BackColor = Color.LightCoral
                Status_CGConnected.Text = "not connected"
            Case 100  'Information about an event. ???
            Case 101  'Information about an event. A line of data Is being returned. ???
         'Successful        
            Case 200    'OK	- The command has been executed And several lines of data (seperated by \r\n) are being returned (terminated with an additional \r\n)
            Case 201    'OK	- The command has been executed And data (terminated by \r\n) Is being returned.
            Case 202    'OK	- The command has been executed.
         '400 ERROR	- Command Not understood
            Case 400
                MsgBox("400 ERROR	- Command Not understood")
            Case 401
                MsgBox("401 [command] ERROR	- Illegal video_channel")
            Case 402
                MsgBox("402 [command] ERROR	- Parameter missing")
            Case 403
                MsgBox("403 [command] ERROR	- Illegal parameter")
            Case 404
                MsgBox("404 [command] ERROR	- Media file Not found")
        'Server Error
            Case 500
                MsgBox("500 FAILED	- Internal server error")
            Case 501
                MsgBox("501 [command] FAILED	- Internal server error")
            Case 502
                MsgBox("502 [command] FAILED	- Media file unreadable")
        End Select
    End Sub

End Class
