VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   ".:[SDM]:.  coded by chown"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6555
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox log 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   6375
   End
   Begin MSWinsockLib.Winsock w 
      Left            =   120
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Download"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   250
      Left            =   5520
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin ComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.TextBox url 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------'
'- Simple Download Manager v0.1.3 -'
'-      Coded by chown, 2006      -'
'- Email: Prof.J.FrinkATgmail.com -'
'----------------------------------'
'
' If you use this code in your own project, remember
' you're not obligated to credit me, but it would be nice
'
' Have fun!
' - chown

'// Declare various public variables
Public rhost, rport, rpath, fname, fsize, cindex, downloadpath, reqsize As Boolean

Private Sub Command1_Click()
    '// Exit when the use clicks 'Abort'
    End
End Sub

Private Sub Command2_Click()
    
     '// Start a download (or resume if it already exists)
    If Command2.Caption = "Download" Or Command2.Caption = "Resume" Then
        
        '// declare various local vars
        Dim spl, addr, hspl, revspl, fn
        
        '// Pause downloading if Command2 is clicked again
        Command2.Caption = "Pause"
        
        '// Make sure the user hasn't specified anything other than the HTTP protocol
        '// (eg ftp://, file://, mms://, etc)
        addr = Replace(url.Text, "http://", "")
        If InStr(1, addr, "://") Then
            logg "Invalid protocol"
            MsgBox "Invalid protocol (only HTTP is supported)", vbExclamation
            Exit Sub
        End If
        
        '// Parse the user-specified URL for the host, path, filename and port
        revspl = Split(StrReverse(addr), "/")
        fname = StrReverse(revspl(0))
        spl = Split(addr, "/")
        hspl = Split(spl(0) & ":80", ":")
        rhost = hspl(0)
        rport = hspl(1)
        rpath = Mid(addr, Len(spl(0)) + 1)
        
        '// Output data extracted from the URL
        logg "Local fname: " & fname
        logg "Remote host: " & rhost
        logg "Remote port: " & rport
        logg "Remote path: " & rpath
        
        '// Check to see wether the .downloading file already exists
        '// If so, resume the download. Otherwise start a new one
        fn = downloadpath & "\" & fname & ".downloading"
        If FileExists(fn) Then
            cindex = FileLen(fn)
            logg "Resuming download..."
        Else
            cindex = 0
            logg "Starting new download..."
        End If
        
        logg "Connecting to " & rhost & ":" & rport & "..."
        
        '// Connect to the remote host to grab the file size
        '// (make sure the socket is ready first, by closing it)
        w.Close
        w.Connect rhost, rport
        
    ElseIf Command2.Caption = "Pause" Then
    
        '// Pause the download simply by closing the socket
        w.Close
        Command2.Caption = "Resume"
    End If
End Sub

Function logg(txt)
    '// Chop off one kb if the log window excedes 10 kb in size
    If Len(logg.Text) > 10240 Then log.Text = Right(log.Text, 9216)
    '// Timestamp and add 'txt' to the log window
    log.Text = log.Text & "(" & DateTime.Time & ") " & txt & vbCrLf
    '// Scroll the text field down
    log.SelStart = Len(log.Text)
    log.SelLength = 1
End Function

Private Sub Form_Load()
    init
End Sub

Sub init()
    '// Init
    downloadpath = App.Path
    reqsize = True
End Sub

'// When w (the socket) connects, send the appropriate HTTP request
Private Sub w_Connect()
    Dim request
    If reqsize = True Then
        request = Replace(Replace(HTTPGET, "{HOST}", rhost), "{PATH}", url.Text)
        logg "Connected to " & w.RemoteHost & ":" & w.RemotePort
        logg "Requesting file size..."
        logg "request:" & vbCrLf & request
        w.SendData request
        logg "Size request sent. Awaiting response..."
    Else
        logg "Requesting file from position " & cindex - fsize & "..."
        request = Replace(Replace(Replace(HTTPGETRANGE, "{HOST}", rhost), "{PATH}", url.Text), "{RANGE}", cindex - fsize)
        logg "request:" & vbCrLf & request
        w.SendData request
        logg "Downloading..."
    End If
End Sub

'// Handle data arriving on w
Private Sub w_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String
    Dim crspl, i, lenspl, dblcrspl, flen, cmpname, cmpext, cmpspl, pcnt
    
    '// If reqsize is true, then parse the HTTP response for the Content-Length value
    If reqsize = True Then
        w.GetData dat, , 1024
        
        '// HTTP 404 recieved from the server; abort.
        If LCase(Mid(dat, 1, 12)) = "http/1.1 404" Then
            logg "Recieved HTTP not found. Aborting."
            MsgBox "404: Not found", vbExclamation, ""
            w.Close
            Command2.Caption = "Download"
            Exit Sub
        End If
        crspl = Split(dat & vbCrLf & "[EOF]", vbCrLf)
        i = 0
        While crspl(i) <> "[EOF]"
            If LCase(Mid(crspl(i) & "XXXXXXXXXXXXXXX", 1, 15)) = "content-length:" Then
                lenspl = Split(crspl(i) & ": 0", ": ")
                fsize = lenspl(1)
                logg "File size is: " & (fsize / 1024) & " Kb"
                reqsize = False
            End If
            i = i + 1
        Wend
        
        '// Failed to parse Content-Length value. Invalid data. Abort.
        If reqsize = True Then
            MsgBox "Recieved unsupported or non-HTTP data from the server. Aborting.", vbExclamation, ""
            w.Close
            Command2.Caption = "Download"
            Exit Sub
        End If
        w.Close
        logg "Reconnecting to server..."
        w.Connect rhost, rport
        Exit Sub
    Else
    
        '// Otherwise, if reqsize is not true, then append the incomming data to the currently downloading file.
        w.GetData dat
    End If
    
    '// Remove the HTTP header
    If LCase(Mid(dat, 1, 12)) = "http/1.1 206" Then
        dblcrspl = Split(dat, vbCrLf & vbCrLf, 2)
        dat = dblcrspl(1)
    End If
    
    '// Append the currently downloading file
    Open downloadpath & "\" & fname & ".downloading" For Binary Access Write As #1
        Put #1, LOF(1) + 1, dat 'Mid(dat, 1, Len(dat) - 2)
    Close #1
    
    '// Check the file size against the servers file size to determine wether it has completed downloading
    flen = FileLen(downloadpath & "\" & fname & ".downloading")
    If Val(flen) = Val(fsize) Then
        
        '// Download has completed. Notify user, rename .downloading file, etc
        Form1.Caption = title
        pb.Value = 0
        Command2.Caption = "Download"
        
        logg "Download completed."
        
        '// If the .downloading file cannot be renamed because a file of that name already exists,
        '// then append a number to the file name. Eg. foobar(1).exe. (Increase the number until
        '// the file name is unique)
        fcmpname = downloadpath & "\" & fname
        If FileExists(fcmpname) Then
            cmpspl = Split(fname & ".", ".")
            cmpname = cmpspl(0)
            cmpext = cmpspl(1)
            i = 1
            fcmpname = downloadpath & "\" & cmpname & "(" & i & ")" & "." & cmpext
            While FileExists(fcmpname)
                i = i + 1
                fcmpname = downloadpath & "\" & cmpname & "(" & i & ")" & "." & cmpext
            Wend
        End If
        
        '// Rename the .downloading file to the name determined above
        Name downloadpath & "\" & fname & ".downloading" As fcmpname
        
        '// Notify user upon completion
        MsgBox "Download completed.", vbInformation, ""
        w.Close
        Exit Sub
    End If
    
    '// Calculate & display progress in %
    pcnt = flen / fsize * 100
    pb.Value = Math.Round(pcnt)
    Form1.Caption = title & ". [downloading: " & Round(flen / fsize * 100, 2) & "%]"
End Sub

Private Sub w_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '// In the event of a socket error, abort the download. (resuming is still possible)
    logg "Socket Error: " & Description
    logg "Aborting."
    Command2.Caption = "Download"
    w.Close
End Sub

'// Simple function to determine if a specified file exists
Public Function FileExists(fn) As Boolean
    If fn = "" Or Right(fn, 1) = "\" Then
      FileExists = False
      Exit Function
    End If
    FileExists = (Dir(fn) <> "")
End Function
