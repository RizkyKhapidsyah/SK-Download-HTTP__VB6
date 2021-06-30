VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1740
      Top             =   780
   End
   Begin VB.CommandButton btnStartDL 
      Caption         =   "Download!"
      Height          =   375
      Left            =   1620
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   1020
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   1380
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   300
      Width           =   4515
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- coded 2001 by Roman Bobik (roman.bobik@aon.at) --
'-------- any comments to roman.bobik@aon.at --------
'
'downloads an (binary)file over the HTTP-protocol and saves it to a file

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long) 'used for direct memory copy


Const Timeout As Long = 60 'Timeout, if no connection can be established (in seconds)
Dim file() As Byte, Destination As String 'Byte-Array containing current file-contents; Destination file on the harddisk
Dim LastSizeCheck As Long, LastSize As Long 'Data needed for speed-status

Private Sub btnStartDL_Click()
    Inet1.RequestTimeout = Timeout 'set timeout
    ProgressBar1.Max = 1024 '--- here you must fill in the expected file size (in KBs), for using the progressbar...
    Destination = App.Path & "\vb6sp4-runtime.exe" 'set destination file
    Label1.Caption = "Establishing connection..."
    Inet1.Execute "http://www.ssgf.at/docs/haupt.htm", "GET" 'Start the download of the specified file
End Sub

'Determines the UBound of the "file"-Array. If the Array is "Empty" ("Erase file"), it returns -1
Private Function SafeUBoundFile() As Long
    On Error GoTo erro
    SafeUBoundFile = UBound(file)
    Exit Function
erro:
    SafeUBoundFile = -1
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Inet1.Cancel 'cancel on exit
End Sub

Private Sub INet1_StateChanged(ByVal State As Integer)
    Static inProc As Boolean
    If Not inProc Then 'only execute this procedure, if it is the first call (DoEvents in this sub may call this event frequently
        inProc = True
        Debug.Print Timer, State
        Select Case State
            Case icResponseReceived 'Received something
                Dim vtData() As Byte 'Current Chunk
                Label1.Caption = "Downloading " & Inet1.URL & "..."
               
                Do While Inet1.StillExecuting
                    DoEvents
                Loop
                Do
                    DoEvents
                    vtData = Inet1.GetChunk(256, 1)
                    If UBound(vtData) = -1 Then Exit Do 'exit loop, if no Chunk could received
                    ReDim Preserve file(SafeUBoundFile + UBound(vtData) + 1) 'enlarge file-array
                    CopyMemory file(UBound(file) - UBound(vtData)), vtData(0), UBound(vtData) + 1 'copy received Chunk to the file-array
                    If UBound(vtData) <> 255 Then Exit Do 'if the length of the chunk is not 255, then it must be the last chunk of the file
                    
                    Dim tmp As Long
                    tmp = UBound(file) / 1024
                    If tmp > ProgressBar1.Max Then tmp = ProgressBar1.Max 'if KBs is higher then ProgressBar1.Maxy then truncated
                    ProgressBar1.Value = tmp 'update ProgressBar1
                Loop
                
                Label1.Caption = "Download complete."
                MsgBox "Download complete."
                
                Inet1.Cancel
            
                Open Destination For Binary As #1 'Write file-array to destination-file
                Put #1, , file
                Close #1
                Erase file 'free file-array
        End Select
        inProc = False
    End If
End Sub

'Updates the status. Think about it....
Private Sub Timer1_Timer()
    Label2.Caption = Format(SafeUBoundFile / 1024, "#,##0.00 KB") & " @ " & _
        Format((SafeUBoundFile - LastSize) / 1024 / (Timer - LastSizeCheck / 1000), "#,##0.00 KB/s")
    LastSizeCheck = Timer * 1000
    LastSize = SafeUBoundFile
End Sub
