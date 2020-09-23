VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectX Tutorial #3 - DirectMusic Intro - by Simon Price - visit www.VBgames.co.uk for more cool VB code and tutorials!"
   ClientHeight    =   1548
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   4344
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1548
   ScaleWidth      =   4344
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   492
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   1212
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   492
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1212
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load MIDI..."
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1212
   End
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   3840
      Top             =   120
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      Filter          =   "*.mid"
   End
   Begin VB.Label lblFilename 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Loaded File = ""midi.mid"""
      Height          =   732
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''
'
'     DIRECTX TUTORIAL #3 - BY SIMON PRICE
'
' DIRECTMUSIC INTRO - LOAD, PLAY, STOP MIDI'S
'
'''''''''''''''''''''''''''''''''''''''''''''''

' this is the main DirectX object
Private DX As New DirectX7
' this loads music, it transfers
' the contents of a file into memory
Private Loader As DirectMusicLoader
' this controls the music
Private Performance As DirectMusicPerformance
' this stores the music in memory
Private Segment As DirectMusicSegment

Private Sub cmdLoad_Click()
On Error Resume Next
' show the open file dialog box
ComDialog.ShowOpen
' if the user pressed cancel, exit sub now
If ComDialog.Filename = "" Then Exit Sub
' load the chosen MIDI file
LoadMIDI ComDialog.Filename
' display filename
lblFilename = ComDialog.FileTitle
' start playing the file
PlayMIDI
End Sub

Private Sub cmdPlay_Click()
' when this button is pressed, play the midi
PlayMIDI
End Sub

Private Sub cmdStop_Click()
' when this button is pressed, stop the midi
StopMIDI
End Sub

' in the Form_Load event, we load all the DirectMusic
' stuff we can so we don't need to do it in the middle
' of the program
Private Sub Form_Load()
On Error Resume Next
' display annoying screen to nag more people to visit my website
MsgBox Caption, vbInformation

' create the loader
Set Loader = DX.DirectMusicLoaderCreate
' create the performance
Set Performance = DX.DirectMusicPerformanceCreate

' start up the performance, telling DirectMusic
' the handle of the form
Performance.Init Nothing, hWnd
' set the port (-1 lets DirectMusic to choose
' the port itself - less work for us)
Performance.SetPort -1, 1
' tell DirectMusic to do all the sound downloading
' stuff itself because we can't be bothered to do it
Performance.SetMasterAutoDownload True

' if there's been an error, report it with message box
If Err.Number <> DD_OK Then MsgBox "ERROR : Could not load DirectMusic!", vbExclamation, "ERROR!"

' now load a default file
LoadMIDI App.Path & "\midi.mid"
' and begin playing it
PlayMIDI
End Sub

' the sub loads a MIDI file and reports
' upon an error
Sub LoadMIDI(Filename As String)
On Error Resume Next
' load from the file given
' what we are actually doing is getting the
' loader to transfer the contents of the midi
' file to the segment object
Set Segment = Loader.LoadSegment(Filename)
' if there was an error, report it with a message box
If Err.Number <> DD_OK Then MsgBox "ERROR : Could not load MIDI file!", vbExclamation, "ERROR!"
End Sub

' this sub plays the currently loaded midi
Sub PlayMIDI()
On Error Resume Next
' tell the performance to being playing the
' music stored in Segment, right away (delay time = 0)
Performance.PlaySegment Segment, 0, 0
End Sub

' this sub stops play
Sub StopMIDI()
On Error Resume Next
' tell the performance object to stop playing
Performance.Stop Segment, Nothing, 0, 0
End Sub
