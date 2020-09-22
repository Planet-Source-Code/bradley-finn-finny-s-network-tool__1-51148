VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   Caption         =   "Finny's Network tool"
   ClientHeight    =   4920
   ClientLeft      =   2580
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   3
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Ping"
      TabPicture(0)   =   "ping.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ProgressBar1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "size"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "windows"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "times(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "host(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Net Send"
      TabPicture(1)   =   "ping.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command14"
      Tab(1).Control(1)=   "Command13"
      Tab(1).Control(2)=   "b1"
      Tab(1).Control(3)=   "b2"
      Tab(1).Control(4)=   "Command4"
      Tab(1).Control(5)=   "b3"
      Tab(1).Control(6)=   "ProgressBar2"
      Tab(1).Control(7)=   "Label11(4)"
      Tab(1).Control(8)=   "Label5(1)"
      Tab(1).Control(9)=   "Label7"
      Tab(1).Control(10)=   "Label8"
      Tab(1).Control(11)=   "Message"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Virtual Drives"
      TabPicture(2)   =   "ping.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Dir2"
      Tab(2).Control(1)=   "Drive2"
      Tab(2).Control(2)=   "v2"
      Tab(2).Control(3)=   "Command7"
      Tab(2).Control(4)=   "v3"
      Tab(2).Control(5)=   "Command8"
      Tab(2).Control(6)=   "Label11(3)"
      Tab(2).Control(7)=   "Label4(2)"
      Tab(2).Control(8)=   "Label3(2)"
      Tab(2).Control(9)=   "Label6(2)"
      Tab(2).Control(10)=   "Label12"
      Tab(2).Control(11)=   "Label13"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Open shares"
      TabPicture(3)   =   "ping.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label9"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label1(3)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label1(2)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label6(1)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label11(6)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "bb"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "hs"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "host(2)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Command5"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "host(1)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Create Shares"
      TabPicture(4)   =   "ping.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label14"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label15"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label17"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label4(1)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label3(1)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label11(2)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "c1"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Command9"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "c3"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Command10"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Dir1"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Drive1"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).ControlCount=   12
      TabCaption(5)   =   "User Config"
      TabPicture(5)   =   "ping.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label10"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label6(0)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label1(1)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label2(3)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label1(5)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label2(4)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Label11(0)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Command6"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "deluser"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Command3"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "pw"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "user"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "Command2"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "nuser(1)"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "npw(1)"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).ControlCount=   15
      TabCaption(6)   =   "FTP"
      TabPicture(6)   =   "ping.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Command12"
      Tab(6).Control(1)=   "f6"
      Tab(6).Control(2)=   "f5"
      Tab(6).Control(3)=   "f4"
      Tab(6).Control(4)=   "f3"
      Tab(6).Control(5)=   "Command11"
      Tab(6).Control(6)=   "f2"
      Tab(6).Control(7)=   "f1"
      Tab(6).Control(8)=   "Label11(1)"
      Tab(6).Control(9)=   "Label23"
      Tab(6).Control(10)=   "Label22"
      Tab(6).Control(11)=   "Label21"
      Tab(6).Control(12)=   "Label20"
      Tab(6).Control(13)=   "Label19"
      Tab(6).Control(14)=   "Label18"
      Tab(6).Control(15)=   "Label6(4)"
      Tab(6).Control(16)=   "Label6(3)"
      Tab(6).ControlCount=   17
      Begin VB.CommandButton Command14 
         Caption         =   "Enable Net Send"
         Height          =   375
         Left            =   -71160
         TabIndex        =   83
         Top             =   3720
         Width           =   1935
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Disable Net Send"
         Height          =   375
         Left            =   -73440
         TabIndex        =   82
         Top             =   3720
         Width           =   2055
      End
      Begin VB.DirListBox Dir2 
         Height          =   1215
         Left            =   -74160
         TabIndex        =   79
         Top             =   1920
         Width           =   4215
      End
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   -74160
         TabIndex        =   78
         Top             =   1320
         Width           =   2055
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   -74280
         TabIndex        =   75
         Top             =   1080
         Width           =   2055
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   -74280
         TabIndex        =   74
         Top             =   1800
         Width           =   4215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Open FTP"
         Height          =   375
         Left            =   -70320
         TabIndex        =   73
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox f6 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -71880
         PasswordChar    =   "*"
         TabIndex        =   71
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox f5 
         Height          =   285
         Left            =   -73680
         TabIndex        =   70
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox f4 
         Height          =   285
         Left            =   -71280
         TabIndex        =   67
         Text            =   "21"
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox f3 
         Height          =   285
         Left            =   -73680
         TabIndex        =   66
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CommandButton Command11 
         Caption         =   "open FTP"
         Height          =   375
         Left            =   -70320
         TabIndex        =   64
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox f2 
         Height          =   285
         Left            =   -71280
         TabIndex        =   63
         Text            =   "21"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox f1 
         Height          =   285
         Left            =   -73680
         TabIndex        =   60
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Unshare"
         Height          =   375
         Left            =   -71520
         TabIndex        =   56
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox c3 
         Height          =   285
         Left            =   -74280
         TabIndex        =   55
         Top             =   4320
         Width           =   2655
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Share"
         Height          =   495
         Left            =   -69600
         TabIndex        =   54
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox c1 
         Height          =   285
         Left            =   -74280
         TabIndex        =   52
         Top             =   3360
         Width           =   3735
      End
      Begin VB.TextBox npw 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   -72240
         PasswordChar    =   "*"
         TabIndex        =   44
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox nuser 
         Height          =   285
         Index           =   1
         Left            =   -73920
         TabIndex        =   43
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Create New User"
         Height          =   375
         Left            =   -70560
         TabIndex        =   42
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox user 
         Height          =   285
         Left            =   -73920
         TabIndex        =   41
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox pw 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -72240
         PasswordChar    =   "*"
         TabIndex        =   40
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Change Password"
         Height          =   375
         Left            =   -70560
         TabIndex        =   39
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox deluser 
         Height          =   285
         Left            =   -73920
         TabIndex        =   38
         Top             =   3240
         Width           =   3135
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Delete User"
         Height          =   375
         Left            =   -70560
         TabIndex        =   37
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox host 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   32
         Top             =   1800
         Width           =   3975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Open Share"
         Height          =   495
         Left            =   5040
         TabIndex        =   31
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox host 
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   30
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox hs 
         Height          =   285
         Left            =   3000
         TabIndex        =   29
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton bb 
         Caption         =   "Open Hidden Share"
         Height          =   495
         Left            =   5040
         TabIndex        =   28
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox v2 
         Height          =   285
         Left            =   -74160
         TabIndex        =   24
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Create Virtual Drive"
         Height          =   495
         Left            =   -69600
         TabIndex        =   23
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox v3 
         Height          =   285
         Left            =   -74160
         TabIndex        =   22
         Top             =   4320
         Width           =   3495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Remove Virtual Drive"
         Height          =   495
         Left            =   -70440
         TabIndex        =   21
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox b1 
         Height          =   285
         Left            =   -74040
         TabIndex        =   16
         Top             =   1740
         Width           =   1695
      End
      Begin VB.TextBox b2 
         Height          =   285
         Left            =   -72120
         TabIndex        =   15
         Text            =   "1"
         Top             =   1740
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Send"
         Height          =   375
         Left            =   -69840
         TabIndex        =   14
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox b3 
         Height          =   285
         Left            =   -74040
         TabIndex        =   12
         Top             =   2340
         Width           =   5295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ping!"
         Height          =   495
         Left            =   -70320
         TabIndex        =   6
         Top             =   1980
         Width           =   1455
      End
      Begin VB.TextBox host 
         Height          =   285
         Index           =   0
         Left            =   -74040
         TabIndex        =   5
         Top             =   1740
         Width           =   1695
      End
      Begin VB.TextBox times 
         Height          =   285
         Index           =   0
         Left            =   -72120
         TabIndex        =   4
         Top             =   1740
         Width           =   1575
      End
      Begin VB.TextBox windows 
         Height          =   285
         Left            =   -74040
         TabIndex        =   3
         Text            =   "1"
         Top             =   2460
         Width           =   1695
      End
      Begin VB.TextBox size 
         Height          =   285
         Left            =   -72120
         TabIndex        =   2
         Text            =   "65500"
         Top             =   2460
         Width           =   1575
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   -74040
         TabIndex        =   1
         Top             =   2940
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   -74040
         TabIndex        =   13
         Top             =   2940
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label11 
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   6
         Left            =   4560
         TabIndex        =   93
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   5
         Left            =   -70440
         TabIndex        =   92
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   4
         Left            =   -70440
         TabIndex        =   91
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   3
         Left            =   -70440
         TabIndex        =   90
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   2
         Left            =   -70440
         TabIndex        =   89
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   1
         Left            =   -70440
         TabIndex        =   88
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   0
         Left            =   -70440
         TabIndex        =   87
         Top             =   4680
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "New Password"
         Height          =   255
         Index           =   4
         Left            =   -72240
         TabIndex        =   86
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "User Name"
         Height          =   255
         Index           =   5
         Left            =   -73920
         TabIndex        =   85
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Select Folder:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   -74160
         TabIndex        =   81
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select drive:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   -74160
         TabIndex        =   80
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select drive:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   -74280
         TabIndex        =   77
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Select Folder:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   -74280
         TabIndex        =   76
         Top             =   1560
         Width           =   1200
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Password"
         Height          =   255
         Left            =   -71880
         TabIndex        =   72
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Username"
         Height          =   255
         Left            =   -73680
         TabIndex        =   69
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "Port"
         Height          =   255
         Left            =   -71160
         TabIndex        =   68
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Computer Name"
         Height          =   255
         Left            =   -73560
         TabIndex        =   65
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Port"
         Height          =   255
         Left            =   -71280
         TabIndex        =   62
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Computer Name"
         Height          =   255
         Left            =   -73560
         TabIndex        =   61
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Access FTP with user and password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74160
         TabIndex        =   59
         Top             =   2280
         Width           =   5055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Access FTP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -74160
         TabIndex        =   58
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label17 
         Caption         =   "Name of share to be unshared"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74280
         TabIndex        =   57
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label Label15 
         Caption         =   "Name of new share (eg Games)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74280
         TabIndex        =   53
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Create Shares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74400
         TabIndex        =   51
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Index           =   3
         Left            =   -71880
         TabIndex        =   50
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "User Name"
         Height          =   255
         Index           =   1
         Left            =   -73560
         TabIndex        =   49
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Existing username"
         Height          =   255
         Index           =   0
         Left            =   -74040
         TabIndex        =   48
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "New password"
         Height          =   255
         Index           =   0
         Left            =   -72240
         TabIndex        =   47
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "User Config"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -74040
         TabIndex        =   46
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Username to delete"
         Height          =   255
         Left            =   -73920
         TabIndex        =   45
         Top             =   3000
         Width           =   3135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Open Shares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   36
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Computername or IP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   35
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Computername or IP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   34
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Hidden Share Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Virtual Drives"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -74160
         TabIndex        =   27
         Top             =   660
         Width           =   5055
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Vitual Drive Letter (eg F)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   26
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label13 
         Caption         =   "Virtual Drive letter to be removed (eg F)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   25
         Top             =   4080
         Width           =   3855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Net Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -74280
         TabIndex        =   20
         Top             =   660
         Width           =   5295
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Computer Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   19
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Number of times to send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72120
         TabIndex        =   18
         Top             =   1500
         Width           =   2175
      End
      Begin VB.Label Message 
         Alignment       =   2  'Center
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   17
         Top             =   2100
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Number of Pings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -72120
         TabIndex        =   11
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Computername or IP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74040
         TabIndex        =   10
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Number of windows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74040
         TabIndex        =   9
         Top             =   2220
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Size (max 65500)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -72120
         TabIndex        =   8
         Top             =   2220
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Pinger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -74160
         TabIndex        =   7
         Top             =   660
         Width           =   5055
      End
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   84
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bb_Click()
Shell ("explorer \\" + (host(2)) + "\" + (hs) + "$")
host(2).Text = ""
End Sub

Private Sub Command1_Click()
ProgressBar1.Max = (windows)
a = 0
X = (windows)
Do
Shell ("ping " + (host(0)) + " -n " + (times(0)) + " -l " + (size))
  X = X - 1
  a = (a + 1)
  ProgressBar1 = (a)
  If X = 0 Then Exit Do
Loop
ProgressBar1 = 0
End Sub

Private Sub Command10_Click()
Shell ("net share " + (c3) + " /delete")
c3.Text = ""
End Sub

Private Sub Command11_Click()
Shell (("C:\Program Files\Internet Explorer\IEXPLORE.EXE ftp://") + (f1) + ":" + (f2))
End Sub

Private Sub Command12_Click()
Shell (("C:\Program Files\Internet Explorer\IEXPLORE.EXE ftp://") + (f5) + ":" + (f6) + "@" + (f3) + ":" + (f4))
End Sub

Private Sub Command13_Click()
Shell "net stop messenger"
End Sub

Private Sub Command14_Click()
Shell "net start messenger"
End Sub

Private Sub Command2_Click()
Shell ("net user /add " + (nuser(1)) + " " + (npw(1)))
nuser(1) = ""
npw(1) = ""
End Sub

Private Sub Command3_Click()
Shell ("net user " + (user) + " " + (pw))
user.Text = ""
pw.Text = ""
End Sub

Private Sub Command4_Click()
ProgressBar2.Max = (b2)
a = 0
X = (b2)
Do
Shell ("net send " + (b1) + " " + (b3))
  X = X - 1
  a = (a + 1)
  ProgressBar2 = (a)
  If X = 0 Then Exit Do
Loop
ProgressBar2 = 0
End Sub

Private Sub Command5_Click()
Shell ("explorer \\" + (host(1)))
host(1).Text = ""
End Sub

Private Sub Timer1_Timer()
If Timer1.Index = 1 Then
ProgressBar2 = 0
ProgressBar1 = 0
Timer1 = False
End If
End Sub

Private Sub Command6_Click()
Shell ("net user /delete " + (deluser))
End Sub

Private Sub Command7_Click()
Shell ("subst " + (v2) + ": " + (Dir2.Path))
v2.Text = ""
End Sub

Private Sub Command8_Click()
Shell ("subst " + (v3) + ": /D")
v3.Text = ""
End Sub

Private Sub Command9_Click()
Shell ("net share " + (c1) + "=" + (Dir1.Path))
c1.Text = ""
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Drive2_Change()
Dir2.Path = Drive2.Drive
End Sub

