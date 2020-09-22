VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form reg_add 
   Caption         =   "Add to Registry"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "reg_add.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Welcome"
      TabPicture(0)   =   "reg_add.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblUsername"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblUserNameText"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Rename Items"
      TabPicture(1)   =   "reg_add.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label8"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command9"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command8"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Text3"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command6"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command5"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdRecycle_Default"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtRename_Recycle"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdRecycle_Rename"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Command10"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text6"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Command13"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Text4"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Command11"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Text5"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Command12"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "Misc Items"
      TabPicture(2)   =   "reg_add.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check7"
      Tab(2).Control(1)=   "Check6"
      Tab(2).Control(2)=   "Check5"
      Tab(2).Control(3)=   "Check4"
      Tab(2).Control(4)=   "Check3"
      Tab(2).Control(5)=   "Check2"
      Tab(2).Control(6)=   "Check1"
      Tab(2).Control(7)=   "Label13"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Context Menus"
      TabPicture(3)   =   "reg_add.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command15"
      Tab(3).Control(1)=   "Text8"
      Tab(3).Control(2)=   "Text7"
      Tab(3).Control(3)=   "Command3"
      Tab(3).Control(4)=   "Command1"
      Tab(3).Control(5)=   "cmdRecycle_Add"
      Tab(3).Control(6)=   "cmdRecycle_Remove"
      Tab(3).Control(7)=   "Label12"
      Tab(3).Control(8)=   "Label11"
      Tab(3).Control(9)=   "Label10"
      Tab(3).Control(10)=   "Label3"
      Tab(3).Control(11)=   "Label1"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "Browser Skin"
      TabPicture(4)   =   "reg_add.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label2"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Image1"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Command4"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Command2"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Text1"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "File1"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Dir1"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Drive1"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).ControlCount=   8
      Begin VB.CheckBox Check7 
         Caption         =   "Show\Hide C:\ Drive"
         Height          =   255
         Left            =   -71280
         TabIndex        =   51
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Show\Hide Logoff in Start Menu"
         Height          =   495
         Left            =   -71280
         TabIndex        =   50
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Show\Hide Find in Start Menu"
         Height          =   495
         Left            =   -71280
         TabIndex        =   49
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Show\Hide Exit in Start Menu"
         Height          =   495
         Left            =   -74760
         TabIndex        =   48
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Show\Hide Run in Start Menu"
         Height          =   495
         Left            =   -74760
         TabIndex        =   47
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show\Hide Recent Documents in Start Menu"
         Height          =   495
         Left            =   -74760
         TabIndex        =   46
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show\Hide Favorites in Start Menu"
         Height          =   495
         Left            =   -74760
         TabIndex        =   45
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Add it now"
         Height          =   495
         Left            =   -72000
         TabIndex        =   42
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   -71760
         TabIndex        =   41
         Top             =   3960
         Width           =   3615
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   -74760
         TabIndex        =   40
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add it now"
         Height          =   495
         Left            =   -72675
         TabIndex        =   36
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Remove it"
         Height          =   495
         Left            =   -71115
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdRecycle_Add 
         Caption         =   "Add it now"
         Height          =   495
         Left            =   -72675
         TabIndex        =   34
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdRecycle_Remove 
         Caption         =   "Remove it"
         Height          =   495
         Left            =   -71115
         TabIndex        =   33
         Top             =   2280
         Width           =   1215
      End
      Begin VB.DriveListBox Drive1 
         Height          =   360
         Left            =   -73320
         TabIndex        =   31
         Top             =   480
         Width           =   2925
      End
      Begin VB.DirListBox Dir1 
         Height          =   1170
         Left            =   -73320
         TabIndex        =   30
         Top             =   960
         Width           =   2925
      End
      Begin VB.FileListBox File1 
         Height          =   1770
         Hidden          =   -1  'True
         Left            =   -73320
         Pattern         =   "*.bmp"
         System          =   -1  'True
         TabIndex        =   29
         Top             =   2160
         Width           =   2955
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   -73320
         TabIndex        =   28
         Top             =   4380
         Width           =   2955
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Apply"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -69240
         TabIndex        =   27
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -69240
         TabIndex        =   26
         Top             =   4620
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Rename it now"
         Height          =   495
         Left            =   4800
         TabIndex        =   23
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3960
         TabIndex        =   22
         Text            =   "Company"
         Top             =   4080
         Width           =   2895
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Rename it now"
         Height          =   495
         Left            =   4800
         TabIndex        =   21
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3960
         TabIndex        =   20
         Text            =   "Name"
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Defaut"
         Height          =   495
         Left            =   5640
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3960
         TabIndex        =   17
         Text            =   "Name of Control Panel"
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Rename it now"
         Height          =   495
         Left            =   4080
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdRecycle_Rename 
         Caption         =   "Rename it now"
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtRename_Recycle 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Text            =   "Name of Recycle bin"
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton cmdRecycle_Default 
         Caption         =   "Defaut"
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Rename it now"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Default"
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Text            =   "Browser name"
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Text            =   "Outlook Express name"
         Top             =   4080
         Width           =   2775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Default"
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Rename it now"
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "This is only an Example of what can be changed. There is still alot more to be added. BJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   -74760
         TabIndex        =   52
         Top             =   3840
         Width           =   6615
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Path Eg... C:\Windows\Notepad.exe %1"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -71760
         TabIndex        =   44
         Top             =   3600
         Width           =   3615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Add custom path to right click context menus"
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   -74760
         TabIndex        =   39
         Top             =   3240
         Width           =   6615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Open with"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -74760
         TabIndex        =   43
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Add Open with Notepad to right click context menus"
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   -72720
         TabIndex        =   38
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Add Empty Recycle Bin to Right Click Context Menus"
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   -73395
         TabIndex        =   37
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   -73320
         Stretch         =   -1  'True
         Top             =   4860
         Width           =   2955
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Current Skin"
         Height          =   240
         Left            =   -73320
         TabIndex        =   32
         Top             =   4140
         Width           =   2985
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Registered Organization"
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   3960
         TabIndex        =   25
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Registered Owner"
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   3960
         TabIndex        =   24
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Type in new Name for Control Panel"
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   3960
         TabIndex        =   19
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblUserNameText 
         Caption         =   "Information"
         ForeColor       =   &H00008000&
         Height          =   3975
         Left            =   -74640
         TabIndex        =   15
         Top             =   1080
         Width           =   6495
      End
      Begin VB.Label lblUsername 
         Alignment       =   2  'Center
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   -74640
         TabIndex        =   14
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Type in new Name for your Recycle Bin here"
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Type in new Name for Internet Explorer browser"
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1980
         Width           =   2895
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Type in new Name for Outlook Express"
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   3600
         Width           =   2895
      End
   End
End
Attribute VB_Name = "reg_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fN As String
Dim Current_Tab
Dim chk
'----------------------------------
'####################################

    'These values must be Numeric values
    'this is the path
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer
    
    'these are the entries

Private Sub Check1_Click()
    
    ' Show\Hide Favorites
    'NoFavoritesMenu
    'Numeric Value 1
If Check1.Value = 1 Then
     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu")  'let's see what is
    If chk = "Error" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", 1
    ElseIf chk = "1" Then
DeleteValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu"
End If
End If
End Sub
'####################################
Private Sub Check2_Click()

    ' Show\Hide Recent Documents
    'NoRecentDocsMenu
    'Numeric Value 1
If Check2.Value = 1 Then
     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu")  'let's see what is
    If chk = "Error" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", 1
    ElseIf chk = "1" Then
DeleteValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu"
End If
End If
End Sub
'####################################
Private Sub Check3_Click()
    'Show\Hide Run
    'NoRun
    'Numeric Value 1
If Check3.Value = 1 Then
     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun")  'let's see what is
    If chk = "Error" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 1
    ElseIf chk = "1" Then
DeleteValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun"
End If
End If
End Sub
'####################################
Private Sub Check4_Click()
    'Show\Hide Exit
    'NoClose
    'Numeric Value 1
If Check4.Value = 1 Then
     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose")  'let's see what is
    If chk = "Error" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", 1
    ElseIf chk = "1" Then
DeleteValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose"
End If
End If
End Sub
'####################################
Private Sub Check5_Click()

    'Show\Hide Find
    'NoFind
    'Numeric Value 1
If Check5.Value = 1 Then
     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind")  'let's see what is
    If chk = "Error" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 1
    ElseIf chk = "1" Then
DeleteValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind"
End If
End If
End Sub
'####################################
Private Sub Check6_Click()
'Still have to work on this one
    ' Show\Hide Logoff
    'This must be Binary
    
    'NoLogOff
    'Binary Value 01 00 00 00
'    If Check6.Value = 1 Then
'     chk = GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff") 'let's see what is
'     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu") 'let's see what is
'    If chk = "Error" Then
'SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff", "01000000"
'SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", 1
'    ElseIf chk = "01000000" Then
'SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff", 0
'SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", 0
'End If
'End If
MsgBox "Still have to work on this one. if you know how to get the Binary Value can you please E-Mail me (bryce3@bigpond.com)", vbInformation, "Can't use No Log Off Yet"
End Sub
'####################################
Private Sub Check7_Click()
    'Show\Hide C:\Drive
    'NoDrives
    'Numeric Value 4
If Check7.Value = 1 Then
     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives")  'let's see what is
    If chk = "Error" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives", 4
    ElseIf chk = "4" Then
DeleteValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives"
End If
End If
End Sub
'####################################
'----------------------------------

Private Sub cmdRecycle_Add_Click()
'Adds Empty Recycle Bin to Right Click Context Menus
If GetStringValue("HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers\{645FF040-5081-101B-9F08-00AA002F954E}", "") = "Error" Then
CreateKey "HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers\{645FF040-5081-101B-9F08-00AA002F954E}"
End If
End Sub

Private Sub cmdRecycle_Remove_Click()
'Removes Empty Recycle Bin from Right Click Context Menus
DeleteKey "HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers\{645FF040-5081-101B-9F08-00AA002F954E}"
End Sub

Private Sub cmdRecycle_Rename_Click()
'Renames Recycle Bin to what you want
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "", txtRename_Recycle.Text
txtRename_Recycle.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "")
Current_Tab = 1
Form_Load
End Sub

Private Sub cmdRecycle_Default_Click()
'Renames Recycle Bin to Recycle Bin
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "", "Recycle Bin"
Current_Tab = 1
  Form_Load
End Sub

Private Sub Command1_Click()
'Removes custom Open with key
DeleteKey "HKEY_CLASSES_ROOT\*\Shell\O"
End Sub

Private Sub Command10_Click()
'Renames Control Panel
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{21EC2020-3AEA-1069-A2DD-08002B30309D}", "", Text6.Text
Current_Tab = 1
Form_Load
End Sub

Private Sub Command11_Click()
'Changes the Registered Owner
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner", Text4.Text
Current_Tab = 1
Form_Load
End Sub

Private Sub Command12_Click()
'Changes the Registered Organization
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization", Text5.Text
Current_Tab = 1
Form_Load
End Sub

Private Sub Command13_Click()
'Renames Control Panel to Control Panel
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{21EC2020-3AEA-1069-A2DD-08002B30309D}", "", "Control Panel"
Current_Tab = 1
Form_Load
End Sub

Private Sub Command15_Click()
'Add custom Open with to Right Click Context Menus
CreateKey "HKEY_CLASSES_ROOT\*\Shell\" & Text7.Text & "\ Command"
SetStringValue "HKEY_CLASSES_ROOT\*\Shell\" & Text7.Text, "", Text7.Text

SetStringValue "HKEY_CLASSES_ROOT\*\Shell\" & Text7.Text & "\ Command", "", Text8.Text

End Sub

Private Sub Command2_Click()
'Changes or Add a skin to your Internet Explorer & Explorer Browsers
  UpdateBitmapValue
  Command2.Enabled = False
  Command4.Enabled = True
End Sub

Private Sub Command3_Click()
'Add Open with Notepad to Right Click Context Menus
CreateKey "HKEY_CLASSES_ROOT\*\Shell\O\Command"
SetStringValue "HKEY_CLASSES_ROOT\*\Shell\O", "", "&Open with Notepad"
SetStringValue "HKEY_CLASSES_ROOT\*\Shell\O\Command", "", "C:\Windows\Notepad.exe %1"

End Sub

Private Sub Command4_Click()
'Removes skin from your Internet Explorer & Explorer Browsers
  SetStringValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\INTERNET EXPLORER\TOOLBAR", "BACKBITMAP", ""                     'reset value
Current_Tab = 4
  Form_Load
End Sub

Private Sub Command5_Click()
'Rename your Internet Explorer Browser
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main", "Window Title", Text2.Text
Current_Tab = 1
Form_Load
End Sub

Private Sub Command6_Click()
'Default name for your Internet Explorer Browser
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main", "Window Title", ""
SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main", "", ""
Current_Tab = 1
  Form_Load
End Sub

Private Sub Command7_Click()
'Exit
End
End Sub

Private Sub Command8_Click()
'Default name for Outlook Express
SetStringValue "HKEY_CURRENT_USER\Identities\{499B2520-600E-11D4-A890-F0BCB48EAB79}\Software\Microsoft\Outlook Express\5.0", "WindowTitle", ""
SetStringValue "HKEY_CURRENT_USER\Identities\{499B2520-600E-11D4-A890-F0BCB48EAB79}\Software\Microsoft\Outlook Express\5.0", "", ""
Current_Tab = 1
  Form_Load
End Sub

Private Sub Command9_Click()
'Rename Outlook Express
SetStringValue "HKEY_CURRENT_USER\Identities\{499B2520-600E-11D4-A890-F0BCB48EAB79}\Software\Microsoft\Outlook Express\5.0", "WindowTitle", Text3.Text
Current_Tab = 1
Form_Load
End Sub

Private Sub Form_Load()
'####################################
'Current tab to load
If Current_Tab = Empty Then
SSTab1.Tab = 0
Else
SSTab1.Tab = Current_Tab
End If
'####################################
'Get the values
If GetStringValue("HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "Recycle Bin") = "Error" Then
txtRename_Recycle.Text = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "")
End If
'####################################
  fN = GetStringValue("HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\INTERNET EXPLORER\TOOLBAR", "BACKBITMAP")         'let's see what is the default
  Command4.Enabled = Len(fN) > 0
  Text1.Text = fN
'####################################
  
  fN = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\INTERNET EXPLORER\Main", "Window Title") 'let's see what is
    If fN = "Error" Then
    Text2.Text = "Enter new name for your browser"
    Else
  Text2.Text = fN
End If
'####################################
 
 fN = GetStringValue("HKEY_CURRENT_USER\Identities\{499B2520-600E-11D4-A890-F0BCB48EAB79}\Software\Microsoft\Outlook Express\5.0", "WindowTitle") 'let's see what is
    If fN = "Error" Then
    Text3.Text = "Enter new name for Outlook Express"
    Else
  Text3.Text = fN
End If
'####################################

 fN = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner") 'let's see what is
    If fN = "Error" Then
    Text4.Text = "Enter new name for registered owner"
    Else
  Text4.Text = fN
lblUsername.Caption = "Well hello " & fN
End If
'####################################

 fN = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization") 'let's see what is
    If fN = "Error" Then
    Text5.Text = "Enter new name for registered organization"
    Else
  Text5.Text = fN
End If
'####################################

 fN = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "SystemRoot") 'let's see what is
    If fN = "Error" Then
Dir1.Path = Drive1.Drive
    Else
Dir1.Path = fN
End If
'####################################

 fN = GetStringValue("HKEY_CLASSES_ROOT\CLSID\{21EC2020-3AEA-1069-A2DD-08002B30309D}", "") 'let's see what is
    If fN = "Error" Then
    Text6.Text = "Enter new name for Control Panel"
    Else
  Text6.Text = fN
End If
'----------------------------------
'####################################
     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu") 'let's see what is
    If chk = "0" Then
Check1.Value = 0
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", 0
ElseIf chk = "1" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFavoritesMenu", 1
Check1.Value = 1
End If
'####################################

     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu") 'let's see what is
    If chk = "0" Then
Check2.Value = 0
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", 0
ElseIf chk = "1" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRecentDocsMenu", 1
Check2.Value = 1
End If
'####################################

     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun") 'let's see what is
    If chk = "0" Then
Check3.Value = 0
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 0
ElseIf chk = "1" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 1
Check3.Value = 1
End If
'####################################

     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose") 'let's see what is
    If chk = "0" Then
Check4.Value = 0
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", 0
ElseIf chk = "1" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", 1
Check4.Value = 1
End If
'####################################

     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind") 'let's see what is
    If chk = "0" Then
Check5.Value = 0
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 0
ElseIf chk = "1" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 1
Check5.Value = 1
End If
'####################################
'Still have to work on this one
'     chk = GetBinaryValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff") 'let's see what is
'    If chk = "0" Then
'Check6.Value = 0
'SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff", "01000000"
'ElseIf chk = "01000000" Then
'SetBinaryValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff", 1
'Check6.Value = 1
'End If
'####################################
     chk = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives") 'let's see what is
    If chk = "0" Then
Check7.Value = 0
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives", 0
ElseIf chk = "1" Then
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives", 4
Check7.Value = 1
End If
'####################################
'----------------------------------
     
lblUserNameText.Caption = "This program lets you change some of the settings in your Registry." & vbNewLine & _
"It allows you to change." & vbNewLine & vbNewLine & _
"1: The name of your Recycle Bin, Internet Explorer browser, Outlook Express and Control Panel Names" & vbNewLine & _
"2: Your Username and Orgination." & vbNewLine & _
"3: Add Empty Recycle Bin and Open with Notepad to context menus. You can even add custom paths to menu." & vbNewLine & _
"4: Change or add skins to your Internet Explorer and Explorer Browsers." & vbNewLine & vbNewLine & _
"Well I hope you find this usefull in some ways. Not advised if you are using Windows 2000. Reg setting are different" & vbNewLine & _
"Thanks for trying this out. You must Logoff for settings to take effect." & vbNewLine & vbNewLine & _
"BJ.    E-Mail me bryce3@bigpond.com"
End Sub

Private Sub Text1_Change()
'Choose the Skin for your Browser
  On Error Resume Next
  Command2.Enabled = Len(Text1.Text) > 0
  Image1.Picture = LoadPicture()
  Image1.Picture = LoadPicture(Text1.Text)   'a simple preview
End Sub

Private Sub UpdateBitmapValue()
'Choose the Skin for your Browser
  If Len(Text1.Text) = 0 Then Exit Sub
  SetStringValue "HKEY_CURRENT_USER\SOFTWARE\MICROSOFT\INTERNET EXPLORER\TOOLBAR", "BACKBITMAP", Text1.Text             'put the path of our
                                        'picture (*.BMP) to
                                        'this entry please.
End Sub
Private Sub Dir1_Change()
'Choose the Skin for your Browser
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
'Choose the Skin for your Browser
  Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
'Choose the Skin for your Browser
  fN = Dir1.Path + "\" + File1.FileName
  Text1.Text = fN
End Sub



