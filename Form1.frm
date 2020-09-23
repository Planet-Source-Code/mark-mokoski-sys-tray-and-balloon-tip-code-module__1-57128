VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3660
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5190
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":1272
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Code by"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label websiteLabel 
      Alignment       =   2  'Center
      Caption         =   "www.cmtelephone.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   360
      MouseIcon       =   "Form1.frx":14AB
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Click to goto my web site"
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Label emailLabel 
      Alignment       =   2  'Center
      Caption         =   "markm@cmtelephone.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      MouseIcon       =   "Form1.frx":17B5
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Click to send mail to Mark Mokoski"
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mark Mokoski"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Form to Screen"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMclose 
         Caption         =   "Close Menu"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "End Program"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Addtional Information"
      Begin VB.Menu mnuFormCode 
         Caption         =   "Form Code"
      End
      Begin VB.Menu mnuMSDN1 
         Caption         =   "MSDN Ref 1"
      End
      Begin VB.Menu mnuMSDN2 
         Caption         =   "MSDN Ref 2"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

    '********************************************************************
    '
    'Systray, Balloon Tool Tip add-in code to the form
    '
    'Mark Mokoski
    'markm@cmtelephone.com
    'www.cmtelephone.com
    '
    '6-NOV-2004
    '
    'See Systray Form Code.txt in the ZIP file for form add-in's to make it all work
    '
    'Also see Microsoft Knowledge base http://support.microsoft.com/default.aspx?scid=kb;en-us;149276
    'for more information.
    '
    'This code is based on the Microsoft Knowledge Base code.
    '********************************************************************
    

Private Sub Form_Load()

    '********************************************************************
    'If you want the form to be in the tray on startup add this

    Call SystrayOn(Me, "The Form is visible on the screen")
    Call PopupBalloon(Me, "Put your Message Here !", "Balloon Tool Tip")

End Sub

Private Sub Form_Resize()

    '********************************************************************
    'Add this to resize event to hide in tray on minimize

        If Me.WindowState = vbMinimized Then
            Call SystrayOn(Me, "Double Click to Restore Me back to the screen")
            Call ChangeSystrayToolTip(Me, "Double Click to Restore Me back to the screen")
            Call PopupBalloon(Me, "App is now hidden in the Systray !" + vbCrLf + "Double click Icon to restore", "Balloon Tool Tip")
            Me.Hide
        End If

End Sub

Private Sub Form_Terminate()

    '********************************************************************
    'If you don't remove icon from tray on double click show, add this
    'good idea

    Call SystrayOff(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SystrayOff(Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '*********************************************************************
    'Add this event code to repond to mouse over and clicks on icon in the tray

    Static lngMsg            As Long
    Dim blnflag              As Boolean

    lngMsg = X / Screen.TwipsPerPixelX

        If blnflag = False Then

            blnflag = True
        
                Select Case lngMsg
                    Case WM_RBUTTONCLK      'to popup menu on right-click
                        Call SetForegroundWindow(Me.hWnd)
                        Call RemoveBalloon(Me)
                        'Reference the menu object of the form below for popup
                        PopupMenu Me.mnuFile

                    Case WM_LBUTTONDBLCLK   'SHow form on left-dblclick
                        'Use line below if you want to remove tray icon on dbclick show form.
                        'If not, be sure to put Systrayoff in form unload and terminate events.
                        'Call SystrayOff(Me)
                        Call ChangeSystrayToolTip(Me, "The Form is visible on the screen")
                        Call SetForegroundWindow(Me.hWnd)
                        Call RemoveBalloon(Me)
                        Me.WindowState = vbNormal
                        Me.Show
                        Me.SetFocus
            
                End Select
        
            blnflag = False
        
        End If
    
End Sub

Private Sub emailLabel_Click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "mailto:name@domain.com?body=hello%0a%0world", vbNullString, vbNullString, vbNormalFocus
    ShellExecute hWnd, vbNullString, "mailto:markm@cmtelephone.com?Subject=Questions or Comments on Systray Code Module. %09 ", vbNullString, vbNullString, vbNormalFocus
  
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)

End Sub

Private Sub mnuEnd_Click()

    Unload Me

End Sub

Private Sub mnuFormCode_Click()

    'Addtional code need in form events for Systray to work OK
    ShellExecute hWnd, vbNullString, App.Path & "\systray form code.txt", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub mnuMSDN1_Click()

    'How To Use Icons with the Windows 95.htm
    ShellExecute hWnd, vbNullString, App.Path & "\html\How To Use Icons with the Windows 95.htm", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub mnuMSDN2_Click()

    'How To Manipulate Icons in the System Tray with Visual Basic.htm
    ShellExecute hWnd, vbNullString, App.Path & "\html\How To Manipulate Icons in the System Tray with Visual Basic.htm", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub mnuRestore_Click()

    Me.WindowState = vbNormal
    Me.Show

End Sub

Private Sub websiteLabel_Click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "http://www.domain.com", vbNullString, vbNullString, vbNormalFocus
    ShellExecute hWnd, vbNullString, "http://www.rjillc.com", vbNullString, vbNullString, vbNormalFocus
  
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)

End Sub
