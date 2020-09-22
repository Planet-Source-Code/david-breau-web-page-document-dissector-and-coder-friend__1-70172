VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Green=New Code For Element Select With R Button.  Red=No New Code For Element Select With R Button"
   ClientHeight    =   10740
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   10740
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ckClipCopy 
      Caption         =   "Auto Copy Selected Code Text To Clipboard"
      Height          =   465
      Left            =   2970
      TabIndex        =   13
      Top             =   45
      Width           =   2265
   End
   Begin VB.TextBox tCode 
      Height          =   780
      Index           =   4
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   2655
      Width           =   5325
   End
   Begin VB.TextBox tCode 
      Height          =   2220
      Index           =   3
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   1215
      Width           =   5325
   End
   Begin VB.TextBox tCode 
      Height          =   780
      Index           =   2
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   180
      Width           =   5325
   End
   Begin VB.TextBox tCode 
      Height          =   780
      Index           =   1
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1620
      Width           =   5325
   End
   Begin WebpageCodersFriend.InstantIE InstantIE1 
      Height          =   7440
      Left            =   45
      TabIndex        =   3
      Top             =   3510
      Width           =   10680
      _ExtentX        =   17965
      _ExtentY        =   13123
      StartUrl        =   "http://www.planetsourcecode.com/vb"
      Show_Addressbar =   -1  'True
      Control_Mode    =   1
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   690
   End
   Begin VB.TextBox txtNav 
      Height          =   285
      Left            =   765
      TabIndex        =   1
      Text            =   "http://www.google.com"
      Top             =   45
      Width           =   2040
   End
   Begin VB.TextBox tCode 
      Height          =   780
      Index           =   0
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   585
      Width           =   5325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Code To Get Selected Element By ClassName"
      Height          =   285
      Index           =   4
      Left            =   90
      TabIndex        =   12
      Top             =   2430
      Width           =   3660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Code To Get Selected Element By All"
      Height          =   285
      Index           =   3
      Left            =   5580
      TabIndex        =   10
      Top             =   1035
      Width           =   3390
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Code To Get Selected Element By InnerText"
      Height          =   285
      Index           =   2
      Left            =   5535
      TabIndex        =   8
      Top             =   0
      Width           =   3390
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Code To Get Selected Element By Name"
      Height          =   285
      Index           =   1
      Left            =   135
      TabIndex        =   6
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Code To Get Selected Element By Id"
      Height          =   285
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   405
      Width           =   2895
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu MnuHowTo 
         Caption         =   "How To Use This"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
  

Private Sub cmdGo_Click()
  InstantIE1.IE.Navigate txtNav
End Sub

Private Sub InstantIE1_ALLcode(scode As String)
   If Len(scode) > 0 Then
      tCode(3) = scode
      tCode(3).BackColor = vbGreen
   Else
      tCode(3).BackColor = vbRed
   End If
End Sub

Private Sub InstantIE1_CLASSNAMEcode(scode As String)
   If Len(scode) > 0 Then
      tCode(4) = scode
      tCode(4).BackColor = vbGreen
   Else
      tCode(4).BackColor = vbRed
   End If
End Sub

Private Sub InstantIE1_IDcode(scode As String)
   If Len(scode) > 0 Then
      tCode(0) = scode
      tCode(0).BackColor = vbGreen
   Else
      tCode(0).BackColor = vbRed
   End If
End Sub

Private Sub InstantIE1_INNERTEXTcode(scode As String)
   If Len(scode) > 0 Then
      tCode(2) = scode
      tCode(2).BackColor = vbGreen
   Else
      tCode(2).BackColor = vbRed
   End If
End Sub
 

Private Sub InstantIE1_KeyUp(KeyCode As Integer, ShiftKey As Integer, oElem As MSHTML.HTMLGenericElement)
  Debug.Print oElem.tagName & "  " & KeyCode & "  " & ShiftKey
End Sub

Private Sub InstantIE1_NAMEcode(scode As String)
   If Len(scode) > 0 Then
      tCode(1) = scode
      tCode(1).BackColor = vbGreen
   Else
      tCode(1).BackColor = vbRed
   End If
End Sub
 
 

Private Sub MnuHowTo_Click()
  Fhelp.Show vbModeless, Me
  Fhelp.Text1 = "This utility serves two purposes.  By setting property [" & _
         "Control_Mode] to (user_mode) this control basically serves " & _
         "as a web browser with lots of additionally functionality " & _
         "when it comes programming web pages. There numerous methods " & _
         "and events that allow you to access elements of the webpage " & _
         "such as its anchors and textboxes, returning each of them, " & _
         "seperately, as a collection. This control gives you the " & _
         "ability to easily, with 1 - 3 lines of code, get a handle " & _
         "on almost any element in the webpage based upon its .id, .name " & _
         ", .classname, or innertext.  So if you wanted to created " & _
         "a webbased application that automatically navigates to the " & _
         "yahoo email webpage fills in the username and password and " & _
         "clicks the submit button for you automatically this control" & _
         "gives you the ability to do that in a hassle free way." & vbCrLf & vbCrLf & _
         "By setting property [Control_Mode] to (code_mode) you can " & _
         "easily and quickly get the code necessary to get the handle " & _
         "to each of the webpages elements so you can manipulate them " & _
         "i.e. the above example given with yahoo.  The way you do that " & _
         "is this.  Using the above example, after setting the [Control_Mode " & _
         "] to (code_mode) navigate to the yahoo login page, right click " & _
         "the username textbox and in this forms textboxes the code " & _
         "necessary to access and manipulate that element is immediately " & _
         "present via the username textboxes .id, .name, .classname, " & _
         "and .innertext properties.  Click on the textbox that has " & _
         "the code you want and it automatically copies it to the " & sHelpCont
End Sub
Private Function sHelpCont() As String
  sHelpCont = "clipboard for you." & vbCrLf & vbCrLf & _
         "I not only welcome but strongly suggest ideas for improving " & _
         "this control.  With all the webpage programming I do I finally " & _
         "got tired of all the same code I had to write over and over " & _
         "again to access and manipulate a webpages elements and the " & _
         "creation of this code/control is one of the best, most time " & _
         "saving utilites I have ever created and I hope you feel the " & _
         "same.  And whether you do or dont I guess your voting will " & _
         "tell me how it is and your suggestions and comments will help " & _
         "me make it even better!!"
End Function

Private Sub tCode_GotFocus(Index As Integer)
  If ckClipCopy.Value = vbChecked Then
     tCode(Index).SelStart = 0
     tCode(Index).SelLength = Len(tCode(Index))
     Clipboard.Clear
     Clipboard.SetText tCode(Index)
  End If
End Sub

Private Sub txtNav_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    cmdGo_Click
  End If
End Sub
