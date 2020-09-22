VERSION 5.00
Begin VB.UserControl InstantIE 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ScaleHeight     =   3960
   ScaleWidth      =   4680
End
Attribute VB_Name = "InstantIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WindowPlacement) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPlacement Lib "user32.dll" (ByVal hwnd As Long, ByRef lpwndpl As WindowPlacement) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const WM_CLOSE As Long = &H10
Private Const WM_DESTROY As Long = &H2
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SW_NORMAL As Long = 1
Private Const SW_MINIMIZE As Long = 6
Private Const SW_MAXIMIZE As Long = 3

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type WindowPlacement
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type
Enum enBy
   byname = 0
   byid = 1
   byclassname = 2
   byinnertext = 3
   All = 4
End Enum
Enum enHtype
   h1 = 1
   h2 = 2
   h3 = 3
   h4 = 4
   h5 = 5
   h6 = 6
End Enum
Enum controlMode
    user_mode = 0
    code_mode = 1
End Enum
Enum enEventToRaise
   event_mousedown = 0
   event_mouseup = 1
   event_keyup = 2
End Enum
Private Type variables
   iehwnd As Long
   old_url As String
   bcancel As Boolean
End Type
Dim v As variables

Public WithEvents IE As InternetExplorer
Attribute IE.VB_VarHelpID = -1
Public WithEvents IEdocument As HTMLDocument
Attribute IEdocument.VB_VarHelpID = -1

'Default Property Values:
Const m_def_Control_Mode = 0
Const m_def_Show_Addressbar = 0
Const m_def_Show_Context_Menu = 0
Const m_def_Show_Toolbar = 0
Const m_def_Show_Statusbar = 0
Const m_def_Show_Menubar = 0
Const m_def_StartUrl = "http://www.ip-mask.com"

'Property Variables:
Dim m_Control_Mode As controlMode
Dim m_Show_Addressbar As Boolean
Dim m_Show_Context_Menu As Boolean
Dim m_Show_Toolbar As Boolean
Dim m_Show_Statusbar As Boolean
Dim m_Show_Menubar As Boolean
Dim m_StartUrl As String


Event IEdocumentReady()
Event MouseUp(X As Integer, Y As Integer, _
             Button As Integer, oElem As HTMLGenericElement)
Event MouseDown(X As Integer, Y As Integer, Button As Integer)
Event KeyUp(KeyCode As Integer, ShiftKey As Integer, oElem As HTMLGenericElement)
Event Error(sProc As String, sDescrip As String)
Event IDcode(scode As String)
Event NAMEcode(scode As String)
Event CLASSNAMEcode(scode As String)
Event INNERTEXTcode(scode As String)
Event ALLcode(scode As String)
 

Sub PrintHtmlToDocument(shtml As String)
Dim d As Object
  
  If IEdocument Is Nothing Then Exit Sub
  Set d = IEdocument
  d.open
  d.write shtml
  d.Close
  Set d = Nothing
End Sub

 
Private Sub IE_DocumentComplete(ByVal pDisp As Object, URL As Variant)
  If v.old_url = IE.LocationURL Then Exit Sub
  v.old_url = IE.LocationURL
  Set IEdocument = IE.Document
  RaiseEvent IEdocumentReady

Exit Sub
err_handler:
  If Err.Number <> 0 Then
    RaiseEvent Error("IE_DocumentComplete", Err.Description)
  End If
End Sub
 
Public Function bSearchDocumentForText(stxt As String) As Boolean
Dim s As String

  If IEdocument Is Nothing Then Exit Function
  s = LCase$(IEdocument.body.innerText)
  If InStr(1, s, stxt) Then bSearchDocumentForText = True
End Function

 
Private Sub IE_DownloadBegin()
  v.bcancel = False
End Sub
 

Private Function IEdocument_oncontextmenu() As Boolean
  IEdocument_oncontextmenu = m_Show_Context_Menu
End Function

Private Sub IEdocument_onkeyup()
  subEventXY event_keyup
End Sub

Private Sub IEdocument_onmousedown()
   subEventXY event_mousedown
End Sub

Private Sub subEventXY(EventToRaise As enEventToRaise, _
                   Optional param1 As Integer)
'OBJECT
Dim oevent As IHTMLEventObj
'INTEGER
Dim X As Integer
Dim Y As Integer


   Set oevent = IEdocument.parentWindow.event
   X = oevent.X
   Y = oevent.Y
   
   If EventToRaise = event_mousedown Then
      RaiseEvent MouseDown(oevent.X, oevent.Y, oevent.Button)
   ElseIf EventToRaise = event_mouseup Then
      RaiseEvent MouseUp(oevent.X, oevent.Y, oevent.Button, IEdocument.activeElement)
   ElseIf EventToRaise = event_keyup Then
      RaiseEvent KeyUp(oevent.KeyCode, oevent.ShiftKey, IEdocument.activeElement)
   End If
End Sub


Private Sub IEdocument_onmouseup()
'OBJECT
Dim ogen As HTMLGenericElement
'STRING
Dim sid As String
Dim sclassname As String
Dim sname As String
Dim sinntxt As String
Dim scodestart As String
Dim scodeend As String
Dim svarsingle As String
Dim svarplural As String
Dim svarfunc As String
Dim svartype As String
 

 
  subEventXY event_mouseup
  If m_Control_Mode = user_mode Then Exit Sub
  sid = LCase$(Trim$(IEdocument.activeElement.id))
  sclassname = LCase$(Trim$(IEdocument.activeElement.className))
  sinntxt = LCase$(Trim$(IEdocument.activeElement.innerText))
 
  
  Select Case LCase$(IEdocument.activeElement.tagName)
     Case Is = "a"
         Dim a As HTMLAnchorElement
         Set a = IEdocument.activeElement
         sname = LCase$(a.Name)
         Set a = Nothing
         
         'for ALL code
         svarsingle = "oA"
         svarplural = "oAs()"
         svarfunc = "oAnchor"
         svartype = "HTMLAnchorElement"
         
         'for all other enum selections
         scodestart = _
           "Dim oA As HTMLAnchorElement" & vbCrLf & _
           "set oA =instantIE1.oAnchor("
            
     Case Is = "font"
         'for ALL code
         svarsingle = "oF"
         svarplural = "oFs()"
         svarfunc = "oFont"
         svartype = "HTMLFontElement"
         
         'for all other enum selections
         scodestart = _
           "Dim oF As HTMLFontElement" & vbCrLf & _
           "Set oF = instantIE1.oFont("
         
     Case Is = "table"
         'for ALL code
         svarsingle = "oT"
         svarplural = "oTs()"
         svarfunc = "oTable"
         svartype = "HTMLTable"
        
         'for all other enum selections
         scodestart = _
           "Dim oT As HTMLTable" & vbCrLf & _
           "Set oT = instantIE1.oTable("
     
     Case Is = "td"
         'for ALL code
         svarsingle = "oTd"
         svarplural = "oTds()"
         svarfunc = "oTableDown"
         svartype = "HTMLTableCell"
          
         'for all other enum selections
         scodestart = _
           "Dim oTd As HTMLTableCell" & vbCrLf & _
           "Set oTd = instantIE1.oTableDown("
           
     Case Is = "tr"
         Dim tr As HTMLTableRow
         Set tr = IEdocument.activeElement
         sname = LCase$(tr.Name)
         Set tr = Nothing
         
         'for ALL code
         svarsingle = "oTr"
         svarplural = "oTrs()"
         svarfunc = "oTableRow"
         svartype = "HTMLTableRow"
         
         'for all other enum selections
         scodestart = _
           "Dim oTr As HTMLTableRow" & vbCrLf & _
           "Set oTr = instantIE1.oTableRow("
           
     Case Is = "input"
         Dim i As HTMLInputElement
         Set i = IEdocument.activeElement
         sname = LCase$(i.Name)
         Set i = Nothing
          
         'for ALL code
         svarsingle = "oInp"
         svarplural = "oInps()"
         svarfunc = "oInput"
         svartype = "HTMLInputElement"
        
         'for all other enum selections
         scodestart = _
           "Dim oInp As HTMLInputElement" & vbCrLf & _
           "Set oInp = instantIE1.oInput("
     
     Case Is = "form"
         Dim f As HTMLFormElement
         Set f = IEdocument.activeElement
         sname = LCase$(f.Name)
         Set f = Nothing
         
         'for ALL code
         svarsingle = "oF"
         svarplural = "oFs()"
         svarfunc = "oForm"
         svartype = "HTMLFormElement"
        
         'for all other enum selections
         scodestart = _
           "Dim oFrm As HTMLFormElement" & vbCrLf & _
           "Set oFrm = instantIE1.oForm("
     
     Case Is = "textarea"
         Dim ta As HTMLTextAreaElement
         Set ta = IEdocument.activeElement
         sname = LCase$(ta.Name)
         Set ta = Nothing
         
         'for ALL code
         svarsingle = "oTA"
         svarplural = "oTAs()"
         svarfunc = "oTextArea"
         svartype = "HTMLTextAreaElement"
        
         'for all other enum selections
         scodestart = _
           "Dim oTxtArea As HTMLTextAreaElement" & vbCrLf & _
           "Set oTxtArea = instantIE1.oTextArea("
           
     Case Is = "select"
         Dim s As HTMLSelectElement
         Set s = IEdocument.activeElement
         sname = LCase$(s.Name)
         Set s = Nothing
         
         'for ALL code
         svarsingle = "oSel"
         svarplural = "oSels()"
         svarfunc = "oSelect"
         svartype = "HTMLSelectElement"
         
         'for all other enum selections
         scodestart = _
           "Dim oSel As HTMLSelectElement" & vbCrLf & _
           "Set oSel = instantIE1.oSelect("
     
     Case Is = "div"
         'for ALL code
         svarsingle = "oDiv"
         svarplural = "oDivs()"
         svarfunc = "oDiv"
         svartype = "HTMLDivElement"
         
         'for all other enum selections
         scodestart = _
           "Dim oDiv As HTMLDivElement" & vbCrLf & _
           "Set oDiv = instantIE1.oDiv("
            
     Case Is = "span"
         'for ALL code
         svarsingle = "oSpan"
         svarplural = "oSpans()"
         svarfunc = "oSpan"
         svartype = "HTMLSpanElement"
        
         'for all other enum selections
         scodestart = _
           "Dim oSpan As HTMLSpanElement" & vbCrLf & _
           "Set oSpan = instantIE1.oSpan("
           
           
     Case Is = "img"
         Dim img As HTMLImg
         Set img = IEdocument.activeElement
         sname = LCase$(img.Name)
         Set img = Nothing
         
         'for ALL code
         svarsingle = "oImg"
         svarplural = "oImgs()"
         svarfunc = "oImg"
         svartype = "HTMLImg"
         
         'for all other enum selections
         scodestart = _
           "Dim oImg As HTMLImg" & vbCrLf & _
           "Set oImg = instantIE1.oImg("
     
     Case Else
         Exit Sub
   End Select
   
   
   If Len(sid) > 0 Then
      scodeend = "byid," & Chr(34) & sid & Chr(34) & ")(0)"
      RaiseEvent IDcode(scodestart & scodeend)
   Else
      RaiseEvent IDcode("")
   End If
   
   If Len(sclassname) > 0 Then
      scodeend = "byclassname," & Chr(34) & sclassname & Chr(34) & ")(0)"
      RaiseEvent CLASSNAMEcode(scodestart & scodeend)
   Else
      RaiseEvent CLASSNAMEcode("")
   End If
   
   If Len(sname) > 0 Then
      scodeend = "byname," & Chr(34) & sname & Chr(34) & ")(0)"
      RaiseEvent NAMEcode(scodestart & scodeend)
   Else
      RaiseEvent NAMEcode("")
   End If
   
   If Len(sinntxt) > 0 Then
      scodeend = "byinnertext, " & Chr(34) & sinntxt & Chr(34) & ")(0)"
      RaiseEvent INNERTEXTcode(scodestart & scodeend)
   Else
      RaiseEvent INNERTEXTcode("")
   End If
   
   
   Dim s0 As String
      
       s0 = "Dim " & svarsingle & " As " & svartype & vbCrLf & _
            "Dim " & svarplural & " As " & svartype & vbCrLf & _
            "Dim i As Integer" & vbCrLf & _
            "Dim icnt As Integer" & vbCrLf & vbCrLf & _
            svarplural & " = " & svarfunc & "(ALL, , False)" & vbCrLf & _
            "icnt = UBound(" & Split(svarplural, "(")(0) & ")" & vbCrLf & vbCrLf & _
            "For i = 0 To icnt" & vbCrLf & _
            "   'code to process items in array [" & svarplural & "]" & vbCrLf & _
            "Next i"
       RaiseEvent ALLcode(s0)
End Sub

Private Sub UserControl_Resize()
Dim iwid As Long
Dim ihei As Long
  
  iwid = (Width / Screen.TwipsPerPixelX)
  ihei = (Height - (Height - ScaleHeight)) / Screen.TwipsPerPixelY
  MoveWindow v.iehwnd, 0, 0, iwid, ihei, True
End Sub

Private Sub UserControl_Show()
On Error Resume Next

  If Ambient.UserMode Then
     subMakeIe
     RemoveAddTitlebar v.iehwnd
     SpecifyWindowPlacement v.iehwnd
     IE.Visible = True
     'for some reason IE created in this manner
     'dont become visible. so by creating a pause
     'and then again setting the visible property
     'usually takes care of that
     DoEvents: DoEvents: DoEvents: DoEvents
     IE.Visible = True
  End If
End Sub

Private Sub UserControl_Terminate()
   subKillIe
   Set IEdocument = Nothing
End Sub

 














Private Sub subMakeIe()
  On Error GoTo errhander:
  'create new IE and set is nav bars
  Set IE = New InternetExplorer
  
  If Len(Trim$(m_StartUrl)) > 0 Then
     IE.Navigate m_StartUrl
  Else
     IE.Navigate "about:blank"
  End If
  
  v.iehwnd = IE.Parent.hwnd
  IE.ToolBar = m_Show_Toolbar
  IE.MenuBar = m_Show_Menubar
  IE.AddressBar = m_Show_Addressbar
  IE.StatusBar = m_Show_Statusbar
  '
  'place IE in a container if desired
  SetParent v.iehwnd, UserControl.hwnd
  
Exit Sub
errhander:
  If Err.Number <> 0 Then
    Select Case Err.Number
      Case Is = 462 'ie has been closed no longer available
         Exit Sub
      Case Else
         RaiseEvent Error("subMakeIe", Err.Description)
    End Select
  End If
End Sub

Private Sub subKillIe()
On Error GoTo errhander:
   
   If IE Is Nothing Then Exit Sub
   IE.Stop
   IE.Visible = False
   SetParent v.iehwnd, GetDesktopWindow
   PostMessage v.iehwnd, WM_CLOSE, 0&, 0&
   Set IE = Nothing
   
Exit Sub
errhander:
  If Err.Number <> 0 Then
    Select Case Err.Number
      Case Is = 462 'ie has been closed no longer available
         Exit Sub
      Case Else
         RaiseEvent Error("subKillIe", Err.Description)
    End Select
  End If
End Sub

Private Sub RemoveAddTitlebar( _
                     lhwnd As Long, _
                     Optional bRemove As Boolean = True)
Dim lStyle As Long
   
   ' Retrieve current style bits.
   lStyle = GetWindowLong(lhwnd, GWL_STYLE)
   
   ' Set requested bit On or Off and Redraw.
   If bRemove Then
      lStyle = lStyle And Not WS_CAPTION
   Else
      lStyle = lStyle Or WS_CAPTION
   End If
   
   Call SetWindowLong(lhwnd, GWL_STYLE, lStyle)
   Call pRedraw(lhwnd)
End Sub

Private Sub pRedraw(mhwnd As Long)
On Error Resume Next

   ' Redraw window with new style.
   Const swpFlags As Long = _
      SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
   Call SetWindowPos(mhwnd, 0, 0, 0, 0, 0, swpFlags)
End Sub



Private Sub SpecifyWindowPlacement(lhwnd As Long)
Dim WinEst As WindowPlacement
Dim rtn As Long
    
    WinEst.Length = Len(WinEst)
    'get the current window placement
    rtn = GetWindowPlacement(lhwnd, WinEst)
    WinEst.showCmd = SW_MAXIMIZE
    'set the new window placement (minimized)
    SetWindowPlacement lhwnd, WinEst
End Sub
 
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_StartUrl = m_def_StartUrl
    m_Show_Context_Menu = m_def_Show_Context_Menu
    m_Show_Toolbar = m_def_Show_Toolbar
    m_Show_Statusbar = m_def_Show_Statusbar
    m_Show_Menubar = m_def_Show_Menubar
    m_Show_Addressbar = m_def_Show_Addressbar
    m_Control_Mode = m_def_Control_Mode
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_StartUrl = PropBag.ReadProperty("StartUrl", m_def_StartUrl)
    m_Show_Context_Menu = PropBag.ReadProperty("Show_Context_Menu", m_def_Show_Context_Menu)
    m_Show_Toolbar = PropBag.ReadProperty("Show_Toolbar", m_def_Show_Toolbar)
    m_Show_Statusbar = PropBag.ReadProperty("Show_Statusbar", m_def_Show_Statusbar)
    m_Show_Menubar = PropBag.ReadProperty("Show_Menubar", m_def_Show_Menubar)
    m_Show_Addressbar = PropBag.ReadProperty("Show_Addressbar", m_def_Show_Addressbar)
    m_Control_Mode = PropBag.ReadProperty("Control_Mode", m_def_Control_Mode)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("StartUrl", m_StartUrl, m_def_StartUrl)
    Call PropBag.WriteProperty("Show_Context_Menu", m_Show_Context_Menu, m_def_Show_Context_Menu)
    Call PropBag.WriteProperty("Show_Toolbar", m_Show_Toolbar, m_def_Show_Toolbar)
    Call PropBag.WriteProperty("Show_Statusbar", m_Show_Statusbar, m_def_Show_Statusbar)
    Call PropBag.WriteProperty("Show_Menubar", m_Show_Menubar, m_def_Show_Menubar)
    Call PropBag.WriteProperty("Show_Addressbar", m_Show_Addressbar, m_def_Show_Addressbar)
    Call PropBag.WriteProperty("Control_Mode", m_Control_Mode, m_def_Control_Mode)
End Sub

'STARTURL
Public Property Get StartUrl() As String
    StartUrl = m_StartUrl
End Property

Public Property Let StartUrl(ByVal New_StartUrl As String)
    m_StartUrl = New_StartUrl
    PropertyChanged "StartUrl"
End Property

'SHOHW_CONTEXTMENU
Public Property Get Show_Context_Menu() As Boolean
Attribute Show_Context_Menu.VB_Description = "whether or not to display IE right click context menu at runtime"
    Show_Context_Menu = m_Show_Context_Menu
End Property

Public Property Let Show_Context_Menu(ByVal New_Show_Context_Menu As Boolean)
    m_Show_Context_Menu = New_Show_Context_Menu
    PropertyChanged "Show_Context_Menu"
End Property

'SHOW_TOOLBAR
Public Property Get Show_Toolbar() As Boolean
Attribute Show_Toolbar.VB_Description = "whether or not to display IE toolbar"
    Show_Toolbar = m_Show_Toolbar
End Property

Public Property Let Show_Toolbar(ByVal New_Show_Toolbar As Boolean)
    m_Show_Toolbar = New_Show_Toolbar
    PropertyChanged "Show_Toolbar"
End Property

'SHOW_STATUSBAR
Public Property Get Show_Statusbar() As Boolean
Attribute Show_Statusbar.VB_Description = "whether or not to display IE status bar"
    Show_Statusbar = m_Show_Statusbar
End Property

Public Property Let Show_Statusbar(ByVal New_Show_Statusbar As Boolean)
    m_Show_Statusbar = New_Show_Statusbar
    PropertyChanged "Show_Statusbar"
End Property

'SHOW_MENUBAR
Public Property Get Show_Menubar() As Boolean
Attribute Show_Menubar.VB_Description = "whether or not to display IE menu bar"
    Show_Menubar = m_Show_Menubar
End Property

Public Property Let Show_Menubar(ByVal New_Show_Menubar As Boolean)
    m_Show_Menubar = New_Show_Menubar
    PropertyChanged "Show_Menubar"
End Property

'SHOW_ADDRESSBAR
Public Property Get Show_Addressbar() As Boolean
Attribute Show_Addressbar.VB_Description = "whether or not to display IE address bar"
    Show_Addressbar = m_Show_Addressbar
End Property

Public Property Let Show_Addressbar(ByVal New_Show_Addressbar As Boolean)
    m_Show_Addressbar = New_Show_Addressbar
    PropertyChanged "Show_Addressbar"
End Property

'CONTROL MODE
Public Property Get Control_Mode() As controlMode
Attribute Control_Mode.VB_Description = "When coding the web page document and you want pre prepared code provided for you in events [IDcode, NAMEcode, INNERTEXTcode, ALLcode] then select code_mode. If using this control as a webbrowser for the software user then select mode_user (default)"
    Control_Mode = m_Control_Mode
End Property

Public Property Let Control_Mode(ByVal New_Control_Mode As controlMode)
    m_Control_Mode = New_Control_Mode
    PropertyChanged "Control_Mode"
End Property



 




































'COLLECTION TO ALL ANCHORS IN THE CURRENT WEBPAGE/DOCUMENT <a>
Public Function oAnchor( _
               get_by As enBy, _
               Optional sIdOrNameOrClassnameOrInnerText As String, _
               Optional bReturnFirstOnly As Boolean = True) _
               As HTMLAnchorElement()
               
  Dim a As HTMLAnchorElement, col As IHTMLElementCollection
  Dim s As String, icnt As Integer, collAtemp() As HTMLAnchorElement
  
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  Set col = IEdocument.getElementsByTagName("a")
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
   
  For Each a In col
    DoEvents
    If v.bcancel Then Exit For
    
    If get_by = byclassname Then
      If InStr(1, LCase$(a.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(a.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(a.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(a.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next a
  
save_this:
    oAnchor = collAtemp
    Erase collAtemp
    Set a = Nothing
    Set col = Nothing

Exit Function
process_this:
        ReDim Preserve collAtemp(icnt)
        Set collAtemp(icnt) = a
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
err_handler:
 If Err <> 0 Then
    RaiseEvent Error("oAnchor", Err.Description)
 End If
End Function

'COLLECTION TO ALL INPUT ELEMENTS (textboxes, submit buttons etc)
'IN THE CURRENT WEBPAGE/DOCUMENT <input>
Function oInput(get_by As enBy, _
                Optional sIdOrNameOrClassnameOrInnerText As String, _
                Optional bReturnFirstOnly As Boolean = True) _
                As HTMLInputElement()
                
Dim i As HTMLInputElement, col As IHTMLElementCollection
Dim s As String, icnt As Integer, colInpTemp() As HTMLInputElement
  
  On Error GoTo err_handler:
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  Set col = IEdocument.getElementsByTagName("input")
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
  
  For Each i In col
    DoEvents
    If v.bcancel Then Exit For
    
    If get_by = byclassname Then
      If InStr(1, LCase$(i.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(i.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(i.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(i.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next i
 
 
save_this:
  oInput = colInpTemp
  Erase colInpTemp
  Set i = Nothing
  Set col = Nothing
 

Exit Function
process_this:
        ReDim Preserve colInpTemp(icnt)
        Set colInpTemp(icnt) = i
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
        
err_handler:
  With Err
    If .Number <> 0 Then
        RaiseEvent Error("oInput", Err.Description)
    End If
  End With
End Function


'GET COLLECTION OF SELECT ELEMENTS (combobox in html terms) <select>
Function oSelect(get_by As enBy, _
                 Optional sIdOrNameOrClassnameOrInnerText As String, _
                 Optional bReturnFirstOnly As Boolean = True) _
                 As HTMLSelectElement()
                        
Dim oS As HTMLSelectElement, col As IHTMLElementCollection
Dim s As String, icnt As Integer, colSelTemp() As HTMLSelectElement
  
  On Error GoTo err_handler:
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  Set col = IEdocument.getElementsByTagName("select")
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
   
  
  For Each oS In col
    DoEvents
    If v.bcancel Then Exit For
    
    If get_by = byclassname Then
      If InStr(1, LCase$(oS.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(oS.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(oS.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(oS.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next oS
  
  
save_this:
    oSelect = colSelTemp
    Erase colSelTemp
    Set oS = Nothing
    Set col = Nothing

Exit Function
process_this:
        ReDim Preserve colSelTemp(icnt)
        Set colSelTemp(icnt) = oS
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
        
err_handler:
  With Err
    If .Number <> 0 Then
       RaiseEvent Error("oSelect", Err.Description)
    End If
  End With
End Function

'COLLECTION OF TABLE ROWS <tr>
Function oTableRow(get_by As enBy, _
                  Optional sIdOrNameOrClassnameOrInnerText As String, _
                  Optional bReturnFirstOnly As Boolean = True) _
                  As HTMLTableRow()

Dim oTr As HTMLTableRow, col As IHTMLElementCollection
Dim s As String, icnt As Integer, colTrTemp() As HTMLTableRow
  
  On Error GoTo err_handler:
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  Set col = IEdocument.getElementsByTagName("tr")
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
   
  
  For Each oTr In col
    DoEvents
    If v.bcancel Then Exit For
    
    If get_by = byclassname Then
      If InStr(1, LCase$(oTr.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(oTr.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(oTr.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(oTr.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next oTr
  
  
save_this:
    oTableRow = colTrTemp
    Erase colTrTemp
    Set oTr = Nothing
    Set col = Nothing

Exit Function
process_this:
        ReDim Preserve colTrTemp(icnt)
        Set colTrTemp(icnt) = oTr
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
        
err_handler:
  If Err.Number <> 0 Then
     RaiseEvent Error("oTableRow", Err.Description)
  End If
End Function

'COLLECTION OF FORMS <form>
Function oForms(get_by As enBy, _
                  Optional sIdOrNameOrClassnameOrInnerText As String, _
                  Optional bReturnFirstOnly As Boolean = True) _
                  As HTMLFormElement()

Dim oF As HTMLFormElement, col As IHTMLElementCollection
Dim s As String, icnt As Integer, colFormTemp() As HTMLFormElement
  
  On Error GoTo err_handler:
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  Set col = IEdocument.getElementsByTagName("form")
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
   
  
  For Each oF In col
    DoEvents
    If v.bcancel Then Exit For
    
    If get_by = byclassname Then
      If InStr(1, LCase$(oF.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(oF.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(oF.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(oF.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next oF
  
  
save_this:
    oForms = colFormTemp
    Erase colFormTemp
    Set oF = Nothing
    Set col = Nothing

Exit Function
process_this:
        ReDim Preserve colFormTemp(icnt)
        Set colFormTemp(icnt) = oF
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
        
err_handler:
  If Err.Number <> 0 Then
     RaiseEvent Error("oForms", Err.Description)
  End If
End Function

'COLLECTION OF FORMS <form>
Function oTextArea(get_by As enBy, _
                  Optional sIdOrNameOrClassnameOrInnerText As String, _
                  Optional bReturnFirstOnly As Boolean = True) _
                  As HTMLTextAreaElement()

Dim oTa As HTMLTextAreaElement, col As IHTMLElementCollection
Dim s As String, icnt As Integer, colTaTemp() As HTMLTextAreaElement
  
  On Error GoTo err_handler:
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  Set col = IEdocument.getElementsByTagName("textarea")
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
   
  
  For Each oTa In col
    DoEvents
    If v.bcancel Then Exit For
    
    If get_by = byclassname Then
      If InStr(1, LCase$(oTa.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(oTa.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(oTa.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(oTa.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next oTa
  
  
save_this:
    oTextArea = colTaTemp
    Erase colTaTemp
    Set oTa = Nothing
    Set col = Nothing

Exit Function
process_this:
        ReDim Preserve colTaTemp(icnt)
        Set colTaTemp(icnt) = oTa
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
        
err_handler:
  If Err.Number <> 0 Then
     RaiseEvent Error("oForms", Err.Description)
  End If
End Function

'COLLECTION OF TABLE DOWNS <td>
Function oTableDown(get_by As enBy, _
                  Optional sIdOrNameOrClassnameOrInnerText As String, _
                  Optional bReturnFirstOnly As Boolean = True) _
                  As HTMLTableCell()
 
Dim otd As HTMLTableRow, col As IHTMLElementCollection
Dim s As String, icnt As Integer, colTdTemp() As HTMLTableCell
  
  On Error GoTo err_handler:
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  Set col = IEdocument.getElementsByTagName("td")
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
   
  
  For Each otd In col
    DoEvents
    If v.bcancel Then Exit For
   
    If get_by = byclassname Then
      If InStr(1, LCase$(otd.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(otd.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(otd.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(otd.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next otd
  
save_this:
    oTableDown = colTdTemp
    Erase colTdTemp
    Set otd = Nothing
    Set col = Nothing

Exit Function
process_this:
        ReDim Preserve colTdTemp(icnt)
        Set colTdTemp(icnt) = otd
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
        
err_handler:
  If Err.Number <> 0 Then
     RaiseEvent Error("oTableDown", Err.Description)
  End If
End Function

'COLLECTION OF FONT ELEMENTS <font>
Function oFont(get_by As enBy, _
                  Optional sIdOrNameOrClassnameOrInnerText As String, _
                  Optional bReturnFirstOnly As Boolean = True) _
                  As HTMLFontElement()


Dim oF As HTMLFontElement, col As IHTMLElementCollection
Dim s As String, icnt As Integer, colFTemp() As HTMLFontElement
  
  On Error GoTo err_handler:
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  Set col = IEdocument.getElementsByTagName("font")
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
   
  
  For Each oF In col
    DoEvents
    If v.bcancel Then Exit For
   
    If get_by = byclassname Then
      If InStr(1, LCase$(oF.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(oF.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(oF.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(oF.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next oF
  
save_this:
    oFont = colFTemp
    Erase colFTemp
    Set oF = Nothing
    Set col = Nothing

Exit Function
process_this:
        ReDim Preserve colFTemp(icnt)
        Set colFTemp(icnt) = oF
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
        
err_handler:
   If Err.Number <> 0 Then
     RaiseEvent Error("oFont", Err.Description)
  End If
End Function

'COLLECTION OF HEADER ELEMENTS <h1>, <h2> etc
Function oH(Htype As enHtype, _
             get_by As enBy, _
             Optional sIdOrNameOrClassnameOrInnerText As String, _
             Optional bReturnFirstOnly As Boolean = True) _
             As HTMLHeaderElement()

Dim oHval As HTMLHeaderElement, col As IHTMLElementCollection
Dim s As String, icnt As Integer, colHTemp() As HTMLHeaderElement
Dim sHtype As String
  
  On Error GoTo err_handler:
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  sHtype = "H" & CStr(Htype)
  Set col = IEdocument.getElementsByTagName(sHtype)
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
 
  
  For Each oHval In col
    DoEvents
    If v.bcancel Then Exit For
    
    If get_by = byclassname Then
      If InStr(1, LCase$(oHval.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(oHval.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(oHval.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(oHval.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next oHval
  
save_this:
    oH = colHTemp
    Erase colHTemp
    Set oHval = Nothing
    Set col = Nothing

Exit Function
process_this:
        ReDim Preserve colHTemp(icnt)
        Set colHTemp(icnt) = oHval
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
        
err_handler:
   If Err.Number <> 0 Then
     RaiseEvent Error("oH", Err.Description)
  End If
End Function

'COLLECTION OF WEB PAGES IMAGES <img>
Function oImg(get_by As enBy, _
                  Optional sIdOrNameOrClassnameOrInnerText As String, _
                  Optional bReturnFirstOnly As Boolean = True) _
                  As HTMLImg()

Dim oImage As HTMLImg, col As IHTMLElementCollection
Dim s As String, icnt As Integer, colImgTemp() As HTMLImg
  
  On Error GoTo err_handler:
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  Set col = IEdocument.getElementsByTagName("img")
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
   
  
  For Each oImage In col
    DoEvents
    If v.bcancel Then Exit For
    
    If get_by = byclassname Then
      If InStr(1, LCase$(oImage.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(oImage.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(oImage.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(oImage.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next oImage
  
  
save_this:
    oImg = colImgTemp
    Erase colImgTemp
    Set oImage = Nothing
    Set col = Nothing

Exit Function
process_this:
        ReDim Preserve colImgTemp(icnt)
        Set colImgTemp(icnt) = oImage
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
        
err_handler:
   If Err.Number <> 0 Then
     RaiseEvent Error("oImg", Err.Description)
  End If
End Function

'COLLECTION OF WEB PAGE SPAN ELEMENTS <span>
Function oSpan(get_by As enBy, _
                  Optional sIdOrNameOrClassnameOrInnerText As String, _
                  Optional bReturnFirstOnly As Boolean = True) _
                  As HTMLSpanElement()

Dim ospn As HTMLSpanElement, col As IHTMLElementCollection
Dim s As String, icnt As Integer, colSpanTemp() As HTMLSpanElement
  
  On Error GoTo err_handler:
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  Set col = IEdocument.getElementsByTagName("span")
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
   
  
  For Each ospn In col
    DoEvents
    If v.bcancel Then Exit For
    
    If get_by = byclassname Then
      If InStr(1, LCase$(ospn.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(ospn.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(ospn.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(ospn.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next ospn
  
  
save_this:
    oSpan = colSpanTemp
    Erase colSpanTemp
    Set ospn = Nothing
    Set col = Nothing

Exit Function
process_this:
        ReDim Preserve colSpanTemp(icnt)
        Set colSpanTemp(icnt) = ospn
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
        
err_handler:
   If Err.Number <> 0 Then
     RaiseEvent Error("oSpan", Err.Description)
  End If
End Function

'COLLECTION OF DIV ELEMENTS IN THE WEB PAGE <div>
Function oDiv(get_by As enBy, _
                  Optional sIdOrNameOrClassnameOrInnerText As String, _
                  Optional bReturnFirstOnly As Boolean = True) _
                  As HTMLDivElement()

Dim od    As HTMLDivElement, col As IHTMLElementCollection
Dim s As String, icnt As Integer, colDivTemp() As HTMLDivElement
  
  On Error GoTo err_handler:
  If IEdocument Is Nothing Then Exit Function
  v.bcancel = False
  Set col = IEdocument.getElementsByTagName("div")
  s = LCase$(sIdOrNameOrClassnameOrInnerText)
   
  
  For Each od In col
    DoEvents
    If v.bcancel Then Exit For
    
    If get_by = byclassname Then
      If InStr(1, LCase$(od.className), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byid Then
      If InStr(1, LCase$(od.id), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = byname Then
      If InStr(1, LCase$(od.Name), s) Then
        GoTo process_this:
      End If
    
    ElseIf get_by = byinnertext Then
      If InStr(1, LCase$(od.innerText), s) Then
        GoTo process_this:
      End If
      
    ElseIf get_by = All Then
       GoTo process_this:
    End If
re_start:
  Next od
  
save_this:
    oDiv = colDivTemp
    Erase colDivTemp
    Set od = Nothing
    Set col = Nothing

Exit Function
process_this:
        ReDim Preserve colDivTemp(icnt)
        Set colDivTemp(icnt) = od
        icnt = (icnt + 1)
        If bReturnFirstOnly Then GoTo save_this:
        GoTo re_start:
        
err_handler:
   If Err.Number <> 0 Then
     RaiseEvent Error("oDiv", Err.Description)
  End If
End Function
 
'RETURNS THE OPTION ELEMENTS WITHING A SELECT ELEMENT
'WHICH IS BASICALLY THE SAME AS THE ITEMS IN A COMBOBOX
Function oOptions(oSel As HTMLSelectElement) As HTMLOptionElement()
Dim i As Integer, icnt As Integer
Dim colOptTemp() As HTMLOptionElement
 

  If oSel Is Nothing Then Exit Function
  v.bcancel = False
  
  With oSel.getElementsByTagName("option")
     icnt = .Length - 1
     If icnt = 0 Then Exit Function
     ReDim Preserve colOptTemp(icnt)
     
     For i = 0 To icnt
       If v.bcancel Then Exit For
       Set colOptTemp(i) = .Item(i)
     Next i
  End With
 
  oOptions = colOptTemp
  Erase colOptTemp
  
Exit Function
err_handler:
   If Err.Number <> 0 Then
     RaiseEvent Error("oOptions", Err.Description)
  End If
End Function
 

