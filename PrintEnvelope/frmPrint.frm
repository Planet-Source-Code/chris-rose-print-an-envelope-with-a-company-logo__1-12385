VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing..."
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Status 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblAction 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Printing..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngID As Long
Dim strAction As String
Dim strLine1 As String
Dim strLine2 As String
Dim strLine3 As String
Dim strPostcode As String
Dim strCounty As String
Dim strCity As String
Dim strFirstName As String
Dim strLastName As String
Dim strLogo As String


Function PrintEnvelope()

'a must have, err handling
    On Error GoTo PrintEnvelope_Err

'NOTE: ALL VALUES ARE IN CENTIMETERS

'set up frm
    Status.Value = 0
    lblAction.Caption = "...Envelope"

'declare
    Dim strFileName As String
    Dim lngTakeLine As Long
    Dim lngX As Long
    Dim lngY As Long

'set up the line scaling, increase or decreease this to change the distance between lines.
    lngTakeLine = 0.3 'this is taken away... from the distance between the lines
    lngX = 16.3 'the text x pos,  the height from the top of envelope
    lngY = 4 'the text y pos, the width from end. envelpe starts around 7 cm

'set scale mode to centimeters
    Printer.ScaleMode = vbCentimeters
    
'set orientation
    Printer.Orientation = vbPRORLandscape

'paint the logo to the buffer '******** PRINT THE LOGO ***********
    Printer.PaintPicture LoadPicture(Me.Logo), 8, 0.5, 5.5, 4

'print the address
    Printer.CurrentX = lngX
    Printer.CurrentY = lngY 'starting height for text display

'set the font + soze here
    Printer.Font.Size = 12

'loop thr db until record matches client user wants, rest of code a bottom
'    RS.MoveFirst
'    Do While Not RS.EOF
'        If RS.Fields("id") = Me.ID Then
            
'************* BEGIN PRINTING BUFFER *************************************
'here we actualy feed the document to the buffer.
        Printer.Print strFirstName & " " & strLastName
            Printer.CurrentY = Printer.CurrentY + (lngTakeLine * 2)
            Printer.CurrentX = lngX
                For i = 1 To 50 / 6
                    Status.Value = Status.Value + 1
                Next i
                            For ii = 0 To 100
                                DoEvents
                            Next ii
                            
'for design porpouses...
    Printer.Font.Size = 10
    
        Printer.Print strLine1
            Printer.CurrentY = Printer.CurrentY + lngTakeLine
            Printer.CurrentX = lngX
                For i = 1 To 50 / 6
                    Status.Value = Status.Value + 1
                Next i
                            For ii = 0 To 100
                                DoEvents
                            Next ii
        Printer.Print strLine2
            Printer.CurrentY = Printer.CurrentY + lngTakeLine
            Printer.CurrentX = lngX
                For i = 1 To 50 / 6
                    Status.Value = Status.Value + 1
                Next i
                            For ii = 0 To 100
                                DoEvents
                            Next ii
        Printer.Print strLine3
            Printer.CurrentY = Printer.CurrentY + lngTakeLine
            Printer.CurrentX = lngX
                For i = 1 To 50 / 6
                    Status.Value = Status.Value + 1
                Next i
                            For ii = 0 To 100
                                DoEvents
                            Next ii
        Printer.Print strCity
            Printer.CurrentY = Printer.CurrentY + lngTakeLine
            Printer.CurrentX = lngX
                For i = 1 To 50 / 6
                    Status.Value = Status.Value + 1
                Next i
                            For ii = 0 To 100
                                DoEvents
                            Next ii
        Printer.Print strCounty
            Printer.CurrentY = Printer.CurrentY + lngTakeLine
            Printer.CurrentX = lngX
                For i = 1 To 50 / 6
                    Status.Value = Status.Value + 1
                Next i
                            For ii = 0 To 100
                                DoEvents
                            Next ii
        Printer.Print strPostcode
            Printer.CurrentY = Printer.CurrentY + lngTakeLine
            Printer.CurrentX = lngX
                For i = 1 To 50 / 6
                    Status.Value = Status.Value + 1
                Next i
                        For ii = 0 To 100
                            DoEvents
                        Next ii
                        
'flush the buffer and print the document
    Printer.EndDoc

'********************** END PRINTING BUFFER ********************************

'.../ continued if clause
            'dont want it to print lots of times...
'            Exit Do
'        Else
'            RS.MoveNext
'        End If
'    Loop
    
'NOTE: CHANGE THIS
    On Error Resume Next
    
'set up the progress bar 1 more time
    For i = 1 To Status.Max - Status.Value
        Status.Value = Status.Value + 1
            For ii = 0 To 1000
                DoEvents
            Next ii
    Next i
    
    Exit Function
    
PrintEnvelope_Err:
    'warn user of error
        MsgBox "An error has occoured: " & vbCrLf & Err.Description & vbCrLf & "The printing job has been canceled, please try again.", vbOKOnly + vbCritical + vbApplicationModal, "Printer Error"
    'kill what we have done else it will print what we have done when we quit the program.
        Printer.KillDoc

End Function


Sub Init()

'set mouse to hourglass
    Screen.MousePointer = vbHourglass

'show me, modal, and refresh so we can actually see it.
    Me.Show vbApplicationModal
    DoEvents
    DoEvents
    
'check to see the action to do and do it
'    Select Case strAction
'        Case rgEnvelope
            PrintEnvelope
'        Case Else
'            MsgBox "Else"
'    End Select
    
'hide me
    Me.Hide
    
'restore mouse pointer
    Screen.MousePointer = vbDefault

End Sub

Public Property Get ID() As Long
    ID = lngID
End Property

Public Property Let ID(ByVal vNewValue As Long)
    lngID = vNewValue
End Property

Public Property Get Action() As String
    Action = strAction
End Property

Public Property Let Action(ByVal vNewValue As String)
    strAction = vNewValue
End Property

Public Property Get FirstName() As String
    FirstName = strFirstName
End Property

Public Property Let FirstName(ByVal vNewValue As String)
    strFirstName = vNewValue
End Property
Public Property Get LastName() As String
    LastName = strLastName
End Property

Public Property Let LastName(ByVal vNewValue As String)
    strLastName = vNewValue
End Property
Public Property Get Line1() As String
    Line1 = strLine1
End Property

Public Property Let Line1(ByVal vNewValue As String)
    strLine1 = vNewValue
End Property
Public Property Get Line2() As String
    Line2 = strLine2
End Property

Public Property Let Line2(ByVal vNewValue As String)
    strLine2 = vNewValue
End Property
Public Property Get Line3() As String
    Line3 = strLine3
End Property

Public Property Let Line3(ByVal vNewValue As String)
    strLine3 = vNewValue
End Property
Public Property Get Postcode() As String
    Postcode = strPostcode
End Property

Public Property Let Postcode(ByVal vNewValue As String)
    strPostcode = vNewValue
End Property
Public Property Get County() As String
    County = strCounty
End Property

Public Property Let County(ByVal vNewValue As String)
    strCounty = vNewValue
End Property
Public Property Get City() As String
    City = strCity
End Property

Public Property Let City(ByVal vNewValue As String)
    strCity = vNewValue
End Property
Public Property Get Logo() As String
    Logo = strLogo
End Property

Public Property Let Logo(ByVal vNewValue As String)
    strLogo = vNewValue
End Property
