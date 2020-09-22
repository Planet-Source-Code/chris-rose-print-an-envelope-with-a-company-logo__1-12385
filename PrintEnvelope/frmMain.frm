VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print an Envelope"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "Path to logo file..."
      Top             =   3120
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "Postcode"
      Top             =   2640
      Width           =   4455
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "County"
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "City"
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Address Line3"
      Top             =   1560
      Width           =   4455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Address Line2"
      Top             =   1200
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Address Line1"
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Last Name"
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "First Name"
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'print...
    With frmPrint
        .FirstName = Text1
        .LastName = Text2
        .Line1 = Text3
        .Line3 = Text4
        .Line3 = Text5
        .City = Text6
        .County = Text7
        .Postcode = Text8
        .Logo = Text9
    End With

    frmPrint.Init
    
End Sub
