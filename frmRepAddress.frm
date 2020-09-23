VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRepAddress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Report"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmRepAddress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5640
   Begin VB.Frame Frame1 
      Caption         =   "Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   4215
      Begin VB.TextBox txtLike 
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Text            =   "A"
         Top             =   435
         Width           =   975
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "All Names"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Names starting with"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOldReport 
      Caption         =   "Old Report"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame fraDestination 
      Caption         =   "Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   4215
      Begin VB.OptionButton optDestination 
         Caption         =   "&Screen"
         Height          =   225
         Index           =   0
         Left            =   870
         TabIndex        =   5
         Top             =   390
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optDestination 
         Caption         =   "P&rinter"
         Height          =   225
         Index           =   1
         Left            =   2790
         TabIndex        =   4
         Top             =   390
         Width           =   825
      End
      Begin VB.Image imgDestination 
         Height          =   480
         Index           =   0
         Left            =   330
         Picture         =   "frmRepAddress.frx":0442
         Top             =   270
         Width           =   480
      End
      Begin VB.Image imgDestination 
         Height          =   480
         Index           =   1
         Left            =   2130
         Picture         =   "frmRepAddress.frx":0D0C
         Top             =   300
         Width           =   480
      End
   End
   Begin VB.TextBox txtReportTitle 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Address List"
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Report Title"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmRepAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConString   As String
Dim rs          As Recordset
Dim SQL         As String


Private Sub cmdOldReport_Click()
On Error GoTo errhandler
Dim Parms(1)    As String
Dim Report      As CrystalReport1
Dim ToPrinter   As Boolean

    With CommonDialog1
        .DefaultExt = "RRS"
        .DialogTitle = "Save Report Data"
        .Filename = "*.RRS"
        .Filter = "Report Record Set (*.RRS)"
        .InitDir = App.path
        .CancelError = True
        .ShowOpen
    End With
    

    ToPrinter = optDestination(1).Value
    Parms(1) = "'" & Trim(txtReportTitle.Text) & "'"
    ' Text Formula Fields must be in single quotes
    
    Set rs = New Recordset
    rs.CursorLocation = adUseClient
    'Load previously saved disconnected recordset
    rs.Open CommonDialog1.Filename
    
    Load frmReportView
    
    Set Report = New CrystalReport1

    
    If Not frmReportView.RepInit(Report, rs, ToPrinter, Parms, False) Then
        MsgBox "There was a problem printing the report.", vbOKOnly + vbExclamation, Me.Caption
        Unload frmReportView
    Else
        If ToPrinter = False Then
            frmReportView.Visible = True
        Else
            Unload frmReportView
        End If
    End If
        
    Set Report = Nothing
    Set ADORst = Nothing
    Unload Me
    
    Exit Sub
errhandler:
    ShowError

End Sub

Private Sub cmdPrint_Click()
Dim Parms(1)    As String
Dim Report      As CrystalReport1
Dim ToPrinter   As Boolean
On Error GoTo errhandler

    ToPrinter = optDestination(1).Value
    Parms(1) = "'" & Trim(txtReportTitle.Text) & "'"
    ' Text Formula Fields must be in single quotes
    
    SQL = ""
    SQL = SQL & "Select " & vbCrLf
    SQL = SQL & "   [Title]," & vbCrLf
    SQL = SQL & "   Init," & vbCrLf
    SQL = SQL & "   [Name]," & vbCrLf
    SQL = SQL & "   Phone_W," & vbCrLf
    SQL = SQL & "   Phone_H," & vbCrLf
    SQL = SQL & "   Phone_F," & vbCrLf
    SQL = SQL & "   Phone_C" & vbCrLf
    SQL = SQL & "from Address" & vbCrLf
    If optSelection(0).Value = True Then
        SQL = SQL & "Where [name] like '" & Trim(txtLike.Text) & "%'" & vbCrLf
    End If
    SQL = SQL & "Order by [Name]" & vbCrLf
    
    Set rs = New Recordset
    rs.CursorLocation = adUseClient
    rs.Open SQL, ConString, adOpenStatic, adLockReadOnly
    
    Load frmReportView
    
    Set Report = New CrystalReport1

    
    If Not frmReportView.RepInit(Report, rs, ToPrinter, Parms, True) Then
        MsgBox "There was a problem printing the report.", vbOKOnly + vbExclamation, Me.Caption
        Unload frmReportView
    Else
        If ToPrinter = False Then
            frmReportView.Visible = True
        Else
            Unload frmReportView
        End If
    End If
        
    Set Report = Nothing
    Set ADORst = Nothing
    Unload Me
    
    Exit Sub

errhandler:
    ShowError
    
End Sub

Private Sub Form_Load()
    
    ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Password=;"
    ConString = ConString & "Data Source=" & App.path & "\Address.mdb;"
    ConString = ConString & "Persist Security Info=True"

End Sub

Private Sub txtLike_Change()
    
    If txtLike = "" Then
        optSelection(1).Value = True
    Else
        optSelection(0).Value = True
    End If

End Sub
