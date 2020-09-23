VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dBASE Viewer"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BrowseButton 
      Caption         =   "Browse"
      Height          =   285
      Left            =   2580
      TabIndex        =   1
      Top             =   2220
      Width           =   1965
   End
   Begin VB.TextBox Descr 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   2220
      Width           =   2505
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3990
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2145
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   3784
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483627
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function GetConnectionString(ByVal mFolder As String) As String
    GetConnectionString = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=" & mFolder
End Function

Private Sub BrowseButton_Click()
On Error GoTo FileOpen_ClickError
    Dim db As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim mItem As MSComctlLib.ListItem
    Dim mIndex As Long
    Dim mValue As String
    Dim mFolder As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "dBASE Files (*.dbf)| *.dbf"
    CommonDialog1.InitDir = App.Path
    Call CommonDialog1.ShowOpen
    Descr.Text = CommonDialog1.FileTitle
    
    If (Len(Trim(CommonDialog1.FileName)) > 0) Then
        Set db = New ADODB.Connection
        mFolder = Replace(CommonDialog1.FileName, CommonDialog1.FileTitle, "", 1)
        Call db.Open(GetConnectionString(mFolder))
        Set rs = New ADODB.Recordset
        Call rs.Open("SELECT * FROM " & Descr.Text, db, adOpenStatic, adLockReadOnly)
        Call ResetList
        Call DefineList(rs)
        If (rs.RecordCount > 0) Then
            rs.MoveFirst
            While Not (rs.EOF)
                If (IsNull(rs.Fields(0))) Then
                    mValue = ""
                Else
                    mValue = CStr(rs.Fields(0))
                End If
                Set mItem = ListView1.ListItems.Add(1, , mValue)
                For mIndex = 1 To rs.Fields.Count - 1
                    If (IsNull(rs.Fields(mIndex))) Then
                        mValue = ""
                    Else
                        mValue = CStr(rs.Fields(mIndex))
                    End If
                    mItem.SubItems(mIndex) = mValue
                Next mIndex
                rs.MoveNext
            Wend
        End If
        Call rs.Close
        Set rs = Nothing
        Call db.Close
        Set db = Nothing
    End If
    Exit Sub
FileOpen_ClickError:
    If Err.Number = cdlCancel Then
        Err.Clear
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub ResetList()
    Dim mCount As Long

    ListView1.ListItems.Clear
    mCount = ListView1.ColumnHeaders.Count
    If (mCount > 0) Then
        While (mCount > 0)
            Call ListView1.ColumnHeaders.Remove(1)
            mCount = mCount - 1
        Wend
    End If
End Sub

Private Sub DefineList(ByRef rs As ADODB.Recordset)
    Dim mCount As Long
    Dim mIndex As Long
    Dim mColumnWidth As Long

    mCount = rs.Fields.Count
    For mIndex = 0 To mCount - 1
        mColumnWidth = IIf(CLng(rs.Fields(mIndex).DefinedSize) > Len(CStr(rs.Fields(mIndex).Name)), rs.Fields(mIndex).DefinedSize, Len(CStr(rs.Fields(mIndex).Name)))
        mColumnWidth = mColumnWidth * 150
        Call ListView1.ColumnHeaders.Add(mIndex + 1, , rs.Fields(mIndex).Name, mColumnWidth)
    Next mIndex
End Sub
