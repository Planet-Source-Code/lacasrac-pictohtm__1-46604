VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PICTURE TO HTML SITE!!!! created by Laca 2003 Jun"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opt2 
      Caption         =   "1024*768"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4200
      Width           =   2295
   End
   Begin VB.OptionButton opt1 
      Caption         =   " 800*600"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3840
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.CheckBox cen 
      Caption         =   "Center Pic to HTM"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
   End
   Begin VB.HScrollBar steps 
      Height          =   255
      Left            =   840
      Max             =   10
      Min             =   1
      TabIndex        =   10
      Top             =   3120
      Value           =   1
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show me The page!"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7200
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   2160
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox htmname 
      Height          =   285
      Left            =   1080
      MaxLength       =   12
      TabIndex        =   5
      Text            =   "Htm.html"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Picture"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kill HTM file"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create HTML site"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   2640
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   291
      TabIndex        =   1
      Top             =   120
      Width           =   4365
   End
   Begin VB.TextBox log 
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4650
      Width           =   6975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Steps:"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File name:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Dim p_array() As Long

'created by Laca in Hungary, 2003
'lostinwar@freemail.hu
'http://lacasrac.srv.hu

Private Sub Command1_Click()

With cmd
    .DialogTitle = "open Pic"
    .CancelError = False
    .Filter = "Pictures (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
    .ShowOpen
    If Len(.FileName) = 0 Then
        Exit Sub
    End If
    sfile = .FileName
End With
pic.Picture = LoadPicture(sfile)
pic2.Picture = LoadPicture(sfile)
    
End Sub

Public Sub Map()

ReDim p_array(3, pic.ScaleWidth - 1, pic.ScaleHeight - 1)
For x = 0 To pic.ScaleWidth - 1
    For y = 0 To pic.ScaleHeight - 1
        c = GetPixel(pic.hdc, x, y)
        
        r = c And 255
        g = (c \ 256 ^ 1) And 255
        b = (c \ 256 ^ 2) And 255
        
        p_array(1, x, y) = r
        p_array(2, x, y) = g
        p_array(3, x, y) = b
    Next y
Next x

End Sub



Public Sub ToHTM()

Dim x           As Long
Dim y           As Long
Dim x2          As Integer
Dim r           As String
Dim g           As String
Dim b           As String
Dim c           As String
Dim htm_data    As String
Dim file        As String
Dim whois       As String, text As String, text1 As String

text = "<HTML><BODY bgcolor=#000000><center><pre><font face=arial size=-10>"
file = App.Path + "\" + htmname.text

'Begin
Open file For Output As #1
     Print #1, text + vbCrLf
Close #1

whois = "<!--Picture to Html site!!!        -->" + vbCrLf + _
        "<!--Created by Laca in Hungary 2003-->" + vbCrLf + _
        "<!--E-mail: lostinwar@freemail.hu  -->" + vbCrLf + _
        "<!--Site: http://lacasrac.srv.hu   -->" + vbCrLf
        
Open file For Append As #1
     Print #1, whois + vbCrLf
Close #1

'BODY
Open file For Append As #1
For x = 0 To pic.ScaleWidth - 1
        For y = 0 To pic.ScaleHeight - 1
            
            r = CStr(Hex(p_array(1, x, y)))
            g = CStr(Hex(p_array(2, x, y)))
            b = CStr(Hex(p_array(3, x, y)))

            If Len(r) = 1 Then r = "0" + r
            If Len(g) = 1 Then g = "0" + g
            If Len(b) = 1 Then b = "0" + b
            c = "#" + r + g + b
            
            If cen.Value = Checked Then
                If opt1.Value = True Then
                    x4 = ((800 / 2) - (steps.Value * (pic.ScaleWidth / 2)) + steps.Value * x - 10)
                    y4 = ((600 / 2) - (steps.Value * (pic.ScaleHeight / 2)) + steps.Value * y - 100)
                ElseIf opt2.Value = True Then
                    x4 = ((1024 / 2) - (steps.Value * (pic.ScaleWidth / 2)) + steps.Value * x - 10)
                    y4 = ((768 / 2) - (steps.Value * (pic.ScaleHeight / 2)) + steps.Value * y - 100)
                End If
            Else
                x4 = (x * steps.Value)
                y4 = (y * steps.Value)
            End If
            text1 = "<FONT STYLE='" + _
                    "LEFT: " + CStr(x4) + "px;" + _
                    "TOP: " + CStr(y4) + "px;" + _
                    "position=absolute' color=" + _
                    CStr(c) + ">.</FONT>"
                
            Print #1, text1
           DoEvents
        Next y
    Next x
Close #1

'End
text = "</font></body></html>"
Open file For Append As #1
    Print #1, text + vbCrLf
Close #1

Log.text = file + " | " + CStr(Int(FileLen(file) / 1024)) + " Kb |"

End Sub

Private Sub Command2_Click()
    
    If pic.ScaleHeight > 128 Or pic.ScaleWidth > 128 Then
        'Max 128*128 because internet explorer/win msg:
        '"Not responding..."
        
        MsgBox "Max 128*128 image!"
        Exit Sub
    End If
    
    Call Map
    Call ToHTM
    MsgBox "Ready now !"
End Sub


Private Sub Command3_Click()
    file = App.Path + "\" + htmname.text
    Kill file
End Sub

Private Sub Command4_Click()
    GoToUrl App.Path + "\" + htmname.text
End Sub


