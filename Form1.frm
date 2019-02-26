VERSION 5.00
Begin VB.Form frmDrawLine 
   BackColor       =   &H80000014&
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   18336
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   18336
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstY 
      Height          =   2160
      Left            =   480
      TabIndex        =   25
      Top             =   4920
      Width           =   1212
   End
   Begin VB.ListBox lstX 
      Height          =   2160
      Left            =   1740
      TabIndex        =   24
      Top             =   4920
      Width           =   1512
   End
   Begin VB.Frame frame1 
      Height          =   3072
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   17592
      Begin VB.HScrollBar HScrollE 
         Height          =   372
         Left            =   7500
         TabIndex        =   26
         Top             =   780
         Width           =   1152
      End
      Begin VB.HScrollBar hsbB 
         Height          =   432
         Left            =   1140
         TabIndex        =   8
         Top             =   2460
         Width           =   5232
      End
      Begin VB.HScrollBar hsbG 
         Height          =   492
         Left            =   1080
         TabIndex        =   7
         Top             =   1860
         Width           =   5472
      End
      Begin VB.HScrollBar hsbR 
         Height          =   372
         Left            =   1140
         TabIndex        =   6
         Top             =   1020
         Width           =   5292
      End
      Begin VB.HScrollBar hsbWidth 
         Height          =   492
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   5472
      End
      Begin VB.Label Label1 
         Caption         =   "eraser"
         Height          =   372
         Index           =   1
         Left            =   7140
         TabIndex        =   27
         Top             =   420
         Width           =   912
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   492
         Left            =   12360
         TabIndex        =   21
         Top             =   1020
         Width           =   432
      End
      Begin VB.Label Label10 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   432
         Left            =   11700
         TabIndex        =   20
         Top             =   1080
         Width           =   432
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   432
         Left            =   11100
         TabIndex        =   19
         Top             =   1080
         Width           =   372
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   432
         Left            =   10620
         TabIndex        =   18
         Top             =   1080
         Width           =   372
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Height          =   432
         Left            =   10020
         TabIndex        =   17
         Top             =   1080
         Width           =   372
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   492
         Left            =   12120
         TabIndex        =   16
         Top             =   360
         Width           =   432
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   492
         Left            =   11580
         TabIndex        =   15
         Top             =   360
         Width           =   432
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   432
         Left            =   11040
         TabIndex        =   14
         Top             =   420
         Width           =   432
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   372
         Left            =   10440
         TabIndex        =   13
         Top             =   420
         Width           =   432
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   372
         Index           =   0
         Left            =   9960
         TabIndex        =   11
         Top             =   420
         Width           =   372
      End
      Begin VB.Label lblcolor 
         BorderStyle     =   1  'Fixed Single
         Height          =   672
         Left            =   10920
         TabIndex        =   10
         Top             =   2220
         Width           =   3132
      End
      Begin VB.Label Label 
         Caption         =   "color"
         Height          =   192
         Index           =   4
         Left            =   9060
         TabIndex        =   9
         Top             =   2520
         Width           =   672
      End
      Begin VB.Label Label 
         Caption         =   "blue"
         Height          =   492
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   2460
         Width           =   852
      End
      Begin VB.Label Label 
         Caption         =   "green"
         Height          =   432
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1860
         Width           =   612
      End
      Begin VB.Label Label 
         Caption         =   "red"
         Height          =   372
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   1020
         Width           =   492
      End
      Begin VB.Label Label 
         Caption         =   "width"
         Height          =   312
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   792
      End
   End
   Begin VB.Label lblY 
      BorderStyle     =   1  'Fixed Single
      Height          =   912
      Left            =   300
      TabIndex        =   23
      Top             =   3600
      Width           =   1812
   End
   Begin VB.Label lblX 
      BorderStyle     =   1  'Fixed Single
      Height          =   552
      Left            =   360
      TabIndex        =   22
      Top             =   3720
      Width           =   1392
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   372
      Left            =   10860
      TabIndex        =   12
      Top             =   720
      Width           =   552
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu open 
         Caption         =   "open"
      End
      Begin VB.Menu clear 
         Caption         =   "clear"
      End
      Begin VB.Menu save 
         Caption         =   "save"
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "save as"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "close"
      End
      Begin VB.Menu exit 
         Caption         =   "exit"
      End
   End
End
Attribute VB_Name = "frmDrawLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gfDrawing As Integer
Const gnMaxPoints = 3000 ' n = global number
'gsng = global single
'Dim gsngX(1 To gnMaxPoints) As Single
'Dim gsngY(1 To gnMaxPoints) As Single
' actual number of points recorded
Dim gnNumPoints As Integer
'gstr = global string
Const gstrTitle = "LineDraw"
Dim r, g, b As Integer
Dim path, strFn, ans As String
Dim i As Long
Dim w As Long
Dim radius As Integer
' declares r ,s, b and line thickness
Dim arrayR(5000), arrayG(5000), arrayB(5000), arrayW(5000), arrayT(5000) As Double
Dim gsngX(5000), gsngY(5000) As Double
' need for thickness

Sub DrawCircle(X As Single, Y As Single)

Dim radius As Integer

     
   ' r = arrayR(1)
    ' g = arrayG(1)
     'b = arrayB(1)
radius = hsbWidth.Value


Circle (X, Y), radius, RGB(r, g, b)
If gnNumPoints < 5000 Then
    gnNumPoints = gnNumPoints + 1
    gsngX(gnNumPoints) = X
    gsngY(gnNumPoints) = Y
    lstX.AddItem str(gsngX(gnNumPoints))
    lstY.AddItem str(gsngX(gnNumPoints))
   



End If

End Sub
 

Sub DrawLines()
    Dim i As Integer
    i = 0
    CurrentX = gsngX(1)
    CurrentY = gsngY(1)
     'r = arrayR(i)
     'g = arrayG(i)
     'b = arrayB(i)
     r = arrayR(1)
     g = arrayG(1)
     b = arrayB(1)
     
    'Draw the first chircle
    'Circle (gsngX(1), gsngY(1)), 50
    Circle (gsngX(1), gsngY(1)), 1, RGB(arrayR(1), arrayG(1), arrayB(1))
    
    '- plot the restof the lines and circles
    'For i = 2 To gnNumPoints
     For i = 1 To gnNumPoints
        Line -(gsngX(i), gsngY(i))
        Circle (gsngX(i), gsngY(i)), 50
    Next i
    
End Sub

'Sub DrawCircleeee(X As Single, Y As Single)

' draw the circle

'Circle (X, Y), 50, RGB(r, g, b)

' too many poinyts ?


 '   If gnNumPoints < gnMaxPoints Then
' add one more point

'gnNumPoints = gnNumPoints + 1
' save x and y  crdinates

'gsngX(gnNumPoints) = X
'gsngY(gnNumPoints) = Y

 '   End If
 
 'End Sub






Private Sub clear_Click()
frmDrawLine.Cls


End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
gnNumPoints = 0
w = 1
gfDrawing = 0
r = 0
g = 0
b = 0
hsbR.Min = 0
hsbR.Max = 255

hsbG.Min = 0
hsbG.Max = 255

hsbB.Min = 0
hsbB.Max = 255
w = 20


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'gfDrawing = True
gfDrawing = 1

'frmDrawWidth.DrawWidth = w


DrawCircle X, Y
'-set form's currentX , currentY properties
CurrentX = X

CurrentY = Y
' draw sub
'as you draw ans as you open

' draw line

'open after >>

' open -reload circle and differentiating

' FIX

'array and stores it all

'array is ust a list it stores x ans y value and rgb value

' mouse up = draw other line
' mouse uo - iuse point as a point i look up
' oue u draw curlce on corener ( -200, 2000 )

'CIRCLE x Y 50 rgb(RGB)


End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If gfDrawing = 1 Then
    Line -(X, Y), RGB(r, g, b)
    lblX.Caption = lstX.ListCount
    lblY.Caption = lstY.ListCount
    
  ' DrawCircle(X, Y), 50
    DrawCircle X, Y
    
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
gfDrawing = 0
If gnNumPoints < 5000 Then
    gnNumPoints = gnNumPoints + 1
    X = -20
    Y = -20
    gsngX(gnNumPoints) = X
    gsngY(gnNumPoints) = Y
    lstX.AddItem X 'str(gsngX(gnNumPoints))
    lstY.AddItem Y 'str(gsngY(gnNumPoints))
    DrawCircle X, Y
End If
End Sub

Private Sub hsbWidth_Change()
hsbWidth.Value = radius + 1

'frmDrawLine.DrawWidth = hsbWidth.Value
End Sub
Private Sub hsbG_Change()

r = hsbR.Value
g = hsbG.Value
b = hsbB.Value
lblcolor.BackColor = RGB(r, g, b)



End Sub

Private Sub hsbR_Change()

r = hsbR.Value
g = hsbG.Value
b = hsbB.Value
lblcolor.BackColor = RGB(r, g, b)



End Sub
Private Sub hsbB_Change()

r = hsbR.Value
g = hsbG.Value
b = hsbB.Value
lblcolor.BackColor = RGB(r, g, b)



End Sub




Private Sub HScrollE_Change()
r = 255
g = 255
b = 255
lblcolor.BackColor = RGB(r, g, b)

End Sub

Private Sub lblcolor_Click()
' Path = "<File location>" + strfn + "dat"



End Sub

Private Sub mnuclose_Click()
gnNumPoints = 0
frmDrawLine.Cls
frmDrawLine.Caption = gstrTitle


End Sub


Private Sub mnuSaveas_Click()
' collect file name
'Dim strFn As String
'strFn = UCase$(Trim$(InputBox("Filename", "OpenFile")))
'open binary file for output
'Open strFn For Binary Access Write As #1

' save actual number of points
'Put #1, , gnNumPoints

' - local var for loop


'Dim i As Integer

' loop to actual number of points

'For i = 1 To gnNumPoints

    'save coordinates
    'Put #1, , gsngX(i)
    'Put #1, , gsngY(i)
'Next i
'Close #1
' reset form caption
'frmDrawLine.Caption = gstrTitle & "-" & strFn


ans = vbNo
Do While ans = vbNo
    Dim strFn As String
    strFn = UCase$(Trim$(InputBox("Filename", "OpenFile", "BOB")))
    path = "C:\Users\twish\Desktop\CP1 NEW NEW\" + strFn + ".dat"
    ans = MsgBox(path, vbYesNo, " ID this the path?")
Loop
If ans = vbYes Then
    Open path For Binary Access Write As #1
    Put #1, , gnNumPoints
    Dim i As Integer
    
    For i = 0 To gnNumPoints
        Put #1, , gsngX(i)
        Put #1, , gsngY(i)
        
' save ur sep R,G, B arrays and ur line width array'

    Next i
    Close #1
    
End If



End Sub

Private Sub open_Click()
'- collect filename feom user
'Dim strFn As String

strFn = UCase(Trim$(InputBox("Filenname", "OpenFile")))
' open binary file for input
'Open strFn For Binary Access Read As #1

'- save actual number of points

'Get #1, , gnNumPoints

'-Local Variable for loop
'Dim i As Integer

' loop to actual number of points

'For i = 1 To gnNumPoints

' collect coordinates from file
  '  Get #1, , gsngX(i)
 '   Get #1, , gsngY(i)
'Next i

' reset caption

'frmDrawLine.Caption = gstrTitle & "-" & strFn
'frmDrawLine.Cls
'DrawLines

ans = vbNo
Do While ans = vbNo
    strFn = UCase$(Trim$(InputBox("Filename", "OpenFile", "BOB")))
    path = "C:\Users\twish\Desktop\CP1 NEW NEW\" + strFn + ".dat"
    ans = MsgBox(path, vbYesNo, " ID this the path?")
Loop
If ans = vbYes Then
    Open path For Binary Access Read As #1
    Get #1, , gnNumPoints
    Dim i As Integer
    
    For i = 0 To gnNumPoints
        Get #1, , gsngX(i)
        Get #1, , gsngY(i)
        Get #1, , arrayR(1)
        Get #1, , arrayG(1)
        Get #1, , arrayB(1)
        
' save ur sep R,G, B arrays and ur line width array'

    Next i
    Close #1
    
    
    frmDrawLine.Cls
    DrawLines
    
    
End If

End Sub








Private Sub save_Click()
ans = vbNo
'Do While ans = vbNo
    Dim strFn As String
    strFn = UCase$(Trim$(InputBox("Filename", "OpenFile", "BOB")))
    path = "C:\Users\twish\Desktop\CP1 NEW NEW\" + strFn + ".dat"
    'ans = MsgBox(path, vbYesNo, " ID this the path?")
'Loop
'If ans = vbYes Then
    Open path For Binary Access Write As #1
    Put #1, , gnNumPoints
    Dim i As Integer
    
    For i = 0 To gnNumPoints
        Put #1, , gsngX(i)
        Put #1, , gsngY(i)
        Put #1, , arrayR(1)
        Put #1, , arrayG(1)
        Put #1, , arrayB(1)
        
        
        
' save ur sep R,G, B arrays and ur line width array'

    Next i
    Close #1
    
'End If
End Sub
