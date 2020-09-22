VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bimap Bits Example"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   3150
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   17
      Top             =   75
      Width           =   3000
   End
   Begin VB.CommandButton cmdRGB2GBR 
      Caption         =   "RGB->GBR"
      Height          =   300
      Left            =   5100
      TabIndex        =   16
      Top             =   3150
      Width           =   975
   End
   Begin VB.CommandButton cmdRGB2GRB 
      Caption         =   "RGB->GRB"
      Height          =   300
      Left            =   4125
      TabIndex        =   15
      Top             =   3150
      Width           =   975
   End
   Begin VB.CommandButton cmdRGB2BRG 
      Caption         =   "RGB->BRG"
      Height          =   300
      Left            =   3150
      TabIndex        =   14
      Top             =   3150
      Width           =   975
   End
   Begin VB.CommandButton cmdRGB2BGR 
      Caption         =   "RGB->BGR"
      Height          =   300
      Left            =   2175
      TabIndex        =   13
      Top             =   3150
      Width           =   975
   End
   Begin VB.CommandButton cmdRGB2RBG 
      Caption         =   "RGB->RBG"
      Height          =   300
      Left            =   1200
      TabIndex        =   12
      Top             =   3150
      Width           =   975
   End
   Begin VB.CommandButton cmdGreyScale 
      Caption         =   "Grey Scale"
      Height          =   300
      Left            =   75
      TabIndex        =   11
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdFlipVert 
      Caption         =   "Flip Vert"
      Height          =   300
      Left            =   1050
      TabIndex        =   10
      Top             =   2700
      Width           =   975
   End
   Begin VB.CommandButton cmdFlipHoriz 
      Caption         =   "Flip Horiz"
      Height          =   300
      Left            =   75
      TabIndex        =   9
      Top             =   2700
      Width           =   975
   End
   Begin VB.HScrollBar HSColorShift 
      Height          =   195
      Index           =   2
      LargeChange     =   5
      Left            =   3000
      Max             =   255
      Min             =   -255
      TabIndex        =   5
      Top             =   2400
      Width           =   3165
   End
   Begin VB.HScrollBar HSColorShift 
      Height          =   195
      Index           =   1
      LargeChange     =   5
      Left            =   3000
      Max             =   255
      Min             =   -255
      TabIndex        =   4
      Top             =   2625
      Width           =   3165
   End
   Begin VB.HScrollBar HSColorShift 
      Height          =   195
      Index           =   0
      LargeChange     =   5
      Left            =   3000
      Max             =   255
      Min             =   -255
      TabIndex        =   3
      Top             =   2850
      Width           =   3165
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Reset"
      Height          =   300
      Left            =   75
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdNegative 
      Caption         =   "Negative"
      Height          =   300
      Left            =   1050
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   75
      Picture         =   "Form1.frx":7972
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   75
      Width           =   3000
   End
   Begin VB.Label Label3 
      Caption         =   "Blue Shift"
      Height          =   165
      Left            =   2100
      TabIndex        =   8
      Top             =   2850
      Width           =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Green Shift"
      Height          =   165
      Left            =   2100
      TabIndex        =   7
      Top             =   2625
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Red Shift"
      Height          =   165
      Left            =   2100
      TabIndex        =   6
      Top             =   2400
      Width           =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* CODED BY: BattleStorm
'* EMAIL: battlestorm@cox.net
'* UPDATED: 08/08/2002

'* PURPOSE: Demonstrates how to use
'*     GetBitmapBits and SetBitmapBits.
'*     Also shows how to make a Negative
'*     of a picture, flip horizontally,
'*     flip vertically and shift RGB colors.
'*     Updated, now also does grey scale and
'*     swaps RGB values 5 different ways.

'* COPYRIGHT: This program and source
'*     code is freeware and can be copied
'*     and/or distributed as long as you
'*     mention the original author. I am
'*     not responsible for any harm as the
'*     outcome of using any of this code.

'* CREDITS: Special thanks go out to
'*     www.allapi.net for their wonderful
'*     API Guide that I use all the time as
'*     my only API reference. Some of this
'*     code is borrowed from their GetBitmapBits
'*     examples and modified slightly for X, Y
'*     coordinate manipulation of pixels.

'* NOTES: Only works reliably for 24 and 32 bit
'*        color depths. I am working on a 16 and
'*        8 bit solution using bit shifting to
'*        extract the RGB values from the pixels.

'API declarations
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

'Type declarations
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type BSBITMAP
    Info As BITMAP
    Bits() As Byte
End Type

'Class declarations
Private CodeTimer As clsTimer

'Private variable declarations
Private SBuffer As BSBITMAP
Private DBuffer As BSBITMAP
Private tColor As Integer
Private temp As Long

'This sub uses GetBitmapBits and SetBitmapBits to grab the
'pixels from a picturebox and just copy it to another.
Private Sub cmdCopy_Click()
    Call StartProcess
    
    'Copy every pixel straight over, this sub
    'can be used for the basis of just about
    'any pixel manupulation you can think of.
    For Y = 0 To SBuffer.Info.bmHeight - 1
        For X = 0 To SBuffer.Info.bmWidth - 1
            DBuffer.Bits(2, X, Y) = SBuffer.Bits(2, X, Y) 'Red Bits
            DBuffer.Bits(1, X, Y) = SBuffer.Bits(1, X, Y) 'Green Bits
            DBuffer.Bits(0, X, Y) = SBuffer.Bits(0, X, Y) 'Blue Bits
        Next X
    Next Y
    
    Call StopProcess
End Sub

'This sub uses GetBitmapBits and SetBitmapBits to grab the
'pixels from a picturebox and then flip all bits Vertically.
Private Sub cmdFlipVert_Click()
    Call StartProcess
    
    'Flip bits Vertically
    temp = SBuffer.Info.bmHeight - 1
    For Y = 0 To SBuffer.Info.bmHeight - 1
        For X = 0 To SBuffer.Info.bmWidth - 1
            DBuffer.Bits(2, X, Y) = SBuffer.Bits(2, X, temp - Y) 'Red Bits
            DBuffer.Bits(1, X, Y) = SBuffer.Bits(1, X, temp - Y) 'Green Bits
            DBuffer.Bits(0, X, Y) = SBuffer.Bits(0, X, temp - Y) 'Blue Bits
        Next X
    Next Y
    
    Call StopProcess
End Sub

Private Sub cmdGreyScale_Click()
    Call StartProcess
    
    'Loop thru each Red, Green and Blue portion of each
    'pixel and turn it to it's negative color
    For Y = 0 To SBuffer.Info.bmHeight - 1
        For X = 0 To SBuffer.Info.bmWidth - 1
            tColor = (SBuffer.Bits(2, X, Y) * 0.222) + _
                     (SBuffer.Bits(1, X, Y) * 0.707) + _
                     (SBuffer.Bits(0, X, Y) * 0.071)
            DBuffer.Bits(2, X, Y) = tColor 'Red Bits
            DBuffer.Bits(1, X, Y) = tColor 'Green Bits
            DBuffer.Bits(0, X, Y) = tColor 'Blue Bits
        Next X
    Next Y
    
    Call StopProcess
End Sub

'This sub uses GetBitmapBits and SetBitmapBits to grab the
'pixels from a picturebox and turn each Red, Blue and Green
'value to its negative color. It then repaints the picture.
Private Sub cmdNegative_Click()
    Call StartProcess
    
    'Loop thru each Red, Green and Blue portion of each
    'pixel and turn it to it's negative color
    For Y = 0 To SBuffer.Info.bmHeight - 1
        For X = 0 To SBuffer.Info.bmWidth - 1
            DBuffer.Bits(2, X, Y) = 255 - SBuffer.Bits(2, X, Y) 'Red Bits
            DBuffer.Bits(1, X, Y) = 255 - SBuffer.Bits(1, X, Y) 'Green Bits
            DBuffer.Bits(0, X, Y) = 255 - SBuffer.Bits(0, X, Y) 'Blue Bits
        Next X
    Next Y
    
    Call StopProcess
End Sub

'This sub uses GetBitmapBits and SetBitmapBits to grab the
'pixels from a picturebox and then flip all bits Horizontally.
Private Sub cmdFlipHoriz_Click()
    Call StartProcess
    
    'Flip bits Horizontally
    temp = SBuffer.Info.bmWidth - 1
    For Y = 0 To SBuffer.Info.bmHeight - 1
        For X = 0 To SBuffer.Info.bmWidth - 1
            DBuffer.Bits(2, X, Y) = SBuffer.Bits(2, temp - X, Y) 'Red Bits
            DBuffer.Bits(1, X, Y) = SBuffer.Bits(1, temp - X, Y) 'Green Bits
            DBuffer.Bits(0, X, Y) = SBuffer.Bits(0, temp - X, Y) 'Blue Bits
        Next X
    Next Y
    
    Call StopProcess
End Sub

Private Sub cmdRGB2BGR_Click()
    Call StartProcess
    
    'Swap Colors
    For Y = 0 To SBuffer.Info.bmHeight - 1
        For X = 0 To SBuffer.Info.bmWidth - 1
            DBuffer.Bits(2, X, Y) = SBuffer.Bits(0, X, Y)
            DBuffer.Bits(1, X, Y) = SBuffer.Bits(1, X, Y)
            DBuffer.Bits(0, X, Y) = SBuffer.Bits(2, X, Y)
        Next X
    Next Y
    
    Call StopProcess
End Sub

Private Sub cmdRGB2BRG_Click()
    Call StartProcess
    
    'Swap Colors
    For Y = 0 To SBuffer.Info.bmHeight - 1
        For X = 0 To SBuffer.Info.bmWidth - 1
            DBuffer.Bits(2, X, Y) = SBuffer.Bits(0, X, Y)
            DBuffer.Bits(1, X, Y) = SBuffer.Bits(2, X, Y)
            DBuffer.Bits(0, X, Y) = SBuffer.Bits(1, X, Y)
        Next X
    Next Y
    
    Call StopProcess
End Sub

Private Sub cmdRGB2GBR_Click()
    Call StartProcess
    
    'Swap Colors
    For Y = 0 To SBuffer.Info.bmHeight - 1
        For X = 0 To SBuffer.Info.bmWidth - 1
            DBuffer.Bits(2, X, Y) = SBuffer.Bits(1, X, Y)
            DBuffer.Bits(1, X, Y) = SBuffer.Bits(0, X, Y)
            DBuffer.Bits(0, X, Y) = SBuffer.Bits(2, X, Y)
        Next X
    Next Y
    
    Call StopProcess
End Sub

Private Sub cmdRGB2GRB_Click()
    Call StartProcess
    
    'Swap Colors
    For Y = 0 To SBuffer.Info.bmHeight - 1
        For X = 0 To SBuffer.Info.bmWidth - 1
            DBuffer.Bits(2, X, Y) = SBuffer.Bits(1, X, Y)
            DBuffer.Bits(1, X, Y) = SBuffer.Bits(2, X, Y)
            DBuffer.Bits(0, X, Y) = SBuffer.Bits(0, X, Y)
        Next X
    Next Y
    
    Call StopProcess
End Sub

Private Sub cmdRGB2RBG_Click()
    Call StartProcess
    
    'Swap Colors
    For Y = 0 To SBuffer.Info.bmHeight - 1
        For X = 0 To SBuffer.Info.bmWidth - 1
            DBuffer.Bits(2, X, Y) = SBuffer.Bits(2, X, Y)
            DBuffer.Bits(1, X, Y) = SBuffer.Bits(0, X, Y)
            DBuffer.Bits(0, X, Y) = SBuffer.Bits(1, X, Y)
        Next X
    Next Y
    
    Call StopProcess
End Sub

'This sub uses GetBitmapBits and SetBitmapBits to grab the
'pixels from a picturebox and shift the Red, Green and Blue
'color values depending on scroll bar values.
Private Sub HSColorShift_Change(Index As Integer)
    Call StartProcess
    
    'Shift Bit color depending on scroll bar values
    For Y = 0 To SBuffer.Info.bmHeight - 1
        For X = 0 To SBuffer.Info.bmWidth - 1
            For Z = 0 To 2
                tColor = SBuffer.Bits(Z, X, Y) + HSColorShift(Z).Value
                'Check for valid color ranges
                Select Case tColor
                    Case Is < 0
                        DBuffer.Bits(Z, X, Y) = 0
                    Case 0 To 255
                        DBuffer.Bits(Z, X, Y) = tColor
                    Case Is > 255
                        DBuffer.Bits(Z, X, Y) = 255
                End Select
            Next Z
        Next X
    Next Y
    
    Call StopProcess
End Sub

'Just used to consolidate common code
Private Sub StartProcess()
    'Initialize CodeTimer
    Set CodeTimer = New clsTimer
    
    'Start CodeTimer
    CodeTimer.StartTimer
    
    'Grab picture's pixels and load to Bit array
    GetBitmapBits Pic1.Image, SBuffer.Info.bmWidthBytes * SBuffer.Info.bmHeight, SBuffer.Bits(0, 0, 0)
End Sub

'Just used to consolidate common code
Private Sub StopProcess()
    'Load Bit array to picture box
    SetBitmapBits Pic2.Image, DBuffer.Info.bmWidthBytes * DBuffer.Info.bmHeight, DBuffer.Bits(0, 0, 0)
    
    'SetBitmapBits normally triggers a redraw event,
    'but just in case it doesn't, we'll do one now
    Pic2.Refresh
    
    'Stop CodeTimer
    CodeTimer.StopTimer
    
    'Display CodeTimer results in Form's caption
    Me.Caption = "Processing Time: " & CodeTimer.Elasped & " ms"
End Sub

Private Sub Form_Load()
    'Inform user that program will run alot faster when compiled
    If App.LogMode = 0 Then
        MsgBox "Compile Me - I'll Run Alot Faster!"
    End If
  
    'Get information about picture boxes and declare arrays to hold bits
    GetObject Pic1.Image, Len(SBuffer.Info), SBuffer.Info
    ReDim SBuffer.Bits(0 To SBuffer.Info.bmWidthBytes / SBuffer.Info.bmWidth - 1, _
                       0 To SBuffer.Info.bmWidth - 1, _
                       0 To SBuffer.Info.bmHeight - 1) As Byte

    GetObject Pic2.Image, Len(DBuffer.Info), DBuffer.Info
    ReDim DBuffer.Bits(0 To DBuffer.Info.bmWidthBytes / DBuffer.Info.bmWidth - 1, _
                       0 To DBuffer.Info.bmWidth - 1, _
                       0 To DBuffer.Info.bmHeight - 1) As Byte
                       
    'If color depth is not 24 or 32 bit color exit program
    If SBuffer.Info.bmBitsPixel < 24 Then
        MsgBox "Desktop color must be 24 or 32 bit" & vbCrLf & _
               "for program to function properly." & vbCrLf & _
               "Please exit and change desktop color" & vbCrLf & _
               "depth before running program."
    End If
End Sub
