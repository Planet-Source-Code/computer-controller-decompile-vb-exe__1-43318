VERSION 5.00
Begin VB.Form FormFinder 
   AutoRedraw      =   -1  'True
   Caption         =   "Form Finder"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "C:\file.exe"
      Top             =   0
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   3735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "FormFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub FormFinder(FileName As String, Lst)
Dim Countt As Long
    FreeF = FreeFile
    Open FileName$ For Binary Access Read As #FreeF
     BByte = 32000
     Countt = 1
     For CountDown = 1 To LOF(FreeF) Step BByte
Again:

     FChr$ = Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(5) + Chr(0)
     FChr1$ = Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(7) + Chr(0)
     FChr2$ = Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(6) + Chr(0)
     FChr3$ = Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(8) + Chr(0)
        
        SpaceByte$ = Space(BByte)
        Get #FreeF, CountDown, SpaceByte$

        If InStr(Countt, SpaceByte$, FChr$, 1) Then
            a = InStr(Countt, SpaceByte$, FChr$, 1) + 6
            For d = 1 To 32
            If Mid(SpaceByte$, a, 1) = Chr(d) Then Countt = a: GoTo Again
            Next
            If InStr(a, SpaceByte$, Chr(1)) Then
                b = InStr(a, SpaceByte$, Chr(1)) - 2
                Lst = Lst & vbCrLf & "Name - " & Mid(SpaceByte$, a, b - a)
                If InStr(b + 1, SpaceByte$, Chr(25) + Chr(1)) Then
                    C = InStr(b + 1, SpaceByte$, Chr(25) + Chr(1))
                    Lst = Lst & "   Caption - " & Mid(SpaceByte$, b + 5, C - b - 6)
                    Countt = C
                Else
                    Countt = b
                End If
            Else
            Countt = a
            End If
            DoEvents
            GoTo Again
        ElseIf InStr(Countt, SpaceByte$, FChr1$, 1) Then
            a = InStr(Countt, SpaceByte$, FChr1$, 1) + 6
            For d = 1 To 32
            If Mid(SpaceByte$, a, 1) = Chr(d) Then Countt = a: GoTo Again
            Next
            If InStr(a, SpaceByte$, Chr(1)) Then
                b = InStr(a, SpaceByte$, Chr(1)) - 2
                Lst = Lst & vbCrLf & "Name - " & Mid(SpaceByte$, a, b - a)
                If InStr(b + 1, SpaceByte$, Chr(25) + Chr(1)) Then
                    C = InStr(b + 1, SpaceByte$, Chr(25) + Chr(1))
                    Lst = Lst & "   Caption - " & Mid(SpaceByte$, b + 5, C - b - 6)
                    Countt = C
                Else
                    Countt = b
                End If
            Else
            Countt = a
            End If
            DoEvents
            GoTo Again
        ElseIf InStr(Countt, SpaceByte$, FChr2$, 1) Then
            a = InStr(Countt, SpaceByte$, FChr2$, 1) + 6
            For d = 1 To 32
            If Mid(SpaceByte$, a, 1) = Chr(d) Then Countt = a: GoTo Again
            Next
            If InStr(a, SpaceByte$, Chr(1)) Then
                b = InStr(a, SpaceByte$, Chr(1)) - 2
                Lst = Lst & vbCrLf & "Name - " & Mid(SpaceByte$, a, b - a)
                If InStr(b + 1, SpaceByte$, Chr(25) + Chr(1)) Then
                    C = InStr(b + 1, SpaceByte$, Chr(25) + Chr(1))
                    Lst = Lst & "   Caption - " & Mid(SpaceByte$, b + 5, C - b - 6)
                    Countt = C
                Else
                    Countt = b
                End If
            Else
            Countt = a
            End If
            DoEvents
            GoTo Again
        ElseIf InStr(Countt, SpaceByte$, FChr3$, 1) Then
            a = InStr(Countt, SpaceByte$, FChr3$, 1) + 6
            For d = 1 To 32
            If Mid(SpaceByte$, a, 1) = Chr(d) Then Countt = a: GoTo Again
            Next
            If InStr(a, SpaceByte$, Chr(1)) Then
                b = InStr(a, SpaceByte$, Chr(1)) - 2
                Lst = Lst & vbCrLf & "Name - " & Mid(SpaceByte$, a, b - a)
                If InStr(b + 1, SpaceByte$, Chr(25) + Chr(1)) Then
                    C = InStr(b + 1, SpaceByte$, Chr(25) + Chr(1))
                    Lst = Lst & "   Caption - " & Mid(SpaceByte$, b + 5, C - b - 6)
                    Countt = C
                Else
                    Countt = b
                End If
            Else
                Countt = a
            End If
            DoEvents
            GoTo Again
        Else
        End If
        Next
    Close #FreeF
End Sub

Private Sub Command1_Click()
FormFinder Text2.Text, a
Text1.Text = a
End Sub

Private Sub Form_Resize()
Text1.Width = Me.Width - 90
Text1.Height = Me.Height - 750
Command1.Left = Me.Width - Command1.Width - 120
Text2.Width = Command1.Left - 10
End Sub
