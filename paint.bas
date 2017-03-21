Attribute VB_Name = "paint"
'---------------------------------------------
'-             VB Programming Project        -
'-                                           -
'-  Project Name  : Paint                    -
'-  Coded By      : Mohammad Shams Javi      -
'-  Mail          : M.Shams.J@Gmail.com      -
'---------------------------------------------

Public fname As String
Public PicChanged As Boolean

Public Type POINT
        X As Long
        Y As Long
End Type
Public p1 As POINT, p2 As POINT


Sub draw(i As Byte, size As Byte, ByVal btn As Byte, ByVal X As Long, ByVal Y As Long)
Randomize Timer
On Error Resume Next

With fp.pic
.ForeColor = fp.c(btn).FillColor
.FillColor = fp.c(btn).FillColor
.DrawWidth = size + 1
.FillStyle = 1
End With

Select Case i
Case 0  'get color
    fp.c(btn).FillColor = fp.pic.POINT(X, Y)
Case 3  'brush
    fp.pic.DrawWidth = (size + 2) ^ 2
    fp.pic.Line -(X, Y)
    fp.pic.Circle (X, Y), (size + 2) ^ 2
Case 5  'ruber
    fp.pic.DrawWidth = (size + 2) ^ 3
    fp.pic.ForeColor = vbWhite
    fp.pic.Line -(X, Y)
Case 6  'spray
    hit = (size + 2) ^ 3
    For lP = 0 To hit
    fp.pic.PSet (X + 10 * hit * Rnd * Sin(lP * hit), Y + 10 * hit * Rnd * Cos(lP * hit)), fp.pic.ForeColor
    Next
Case 7  'pencil
    fp.pic.Line -(X, Y)
Case 8  'line
    fp.pic.Line (p1.X, p1.Y)-(p2.X, p2.Y), fp.pic.POINT(X, Y)
    fp.pic.Line (p1.X, p1.Y)-(X, Y)
    p2.X = X: p2.Y = Y
Case 9  'rectangle
    If fp.f(0).Value Then
        fp.pic.Line (p1.X, p1.Y)-(X, Y), fp.pic.ForeColor, BF
    Else
        fp.pic.Line (p1.X, p1.Y)-(p2.X, p2.Y), fp.pic.POINT(X, Y), B
        fp.pic.Line (p1.X, p1.Y)-(X, Y), fp.pic.ForeColor, B
    End If
    p2.X = X: p2.Y = Y

Case 10 'elipse
    If fp.f(0).Value Then
        fp.pic.FillStyle = 0
        fp.pic.Circle (p1.X, p1.Y), IIf(Abs(X - p1.X) > Abs(Y - p1.Y) _
        , Abs(X - p1.X), Abs(Y - p1.Y)), fp.pic.ForeColor, 0, 0, IIf(Abs(X - p1.X) > _
        Abs(Y - p1.Y), Abs(Y / X), Abs(X / Y))
    Else
        fp.pic.Circle (p1.X, p1.Y), IIf(Abs(X - p1.X) > Abs(Y - p1.Y) _
        , Abs(X - p1.X), Abs(Y - p1.Y)), fp.pic.ForeColor, 0, 0, IIf(Abs(X - p1.X) > _
        Abs(Y - p1.Y), Abs(Y / X), Abs(X / Y))
        fp.pic.Circle (p1.X, p1.Y), IIf(Abs(p2.X - p1.X) > Abs(p2.Y - p1.Y) _
        , Abs(p2.X - p1.X), Abs(p2.Y - p1.Y)), fp.pic.POINT(X, Y), 0, 0, IIf(Abs(p2.X - p1.X) > Abs(p2.Y - p1.Y) _
        , Abs(p2.Y / p2.X), Abs(p2.X / p2.Y))
    End If
    p2.X = X: p2.Y = Y
    
Case 11     'circle
    If fp.f(0).Value Then
        fp.pic.FillStyle = 0
        fp.pic.Circle (p1.X, p1.Y), IIf(Abs(X - p1.X) > Abs(Y - p1.Y), Abs(X - p1.X), Abs(Y - p1.Y)), fp.pic.ForeColor
    Else
        fp.FillColor = 5
        fp.pic.Circle (p1.X, p1.Y), IIf(Abs(X - p1.X) > Abs(Y - p1.Y), Abs(X - p1.X), Abs(Y - p1.Y)), fp.pic.ForeColor
        fp.pic.Circle (p1.X, p1.Y), IIf(Abs(p2.X - p1.X) > Abs(p2.Y - p1.Y), Abs(p2.X - p1.X), Abs(p2.Y - p1.Y)), fp.pic.POINT(X, Y)
    End If
    p2.X = X: p2.Y = Y
End Select

End Sub

Public Sub FillCol(ByVal btn As Byte, ByVal X As Long, ByVal Y As Long)
't=top d=down l=left r=right
Dim t As Boolean, d As Boolean, l As Boolean, r As Boolean
Dim col As Long, a As Long, b As Long, i As Long, j As Long
'col=color  a=x,b=y

MsgBox "Because of VB bugs !! it cant work successfully."
'Note : "Pset" routin in picture control , cant set color of
'       one pixel. setted pixel in this command is wrong with
'       scale mode of picture box.
'BUT... i try to found another way to solve this problem .comming soon.
Exit Sub

With fp.pic
.ForeColor = fp.c(btn).FillColor
.DrawWidth = 1
.DrawStyle = 2
col = .POINT(X, Y)      'get pixel color
End With
a = X: b = Y  'set a,b cordination

'it is linear algorithm , to start from a cell in a table ,
'   and jump to all around cells
Do While Not (t Or d Or l Or r)

For q = 1 To j  'right cells
    a = a + i   'x=x+1
    If fp.pic.POINT(a, b) = col Then
        fp.pic.PSet (a, b), fp.c(btn).FillColor
    Else
        r = True
    End If
Next
For q = 1 To j  'down cells
    b = b + i   'y=y-1
    If fp.pic.POINT(a, b) = col Then
        fp.pic.PSet (a, b), fp.c(btn).FillColor
    Else
        d = True
    End If
Next

j = j + 1      'increase counter after 2 jump (R&D)

For q = 1 To j  'left cells
    a = a - i   'x=x-1
    If fp.pic.POINT(a, b) = col Then
        fp.pic.PSet (a, b), fp.c(btn).FillColor
    Else
        l = True
    End If
Next
For q = 1 To j  'top cells
    b = b - i   'y=y-1
    If fp.pic.POINT(a, b) = col Then
        fp.pic.PSet (a, b), fp.c(btn).FillColor
    Else
        t = True
    End If
Next

i = 1
j = j + 1        'increase counter after 2 jump (L&T)
Loop

End Sub

Public Sub SetCur(index As Integer)
With fp.pic     'set cursor
.MousePointer = 99
Select Case index
Case 0: .MouseIcon = LoadPicture(App.Path + "\pics\col.cur")
Case 4: .MouseIcon = LoadPicture(App.Path + "\pics\fil.cur")
Case 6: .MouseIcon = LoadPicture(App.Path + "\pics\spr.cur")
Case 7: .MouseIcon = LoadPicture(App.Path + "\pics\pen.cur")
Case 5, 3: .MouseIcon = LoadPicture(App.Path + "\pics\c.cur")
Case 8 To 11, 2: .MouseIcon = LoadPicture(App.Path + "\pics\z.cur")
Case 1: .MousePointer = 0
End Select
End With
End Sub


Public Sub SetEn(index As Integer)
With fp
Select Case index
Case 9 To 11: .f(0).Value = True
    .s(0).Enabled = False
    .s(1).Enabled = False
    .s(2).Enabled = False
    .f(0).Enabled = True
    .f(1).Enabled = True
Case 5 To 8, 3: .s(0).Value = True
    .s(0).Enabled = True
    .s(1).Enabled = True
    .s(2).Enabled = True
    .f(0).Enabled = False
    .f(1).Enabled = False
Case Else
    .s(0).Enabled = False
    .s(1).Enabled = False
    .s(2).Enabled = False
    .f(0).Enabled = False
    .f(1).Enabled = False
End Select
End With
End Sub
