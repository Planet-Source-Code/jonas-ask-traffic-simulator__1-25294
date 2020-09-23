Attribute VB_Name = "Publics"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCCOPY = &HCC0020

Public Type Coor
 X As Integer
 Y As Integer
End Type
Public Type Cars
 X As Integer
 Y As Integer
 Xdire As Single
 Ydire As Single
 LastPos As Coor
 Type As Integer
End Type

Public FirstTime As Boolean

Public GameSpeed As Long
Public NumCars As Integer


Public ColAsp As Long
Public ColGround As Long

Public NPosX As Integer
Public NPosY As Integer

Public Nett() As Integer
Public Bredde As Integer
Public Hoyde As Integer
Public Car() As Cars
Public Const Size As Integer = 11


Public Sub SetUpNet()
Dim X, Y As Integer
    
    FirstTime = True '
    
    ReDim Nett(0 To Bredde, 0 To Hoyde)
    
    For Y = 0 To Hoyde
        For X = 0 To Bredde
            If Form1.picKilde.Point(X, Y) = 0 Then Nett(X, Y) = 1
        Next X
    Next Y
    
    
    ' Make Cars
    ReDim Car(1 To NumCars)
    For N = 1 To UBound(Car)
        Do
            DoEvents
            X = rndTall(0, Bredde)
            Y = rndTall(0, Hoyde)
        Loop Until Nett(X, Y) = 1
        
        Car(N).X = X
        Car(N).Y = Y
        Car(N).Type = rndTall(1, 3)
        SetFirstDire N
    Next N
End Sub


Public Sub Paintboard(Pic As PictureBox)
Dim tColor
    t = GetTickCount
    Pic.Cls
    
    If FirstTime Then
        FirstTime = False
        
        For Y = 0 To Hoyde
        For X = 0 To Bredde
            Select Case Nett(X, Y)
            Case 1
                Form1.picBuffer.Line (X * Size, Y * Size)-Step(Size, Size), ColAsp, BF
            Case 0
                BitBlt Form1.picBuffer.hDC, X * Size, Y * Size, Size, Size, Form1.pGround(rndTall(0, 4)).hDC, 0, 0, SRCCOPY
            End Select
            
        Next X
        Next Y
        
    End If
    
    BitBlt Pic.hDC, 0, 0, Bredde * Size, Hoyde * Size, Form1.picBuffer.hDC, 0, 0, SRCCOPY
    
    Form1.picMain.Line (-1, -1)-(Bredde * Size, Hoyde * Size), , B
    'cars
    Dim Minus As Integer
    For a = 1 To UBound(Car)
        Minus = 1 'Dette er en bug for å "parkere" biler hvis de står på en enklet rute
        Minus = IIf(Car(a).Xdire = -1, 4, Minus)
        Minus = IIf(Car(a).Ydire = 1, 3, Minus)
        Minus = IIf(Car(a).Xdire = 1, 2, Minus)
        Minus = IIf(Car(a).Ydire = -1, 1, Minus)
        BitBlt Form1.picMain.hDC, Car(a).X * Size, Car(a).Y * Size, Size, Size, Form1.pCarm(((Car(a).Type * 4) - Minus)).hDC, 0, 0, SRCAND
        BitBlt Form1.picMain.hDC, Car(a).X * Size, Car(a).Y * Size, Size, Size, Form1.pCar(((Car(a).Type * 4) - Minus)).hDC, 0, 0, SRCPAINT
    Next a
    
End Sub
Sub SetFirstDire(N)
Dim Tell As Integer
Dim X, Y As Integer
Dim Ar() As Integer
    X = Car(N).X
    Y = Car(N).Y
    
    If Nett(X - 1, Y) = 1 Then Tell = Tell + 1
    If Nett(X, Y + 1) = 1 Then Tell = Tell + 1
    If Nett(X + 1, Y) = 1 Then Tell = Tell + 1
    If Nett(X, Y - 1) = 1 Then Tell = Tell + 1
    
    If Tell = 0 Then Exit Sub
    
    ReDim Ar(1 To Tell)
    i = 1
    If Nett(X - 1, Y) = 1 Then Ar(i) = 1: i = i + 1
    If Nett(X, Y + 1) = 1 Then Ar(i) = 2: i = i + 1
    If Nett(X + 1, Y) = 1 Then Ar(i) = 3: i = i + 1
    If Nett(X, Y - 1) = 1 Then Ar(i) = 4: i = i + 1
    
    
    Valg = Ar(rndTall(1, UBound(Ar)))
    Select Case Valg
    Case 1: Car(N).Xdire = -1
    Case 2: Car(N).Ydire = 1
    Case 3: Car(N).Xdire = 1
    Case 4: Car(N).Ydire = -1
    End Select
    
End Sub

Public Sub MoveCars()
Dim N As Integer
    For N = 1 To UBound(Car)
        With Car(N)

            SetNewDire N
            .LastPos.X = .X
            .LastPos.Y = .Y
            .X = .X + .Xdire
            .Y = .Y + .Ydire
        End With
    Next N
End Sub

Sub SetNewDire(N)
Dim Tell As Integer
Dim X, Y As Integer
Dim Ar() As Integer
    X = Car(N).X
    Y = Car(N).Y
    
    If Nett(X - 1, Y) = 1 Then Tell = Tell + 1
    If Nett(X, Y + 1) = 1 Then Tell = Tell + 1
    If Nett(X + 1, Y) = 1 Then Tell = Tell + 1
    If Nett(X, Y - 1) = 1 Then Tell = Tell + 1
    
    If Tell = 2 Then
        If Nett(X + Car(N).Xdire, Y + Car(N).Ydire) = 1 And Nett(X - Car(N).Xdire, Y - Car(N).Ydire) = 1 Then
            Exit Sub
        End If
    End If
    If Tell = 1 Then
        Car(N).Xdire = Car(N).Xdire * -1
        Car(N).Ydire = Car(N).Ydire * -1
        Exit Sub
    End If
    If Tell = 0 Then
        Car(N).Xdire = 0
        Car(N).Ydire = 0
        Exit Sub
    End If
    
    Car(N).Xdire = 0
    Car(N).Ydire = 0
    
    Tell = 0
    If Nett(X - 1, Y) = 1 And Not X - 1 = Car(N).LastPos.X Then Tell = Tell + 1
    If Nett(X, Y + 1) = 1 And Not Y + 1 = Car(N).LastPos.Y Then Tell = Tell + 1
    If Nett(X + 1, Y) = 1 And Not X + 1 = Car(N).LastPos.X Then Tell = Tell + 1
    If Nett(X, Y - 1) = 1 And Not Y - 1 = Car(N).LastPos.Y Then Tell = Tell + 1
    
    ReDim Ar(1 To Tell)
    i = 1
    If Nett(X - 1, Y) = 1 And Not X - 1 = Car(N).LastPos.X Then Ar(i) = 1: i = i + 1
    If Nett(X, Y + 1) = 1 And Not Y + 1 = Car(N).LastPos.Y Then Ar(i) = 2: i = i + 1
    If Nett(X + 1, Y) = 1 And Not X + 1 = Car(N).LastPos.X Then Ar(i) = 3: i = i + 1
    If Nett(X, Y - 1) = 1 And Not Y - 1 = Car(N).LastPos.Y Then Ar(i) = 4
    
    Dim Valg As Single
    Valg = Ar(rndTall(1, Tell))
    Select Case Valg
    Case 1: Car(N).Xdire = -1
    Case 2: Car(N).Ydire = 1
    Case 3: Car(N).Xdire = 1
    Case 4: Car(N).Ydire = -1
    End Select
    
    
End Sub
Public Function rndTall(min, max)
    Randomize
    rndTall = Int((Rnd * (max - min + 1)) + min)
End Function
