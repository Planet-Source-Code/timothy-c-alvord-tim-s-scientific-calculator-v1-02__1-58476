Attribute VB_Name = "Gradient"
' GradientForm( Form, vbBlack, RGB(0,0,255), [bHorizontal=False]
'
Public Sub GradientForm(ByVal frmIn As Form, oleFromColor As OLE_COLOR, oleToColor As OLE_COLOR, Optional bHorizontal As Boolean = False)
'****************************************************************************
'   GradientForm( frmIn, oleFromColor, oleToColor, [bHorizontal=False] )
'       Draws a Gradient background either horizontally or vertically on the
'       form
'   Inputs:
'       frmIn        - Form to apply gradient to
'       oleFromColor - OLE Color to start gradient from
'       oleToColor   - OLE Color to end gradient on
'       bHorizontal  - Optional Horizontal Gradient [defaults to False]
'   Outputs:
'       None
'****************************************************************************
    Dim VR, VG, VB As Single
    Dim R, G, B, R2, G2, B2 As Integer
    Dim temp As Long
    
    frmIn.AutoRedraw = True
    ' Extract the RGB values from oleFromColor and oleToColor
    temp = (oleFromColor And 255)
    R = temp And 255
    temp = Int(oleFromColor / 256)
    G = temp And 255
    temp = Int(oleFromColor / 65536)
    B = temp And 255
    temp = (oleToColor And 255)
    R2 = temp And 255
    temp = Int(oleToColor / 256)
    G2 = temp And 255
    temp = Int(oleToColor / 65536)
    B2 = temp And 255
    
    If bHorizontal Then
        'create a calculation variable for determining the step between
        'each level of the gradient; this also allows the user to create
        'a perfect gradient regardless of the form size
        VR = Abs(R - R2) / frmIn.ScaleWidth
        VG = Abs(G - G2) / frmIn.ScaleWidth
        VB = Abs(B - B2) / frmIn.ScaleWidth
        'if the second value is lower then the first value, make the step
        'negative
        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < B Then VB = -VB
        'run a loop through the form width, incrementing the gradient color
        'according to the width of the line being drawn
        For x = 0 To frmIn.ScaleWidth
            R2 = R + VR * x
            G2 = G + VG * x
            B2 = B + VB * x
            frmIn.Line (x, 0)-(x, frmIn.ScaleHeight), RGB(R2, G2, B2)
        Next x
    Else
        'create a calculation variable for determining the step between
        'each level of the gradient; this also allows the user to create
        'a perfect gradient regardless of the form size
        VR = Abs(R - R2) / frmIn.ScaleHeight
        VG = Abs(G - G2) / frmIn.ScaleHeight
        VB = Abs(B - B2) / frmIn.ScaleHeight
        'if the second value is lower then the first value, make the step
        'negative
        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < B Then VB = -VB
        'run a loop through the form height, incrementing the gradient color
        'according to the height of the line being drawn
        For y = 0 To frmIn.ScaleHeight
            R2 = R + VR * y
            G2 = G + VG * y
            B2 = B + VB * y
            
            frmIn.Line (0, y)-(frmIn.ScaleWidth, y), RGB(R2, G2, B2)
        Next y
    End If
End Sub
