Attribute VB_Name = "CmdButtons"
'**********      VBButton Control v1.0 for VB 3.0-6.0    **********
'*******                                                     ******
'****         Property of WolfeByte Solutions 1995-2002        ****
'**                                                              **
'**      This program is protected by and subject to all         **
'**    Federal copyright laws governing the duplication and      **
'**  distribution of authored software.  With the purchase and   **
'** use of this program you agree to release WolfeByte Solutions **
'**  of all liability and/or damages as related to the use of    **
'**    this program and also acknowledge that no claims or       **
'**      warranties regarding its usage have been offered.       **
'****                                                          ****
'*******                                                     ******
'**********        Update Version 1.1  June 7, 1996      **********

Option Explicit

'All subs are in this module.  To create a new control - create a new
'picturebox and size it to the desired button size.  Then copy the code
'from the demo applications following events and paste them into the same
'events in your project: 1) Form_Load  2) Picture1_DblClick,
'...GotFocus, ...LostFocus, ...MouseMove, ...MouseUp and ...MouseDown. Rename any
'picture1 references in this code to the names of your new controls.

' Font Type Structure
Public Type fontType
  sName As String
  iSize As Integer
  bBold As Boolean
  bItalic As Boolean
  iUnderline As Integer
  lColor  As Long
End Type

' pbCommandButton Structure
Public Type ECommandButton
    State As Integer
    Bevel As Integer
    Font As fontType
    Text As String
    VAlign As String
    HAlign As String
    bMultiLine As Boolean
    bFocus As Boolean
End Type

Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Sub DrawButtonText(pbIn As PictureBox, rBox As RECT, _
                   sText As String, sHorzAlign As String, sVertAlign As String, _
                   bMultiLine As Boolean)

'Called from within this module to draw text on picturebox and align accordingly
'within area passed in.  The vertical alignment of the text is a percent of
'distance from the top of the area - pass in one of these -> 0 = top,
'.5 = middle, 1 = bottom.

'The Multi-Line will split on words that have a space between them as well
'   as split up long words.  If the word is too long to fit horizontal on the
'   button it will be split between two or more lines.  It also implements
'   a char by char compare instead of using instr for the word break so that
'   individual chars can be filtered out and acted upon specifically
'   i.e. as with the CRLFChr.  The CRLFChr "~" acts as an immediate break
'   at the word and starts the new text on the next line.
'   This can be modified for the '&' character as well.
   
    Dim iBegin As Integer      'Beginning position
    Dim iEnd As Integer        'Ending position
    Dim iBreak As Integer      'Where to break the line
    Dim iTextLen As Integer    'Length of text to be output
    Dim iPctTOLen As Integer   'Text Output Length
    Dim sTemp As String       'Work String
    Dim sChar As String        'Character working with
    Dim iElement As Integer    'Element working with
    Dim arText()  As String     'Array of text
    Dim iTextHeight As Integer 'Height of all text lines
    Dim lbVertAlign As Single   'Vertical Alignment of Text
    Const CRLFChr = "~"         'Character to act as CRLF
    
     sTemp = ""
    'Set the vertical alignment - Default is "center"
    If LCase(sVertAlign) = "top" Then
        lbVertAlign = 0
    ElseIf LCase(sVertAlign) = "bottom" Then
       lbVertAlign = 1
    Else
       lbVertAlign = 0.5
    End If
    
    iPctTOLen = rBox.Right - rBox.Left  'Text Output Length
    iTextLen = pbIn.TextWidth(sText) 'Length of text
    
    'This checks to see if the line needs to be split
    '   Always split if it is a multi-line
    If bMultiLine Or ((iTextLen > iPctTOLen) And bMultiLine) Then
        iBegin = 1        'Begin with first position
        iElement = 1      'We know we have at least 1 element
        
        For iEnd = 1 To Len(sText) + 1  'Look at all chrs
            sChar = Mid$(sText, iEnd, 1)
            If sChar = " " Then
                iBreak = iEnd    'Keep Break on a space position
            End If
            
            'This is where the check for CRLFChr is
            If (pbIn.TextWidth(sTemp) + pbIn.TextWidth(sChar)) < iPctTOLen And sChar <> "" And sChar <> CRLFChr Then
                sTemp = sTemp & sChar
            Else
                ReDim Preserve arText(iElement)
                If sChar = "" Then  'End of text
                    iBreak = iEnd
                End If
                If sChar = CRLFChr Then  'CRLF character
                    iBreak = iEnd
                End If
                If iBreak = 0 Then  'Break on words
                    iBreak = iEnd - 1
                    arText(iElement) = Mid$(sTemp, 1, iEnd - iBegin)
                Else
                   arText(iElement) = Mid$(sTemp, 1, iBreak - iBegin)
                End If
                
                'Text height of each line for centering
                iTextHeight = iTextHeight + pbIn.TextHeight(arText(iElement))
                
                'End of Text ?
                If sChar <> "" Then
                    sTemp = ""
                    iBegin = iBreak + 1
                    iEnd = iBreak
                    iBreak = 0
                    iElement = iElement + 1
                End If
            End If
        Next
    Else
        'We still need to move the text to be processed
        ReDim arText(1)   'There is only 1 element
        arText(1) = sText 'Move the text in
        iTextHeight = pbIn.TextHeight(arText(1)) 'Get the heigth of text
    End If
    
    'Calculate the Y position
    pbIn.CurrentY = rBox.Top + (((rBox.Bottom - rBox.Top) * lbVertAlign) - (iTextHeight * lbVertAlign))
    
    'Don't let it go over the top
    If pbIn.CurrentY < rBox.Top Then
        pbIn.CurrentY = rBox.Top
    End If
    
    For iElement = 1 To UBound(arText) 'Loop thru all element strings
       Select Case LCase(sHorzAlign)
          Case "left"
             pbIn.CurrentX = rBox.Left
          Case "right"
             pbIn.CurrentX = rBox.Right - pbIn.TextWidth(arText(iElement))
          Case "center"
             pbIn.CurrentX = rBox.Left + (((rBox.Right - rBox.Left) / 2) - (pbIn.TextWidth(arText(iElement)) / 2))
          Case Else  'Default "center"
             pbIn.CurrentX = rBox.Left + (((rBox.Right - rBox.Left) / 2) - (pbIn.TextWidth(arText(iElement)) / 2))
       End Select
       pbIn.Print arText(iElement)
    Next
    
    ReDim arText(0)
End Sub

Public Sub DrawCommandButton(ecbInCmd As ECommandButton, pbIn As PictureBox)
    'Called each time the button changes focus or up/down
    Dim x%, y%
    Dim rBox As RECT

    pbIn.Cls                  'Used to clear text (not need if not using picture)
    Select Case ecbInCmd.bFocus 'Draw button according to focus state
        Case True
            pbIn.FillStyle = 1
            pbIn.Line (0, 0)-(pbIn.ScaleWidth - 1, pbIn.ScaleHeight - 1), 0, B
            'Focus rectangle in these two loops
            For x = ecbInCmd.Bevel + 3 To pbIn.ScaleWidth - ecbInCmd.Bevel - 3 Step 2
                pbIn.PSet (x, ecbInCmd.Bevel + 2), 0
                pbIn.PSet (x, pbIn.ScaleHeight - ecbInCmd.Bevel - 3), 0
            Next
            For y = ecbInCmd.Bevel + 3 To pbIn.ScaleHeight - ecbInCmd.Bevel - 3 Step 2
                pbIn.PSet (ecbInCmd.Bevel + 2, y), 0
                pbIn.PSet (pbIn.ScaleWidth - ecbInCmd.Bevel - 3, y), 0
            Next
        Case False
            pbIn.Line (0, pbIn.ScaleHeight - 1)-(pbIn.ScaleWidth, pbIn.ScaleHeight - 1), 0
            pbIn.Line (pbIn.ScaleWidth - 1, 0)-(pbIn.ScaleWidth - 1, pbIn.ScaleHeight), 0
    End Select
    
    SetFonts ecbInCmd.Font, pbIn 'Prepare font settings
    
    Select Case ecbInCmd.State
        Case 0                     'Button disabled
            For x = 0 To ecbInCmd.Bevel - 1
                pbIn.Line (1 + x, pbIn.ScaleHeight - 2 - x)-(pbIn.ScaleWidth - 2 - x + 1, pbIn.ScaleHeight - 2 - x), RGB(92, 92, 92)
                pbIn.Line (1 + x, 1 + x)-(pbIn.ScaleWidth - 2 - x + 1, 1 + x), RGB(255, 255, 255)
                pbIn.Line (pbIn.ScaleWidth - 2 - x, 1 + x)-(pbIn.ScaleWidth - 2 - x, pbIn.ScaleHeight - 2 - x), RGB(92, 92, 92)
                pbIn.Line (1 + x, 1 + x)-(1 + x, pbIn.ScaleHeight - 2 - x), RGB(255, 255, 255)
            Next
            rBox.Left = ecbInCmd.Bevel + 3
            rBox.Top = ecbInCmd.Bevel + 1
            rBox.Right = pbIn.ScaleWidth - ecbInCmd.Bevel - 3
            rBox.Bottom = pbIn.ScaleHeight - ecbInCmd.Bevel - 3
            pbIn.ForeColor = QBColor(8)
            DrawButtonText pbIn, rBox, ecbInCmd.Text, ecbInCmd.HAlign, ecbInCmd.VAlign, ecbInCmd.bMultiLine
        Case 1                     'Button up
            For x = 0 To ecbInCmd.Bevel - 1
                pbIn.Line (1 + x, pbIn.ScaleHeight - 2 - x)-(pbIn.ScaleWidth - 2 - x + 1, pbIn.ScaleHeight - 2 - x), RGB(92, 92, 92)
                pbIn.Line (1 + x, 1 + x)-(pbIn.ScaleWidth - 2 - x + 1, 1 + x), RGB(255, 255, 255)
                pbIn.Line (pbIn.ScaleWidth - 2 - x, 1 + x)-(pbIn.ScaleWidth - 2 - x, pbIn.ScaleHeight - 2 - x), RGB(92, 92, 92)
                pbIn.Line (1 + x, 1 + x)-(1 + x, pbIn.ScaleHeight - 2 - x), RGB(255, 255, 255)
            Next
            rBox.Left = ecbInCmd.Bevel + 3
            rBox.Top = ecbInCmd.Bevel + 1
            rBox.Right = pbIn.ScaleWidth - ecbInCmd.Bevel - 3
            rBox.Bottom = pbIn.ScaleHeight - ecbInCmd.Bevel - 3
            DrawButtonText pbIn, rBox, ecbInCmd.Text, ecbInCmd.HAlign, ecbInCmd.VAlign, ecbInCmd.bMultiLine
        Case 2                      'Button down
            rBox.Left = ecbInCmd.Bevel + 5
            rBox.Top = ecbInCmd.Bevel + 2
            rBox.Right = pbIn.ScaleWidth - ecbInCmd.Bevel - 1
            rBox.Bottom = pbIn.ScaleHeight - ecbInCmd.Bevel - 2
            DrawButtonText pbIn, rBox, ecbInCmd.Text, ecbInCmd.HAlign, ecbInCmd.VAlign, ecbInCmd.bMultiLine
    End Select
End Sub

Sub SetFonts(rFont As fontType, pbIn As PictureBox)
    'Called from within this module to set the fonts of the picturebox.  The font
    'name will not be checked for correct spelling (if errors - check name spelling first).
    pbIn.FontName = rFont.sName
    pbIn.FontSize = rFont.iSize
    pbIn.FontBold = rFont.bBold
    pbIn.FontItalic = rFont.bItalic
    pbIn.FontUnderline = rFont.iUnderline
    pbIn.ForeColor = rFont.lColor
End Sub
