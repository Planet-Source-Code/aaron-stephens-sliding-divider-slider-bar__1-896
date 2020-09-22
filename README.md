<div align="center">

## Sliding Divider \(slider bar\)


</div>

### Description

Resizes two text boxes AS a divider is dragged left or right. Maintains full bounds checking. The methods used can be applied to other controls as well. This is a form of splitter bar.
 
### More Info
 
This is a very simple form with only two text boxes and a picture box. All the code is fully documented and fully adaptable.

----

Create a new form. The name is irrelevant to this code.

Place two text boxes and a picture box on the form. Name one text box "TextLeft" and the other "TextRight". Name the picture box "SlidingDivider".

All other attributes should be left as-is. Location and dimenstion of the controls are irrelevant.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aaron Stephens](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aaron-stephens.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aaron-stephens-sliding-divider-slider-bar__1-896/archive/master.zip)

### API Declarations

None


### Source Code

```
'General declarations section
'Sliding Divider between two controls.
'Written by: Aaron Stephens
'      Midnight Hour Enterprises, 1998.05.21
'This code may be freely distributed and may be
'altered in any way shape and form, if the author's
'name is removed.
'
'If this code is used in it's un-altered form,
'please give me some credit. Thanks.
'Flag for to tell MouseMove wether the sliding divider
'has been clicked.
Dim SDActive As Boolean
'Define the minimum with of the right and left
'controls.
Const MinRightWidth = 0
Const MinLeftWidth = 0
'End general declarations section
Private Sub Form_Load()
  'Set the text boxes and sliding divider to their
  'default parameters. In an adaptation, these
  'options could be loaded at startup, having been
  'saved at the last shutdown.
  'In addition, and controls (tool or status bars)
  'at the top or bottom of the form would need to
  'be compensated for. It would be preferable to
  'use a variable containing the offsets they
  'produce, instead of hard-coding the values
  'into every occurance in this form.
  TextLeft.Top = 0
  TextLeft.Left = 0
  TextLeft.Width = Me.ScaleWidth * 0.25
  TextLeft.Height = Me.ScaleHeight
  SlidingDivider.Top = 0
  SlidingDivider.Left = TextLeft.Width
  SlidingDivider.Width = 30
  SlidingDivider.Height = TextLeft.Height
  TextRight.Top = 0
  TextRight.Left = TextLeft.Width + SlidingDivider.Width
  TextRight.Width = Me.ScaleWidth - TextLeft.Width - SlidingDivider.Width
  TextRight.Height = TextLeft.Height
End Sub
Private Sub Form_Resize()
  'This resizes all controls on the form when the
  'form itself is resized.
  'Set the sliding divider to be at the same relative
  'position in the new form size.
  SlidingDivider.Left = CInt(Me.ScaleWidth * (SlidingDivider.Left / (TextLeft.Width + SlidingDivider.Width + TextRight.Width)))
  'Set the left text box's height.
  TextLeft.Height = Me.ScaleHeight
  'Set the left text box's width.
  TextLeft.Width = SlidingDivider.Left
  'Set the sliding divider and the right text box
  'height to the the same height as the left.
  SlidingDivider.Height = TextLeft.Height
  TextRight.Height = TextLeft.Height
  'Set the right text box to fill the remainder
  'of the form.
  TextRight.Left = TextLeft.Width + SlidingDivider.Width
  TextRight.Width = Me.ScaleWidth - TextLeft.Width - SlidingDivider.Width
End Sub
Private Sub SlidingDivider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'This sets a variable to tell the MouseMove routine
  'that the user has clicked the sliding divider.
  If Button = vbLeftButton Then
    SDActive = True
  End If
End Sub
Private Sub SlidingDivider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'This sets the sliding divider position to the mouse
  'position. I does check to make sure the sliding
  'divider and the objects that adjust to it do not
  'exceed the legal bounds of the form.
  'If the divider is clicked and the mouse has moved...
  If SDActive = True And CLng(X) <> SlidingDivider.Left Then
    'Set the DividerPosition
    SlidingDivider.Left = SlidingDivider.Left + (X - (SlidingDivider.Width / 2))
    'Check the bounds of the divider position and
    'correct if nesecary.
    If SlidingDivider.Left < MinLeftWidth Then SlidingDivider.Left = MinLeftWidth
    If SlidingDivider.Left + SlidingDivider.Width + MinRightWidth >= Me.ScaleWidth Then SlidingDivider.Left = Me.ScaleWidth - SlidingDivider.Width - MinRightWidth
    'Resize the text boxes.
    TextLeft.Width = SlidingDivider.Left
    TextRight.Left = TextLeft.Width + SlidingDivider.Width
    TextRight.Width = Me.ScaleWidth - TextLeft.Width - SlidingDivider.Width
  End If
End Sub
Private Sub SlidingDivider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'This calls the MouseMove routine to set the final
  'sliding divider position the sets a variable to
  'tell the MouseMove routine that the sliding
  'divider is no longer clicked.
  SlidingDivider_MouseMove Button, Shift, X, Y
  SDActive = False
End Sub
```

