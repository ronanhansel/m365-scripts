Attribute VB_Name = "NewMacros"
Sub Example1()
Dim lngPercent2Scale As Long
Dim lngOriginalHeight As Long
Dim lngScaledHeight As Long

'percent to resize
lngPercent2Scale = 70
'the height of the scaled image
lngScaledHeight = InlineShapes.Item(1).Height
'rescale to original size
InlineShapes.Item(1).ScaleHeight = 100
'the size of the original image
lngOriginalHeight = InlineShapes.Item(1).Height
'rescale image
InlineShapes.Item(1).ScaleHeight = _
lngScaledHeight / lngOriginalHeight * 100
'resize
InlineShapes.Item(1).ScaleHeight _
= lngPercent2Scale * lngScaledHeight / lngOriginalHeight
End Sub

