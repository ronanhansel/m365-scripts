Attribute VB_Name = "Module1"
Option Explicit
Sub ReplaceColours()
Dim oSld As
    Slide
Dim oShp As Shape
Dim I As Integer
For Each oSld In
    ActivePresentation.Slides
    For Each oShp In oSld.Shapes
        
    If oShp.Type = msoGroup Then
            
    For I = 1 To oShp.GroupItems.Count
                
    Call FindAndReColourText(oShp.GroupItems(I), _

    RGB(100, 100, 100), RGB(255, 0, 255))
                
    Call FindAndReColourFill(oShp.GroupItems(I), _

    RGB(255, 255, 255), RGB(255, 0, 255))
                
    Call FindAndReColourBorder(oShp.GroupItems(I), _

    RGB(100, 100, 100), RGB(255, 0, 255))
            
    Next I
        Else
            
    Call FindAndReColourText(oShp, _

    RGB(100, 100, 100), RGB(255, 0, 255))
            
    Call FindAndReColourFill(oShp, _

    RGB(255, 255, 255), RGB(255, 0, 255))
            
    Call FindAndReColourBorder(oShp, _

    RGB(100, 100, 100), RGB(255, 0, 255))
        
    End If
    Next oShp
Next oSld
End Sub


    Function FindAndReColourText(oShp As Shape, _

    oRGB As Long, oNewRGB As Long)
Dim I As Integer
Dim oTxtRng As
    TextRange
On Error Resume Next
If oShp.HasTextFrame Then
    
    If oShp.TextFrame.HasText Then
        
    Set oTxtRng = oShp.TextFrame.TextRange
        
    For I = 1 To oTxtRng.Runs.Count
            
    With oTxtRng.Runs(I).Font.Color
                
    If .Type = msoColorTypeRGB Then
                    
    If .RGB = oRGB Then
                        
    .RGB = oNewRGB
                    
    End If
                
    End If
            
    End With
        Next I
    
    End If
End If
End Sub
Function FindAndReColourFill(oShp As
    Shape, _

    oRGB As Long, oNewRGB As Long)
On Error Resume Next
If
    oShp.Fill.Visible Then
    If oShp.Fill.ForeColor.RGB =
    oRGB Then
        
    oShp.Fill.ForeColor.RGB = oNewRGB
    End If
End If

    End Sub
Function FindAndReColourBorder(oShp As Shape, _

    oRGB As Long, oNewRGB As Long)
On Error Resume Next
If
    oShp.Line.Visible Then
    If oShp.Line.ForeColor.RGB =
    oRGB Then
        
    oShp.Line.ForeColor.RGB = oNewRGB
    End If
End If

    End Sub

