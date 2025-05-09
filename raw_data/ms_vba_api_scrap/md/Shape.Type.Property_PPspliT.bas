Attribute VB_Name = "PPspliT"
'
'
'    _____  _____           _ _ _______
'   |  __ \|  __ \         | (_)__   __|
'   | |__) | |__) |__ _ __ | |_   | |
'   |  ___/|  ___/ __| '_ \| | |  | |
'   | |    | |   \__ \ |_) | | |  | |
'   |_|    |_|   |___/ .__/|_|_|  |_|
'                    | |
'                    |_| by Massimo Rimondini - version 1.27
'
' first written by Massimo Rimondini in November 2009
' last update: April 2022
' Source code for PowerPoint 2003-
'
'





' This global variable indicates whether and how slide numbers should be kept
' consistent with the original set of slides. For example, if slide 6 is split
' into 3 slides, then all those 3 slides will be numbered 6 after splitting.
' As an alternative option, a subindex can be added to slide numbers, so that,
' for example, slide 6 is split into 6.1, 6.2, 6.3, etc.
Public slideNumbersAdjustMode As Integer
Public Const SLIDENUMBER_DONOTHING = 1
Public Const SLIDENUMBER_BAKE = 2
Public Const SLIDENUMBER_SUBINDEX = 3

' This global variable indicates whether animations should be split
' at each mouse-triggered event. If set to false, a separate slide is
' created for each and every animation.
Public doNotSplitMouseTriggered As Boolean

' The following variables are for internal use only.
Public cancelStatus As Boolean
Public slide_number As Integer

'
' Convert decimal separators in the argument string from '.' to the most
' appropriate character for the system-configured locale.
'
Private Function localizeDecimalSeparators(ByVal s As String)
    Dim d As Double, useCommaAsSeparator As Boolean
    useCommaAsSeparator = False
    
    ' Use a test value to check for the currently used decimal
    ' separator. In principle, we could use the user-supplied
    ' argument, but if it is a value between 0 and 1, it could
    ' miss the leading zero (e.g., -.1234), thus raising errors
    ' if we are not using the correct decimal separator in the
    ' assignment (which is exactly what we are trying to
    ' discover here).
    
    d = "1,2"
    ' If "," is not the decimal separator in use for the current
    ' system locale, this assignment results in losing the decimal
    ' separator.
    ' Now, this test requires care: in fact, localization of
    ' Double values seems to happen whenever a value is output on
    ' screen or is converted from a string, but in some way it does
    ' not seem to affect the internal representation of the Double
    ' value. Therefore, to check whether the decimal separator
    ' has survived the assignment, we need to look for its
    ' internal representation (which is "."), not its localized one.
    useCommaAsSeparator = (InStr(Trim(Str$(d)), ".") > 0)
    
    If useCommaAsSeparator Then
        d = Replace(s, ".", ",")
    Else
        d = s
    End If
    localizeDecimalSeparators = d
End Function

'
' Hide a paragraph in a text box.
' Arguments are the shape containing the text frame and the index of
' the paragraph to be hidden. The subroutine takes care of preserving
' the space occupied by the paragraph, so that a text frame with
' auto-fit enabled will still be rendered accurately.
'
Private Sub clearParagraph(sh As Shape, par)
    If sh.TextFrame.TextRange.Paragraphs(par).Lines.Count > 1 Then
        ' This is a word wrapped or multi-line paragraph: turn every
        ' word wrap into a real new line. This is required because the
        ' paragraph contents will be soon replaced with spaces, which
        ' have a different width than original characters, can therefore
        ' mess up word wrapping, hence the number of lines of this paragraph,
        ' hence the rendering of any following paragraphs.
        For i = 2 To sh.TextFrame.TextRange.Paragraphs(par).Lines.Count
            If Asc(sh.TextFrame.TextRange.Paragraphs(par).Lines(i - 1).Characters(sh.TextFrame.TextRange.Paragraphs(par).Lines(i - 1).Characters.Count)) <> 11 _
                And Asc(sh.TextFrame.TextRange.Paragraphs(par).Lines(i - 1).Characters(sh.TextFrame.TextRange.Paragraphs(par).Lines(i - 1).Characters.Count)) <> 13 Then
                sh.TextFrame.TextRange.Paragraphs(par).Lines(i).Characters(1).InsertBefore Chr$(11)
            End If
        Next i
    End If
    Set p = sh.TextFrame.TextRange.Paragraphs(par)
    i = 1
    While i <= p.Characters.Count
        ' Replace paragraph contents with spaces. This is the best and
        ' most compatible way I found to "hide" a paragraph while keeping
        ' its original space occupied.
        If Asc(p.Characters(i)) <> 13 And Asc(p.Characters(i)) <> 11 Then
            p.Characters(i) = " "
        End If
        i = i + 1
    Wend
    ' Set bullet symbol too to " " (32 is the Unicode value)
    p.ParagraphFormat.Bullet.Character = 32
End Sub

'
' Copies the contents of p2 into p1.
' This is used to restore a previously hidden paragraph.
'
Private Sub copyParagraph(p1 As TextRange, p2 As TextRange)
    Dim newLineInserted As Boolean

    ' Sometimes text paragraphs are just empty. In this case
    ' return immediately
    If p2.Characters.Count = 0 Then Exit Sub
    
    If Asc(p2.Characters(p2.Characters.Count)) <> 13 Then
        ' This paragraph does not end with a new line (most
        ' likely because it is the last paragraph in the text
        ' frame). Here I add it because I can get all the
        ' formatting attributes of a paragraph only if it
        ' ends with a new line (this is PowerPoint magic...).
        ' In addition, although supported, using
        ' p2.Characters.InsertAfter here with PowerPoint <= 2003
        ' has the adverse effect that the paragraph text
        ' (property p2.Text) as well as its length are not
        ' updated, causing subsequent text editing steps
        ' (including, e.g., removal of the inserted newline)
        ' to fail.
        p2.InsertAfter Chr$(13)
        newLineInserted = True
    End If
    
    ' Apply contents and formatting from the original paragraph
    p2.Copy
    
    ' It seems that the following 3 assignments, applied *before* pasting
    ' the paragraph, reduce the number of cases in which bullet symbols
    ' are lost. The reason why this happens is completely obscure to me, but
    ' repeating the assignment *after* pasting (where this should happen)
    ' seems to be harmless.
    p1.ParagraphFormat.SpaceAfter = p2.ParagraphFormat.SpaceAfter
    p1.ParagraphFormat.SpaceBefore = p2.ParagraphFormat.SpaceBefore
    p1.ParagraphFormat.SpaceWithin = p2.ParagraphFormat.SpaceWithin
    
    p1.Paste
    
    p1.IndentLevel = p2.IndentLevel
    p1.ParagraphFormat.SpaceAfter = p2.ParagraphFormat.SpaceAfter
    p1.ParagraphFormat.SpaceBefore = p2.ParagraphFormat.SpaceBefore
    ' Try hard to set inter-line spacing. Applying a small variation should
    ' force PowerPoint to honor the value.
    p1.ParagraphFormat.SpaceWithin = p2.ParagraphFormat.SpaceWithin - 0.01
    p1.ParagraphFormat.SpaceWithin = p2.ParagraphFormat.SpaceWithin + 0.01
    
    ' Restore bullet formatting. Since there seems to be no
    ' way to get the currently used image for a bullet, care
    ' must be taken in updating the bullet attributes only if
    ' required, otherwise the applied image may be messed up
    ' and I may be unable to restore it.
    If p1.ParagraphFormat.Bullet.Type <> p2.ParagraphFormat.Bullet.Type Then
        p1.ParagraphFormat.Bullet.Type = p2.ParagraphFormat.Bullet.Type
    End If
    If p2.ParagraphFormat.Bullet.Type = ppBulletUnnumbered And p1.ParagraphFormat.Bullet.Character <> p2.ParagraphFormat.Bullet.Character Then
        p1.ParagraphFormat.Bullet.Character = p2.ParagraphFormat.Bullet.Character
        ' Apparently, not all the font attributes of a bullet can be reset (assigning
        ' some of them triggers an error). So, here we reimplement the relevant part
        ' of copyFontAttributes
        With p1.ParagraphFormat.Bullet.Font
            .Name = p2.ParagraphFormat.Bullet.Font.Name
            .Size = p2.ParagraphFormat.Bullet.Font.Size
            assignColor .Color, p2.ParagraphFormat.Bullet.Font.Color
        End With
    End If

    If p2.ParagraphFormat.Bullet.Type = ppBulletNumbered And p1.ParagraphFormat.Bullet.StartValue <> p2.ParagraphFormat.Bullet.StartValue Then
        p1.ParagraphFormat.Bullet.StartValue = p2.ParagraphFormat.Bullet.StartValue
    End If
    If p2.ParagraphFormat.Bullet.Type = ppBulletNumbered And p1.ParagraphFormat.Bullet.Style <> p2.ParagraphFormat.Bullet.Style Then
        p1.ParagraphFormat.Bullet.Style = p2.ParagraphFormat.Bullet.Style
    End If
    ' It's not over yet.
    ' Paste often acts in an "intelligent" way, by cutting away
    ' apparently useless spaces and other stuff. Here I need a
    ' really accurate paste, which preserves all the characters,
    ' therefore I overwrite (or enrich) the set of previously
    ' pasted characters. Overwriting the characters one by one
    ' ensures that the rest of formatting is left untouched, but
    ' here I may still be adding new text (e.g., new spaces), to
    ' which formatting must be applied. This is the reason of the
    ' call to copyFontAttributes.
    For i = 1 To p2.Characters.Count
        ' It's better to explicitly handle the case for added characters
        ' here. Failure to do so has caused inconsistent text rendering
        ' in some cases.
        If i <= p1.Characters.Count Then
            p1.Characters(i) = p2.Characters(i)
        Else
            p1.InsertAfter p2.Characters(i)
        End If
        copyFontAttributes p1.Characters(i).Font, p2.Characters(i).Font
    Next i

    ' Remove any previously inserted new line characters
    If newLineInserted Then
        p1.Characters(p1.Characters.Count).Delete
        p2.Characters(p2.Characters.Count).Delete
    End If
End Sub

'
' Copies fundamental font attributes from f2 to f1.
'
Private Sub copyFontAttributes(f1 As Font, f2 As Font)
    f1.Name = f2.Name
    f1.Size = f2.Size
    f1.Bold = f2.Bold
    f1.Italic = f2.Italic
    f1.Underline = f2.Underline
    ' Warning: assigning just one between the Subscript and the Superscript
    ' attributes, even to the msoFalse value, may impact the other. Therefore
    ' these attributes must be assigned only when strictly required.
    If f2.Subscript Then f1.Subscript = msoTrue
    If f2.Superscript Then f1.Superscript = msoTrue
    If Not f2.Subscript And Not f2.Superscript Then
        f1.Subscript = msoFalse
        f1.Superscript = msoFalse
    End If
    assignColor f1.Color, f2.Color
End Sub

'
' This subroutine applies the ZOrder (depth) of shapes in s2 to shapes in s1.
' Corresponding shapes in s1 and in s2 are different objects, therefore, in order
' to be matched, shape IDs must have been copied in advance to a shape property
' that is more persistent by using the copyShapeIds subroutine.
' Note: the algorithm used to sort shapes in s2 by increasing ZOrder could be
' improved.
'
Private Sub matchZOrder(s1 As Slide, s2 As Slide)
    Dim sortedShapes(255) As Shape
    ProgressForm.infoLabel = "Matching shape Z order..."
    ProgressForm.Repaint
    zThreshold = 0
    j = 1
    For i = 1 To s2.Shapes.Count
        minZ = 65536
        ' Find shape in s2 with minimum ZOrder greater than zThreshold
        For Each sh2 In s2.Shapes
            ' Inequalities are strict because there should be no
            ' two shapes with the same ZOrder
            If sh2.ZOrderPosition < minZ And sh2.ZOrderPosition > zThreshold Then
                minZ = sh2.ZOrderPosition
                minZshapeId = sh2.Tags("shapeId")
            End If
        Next sh2
        zThreshold = minZ
        shapeIdInS1 = findShape(s1, minZshapeId)
        If shapeIdInS1 > 0 Then
            ' The same shape exists also in s1: add the shape to the array of sorted shapes
            Set sortedShapes(j) = s1.Shapes(shapeIdInS1)
            j = j + 1
        End If
    Next i
    
    ' Bring to front shapes in s1 by increasing values of ZOrder
    For i = 1 To j - 1
        sortedShapes(i).ZOrder msoBringToFront
    Next i
    ProgressForm.infoLabel = ""
    ProgressForm.Repaint
End Sub


'
' This subroutine deletes a shape from a slide. If the shape is a textbox
' and its paragraphs are animated independently from each other, then only
' the affected paragraph will be deleted. It takes as input the affected
' shape, a timeline and the index of the effect to be removed from the timeline.
' The returned value is true if and only if the function also deleted the
' effect (besides the shape or paragraph).
'
Private Function deleteShape(sh As Shape, theTimeline As Sequence, effectId)
    theParagraph = getEffectParagraph(theTimeline(effectId))
    If theParagraph > 0 Then
        ' This appears to be a text paragraph effect
        oldCount = theTimeline.Count
        If oldCount > effectId Then
            ' There are other effects following this one.
            ' Save the trigger type of the next effect for restoring it later
            animType = theTimeline(effectId + 1).Timing.TriggerType
        End If
        ' Delete (or better, hide) the paragraph
        clearParagraph sh, theParagraph
        If theTimeline.Count < oldCount Then
            ' The removed paragraph was not the last one in the shape, and therefore
            ' the effect has been automatically removed. Restore the trigger
            ' type if required
            If theTimeline.Count >= effectId Then
                ' Restore the trigger type
                theTimeline(effectId).Timing.TriggerType = animType
            End If
            deleteShape = True
        Else
            ' The removed paragraph was the last one in the shape, therefore
            ' the effect is still there.
            deleteShape = False
        End If
    Else
        ' Whole shape effect
        sh.Delete
        deleteShape = True
    End If
End Function

'
' This subroutine assigns the color in the ColorFormat object
' col2 to the ColorFormat object col1.
' Care must be taken in that the color may be specified as an
' index referring to the slide color scheme or as an RGB value.
'
Private Sub assignColor(col1 As ColorFormat, col2 As ColorFormat)
    If col2.Type <> msoColorTypeRGB Then
        ' I must protect from invalid assignments of color
        ' scheme indexes.
        On Error Resume Next
        col1.SchemeColor = col2.SchemeColor
        ' The brightness attribute does not seem to be accessible
        ' in PowerPoint releases prior to 2010, so we are not setting
        ' it here.
        On Error GoTo 0
    Else
        col1.RGB = col2.RGB
    End If
End Sub

'
' This subroutine converts a color value from the RGB space to the
' HSL space. The result will be put in the last 3 arguments.
' The procedure is taken from http://en.wikipedia.org/wiki/HSL_and_HSV#Conversion_from_RGB_to_HSL_overview
'
Private Sub RGBtoHSL(r, g, b, h, s, l)
    max = 0: min = 255
    r = r / 255: g = g / 255: b = b / 255
    If r > max Then max = r
    If g > max Then max = g
    If b > max Then max = b
    If r < min Then min = r
    If g < min Then min = g
    If b < min Then min = b
    If max = min Then
        h = 0
    ElseIf max = r Then
        h = (60 * (g - b) / (max - min) + 360) Mod 360
    ElseIf max = g Then
        h = 60 * (b - r) / (max - min) + 120
    ElseIf max = b Then
        h = 60 * (r - g) / (max - min) + 240
    End If
    l = (max + min) / 2
    If max = min Then
        s = 0
    ElseIf l <= 1 / 2 Then
        s = (max - min) / (2 * l)
    ElseIf l > 1 / 2 Then
        s = (max - min) / (2 - 2 * l)
    End If
End Sub

'
' This subroutine converts a color value from the HSL space to the
' RGB space. The result will be put in the last 3 arguments.
' The procedure is taken from http://en.wikipedia.org/wiki/HSL_and_HSV#Conversion_from_RGB_to_HSL_overview
'
Private Sub HSLtoRGB(h, s, l, r, g, b)
    If l < 1 / 2 Then
        q = l * (1 + s)
    Else
        q = l + s - l * s
    End If
    p = 2 * l - q
    hk = h / 360
    tr = hk + 1 / 3
    ' Cannot use the Mod operator here, as it only supports integer arithmetic
    If tr < 0 Then tr = tr + 1
    If tr > 1 Then tr = tr - 1
    tg = hk
    If tg < 0 Then tg = tg + 1
    If tg > 1 Then tg = tg - 1
    tb = hk - 1 / 3
    If tb < 0 Then tb = tb + 1
    If tb > 1 Then tb = tb - 1

    If tr < 1 / 6 Then
        r = p + ((q - p) * 6 * tr)
    ElseIf tr >= 1 / 6 And tr < 1 / 2 Then
        r = q
    ElseIf tr >= 1 / 2 And tr < 2 / 3 Then
        r = p + ((q - p) * 6 * (2 / 3 - tr))
    Else
        r = p
    End If
    If tg < 1 / 6 Then
        g = p + ((q - p) * 6 * tg)
    ElseIf tg >= 1 / 6 And tg < 1 / 2 Then
        g = q
    ElseIf tg >= 1 / 2 And tg < 2 / 3 Then
        g = p + ((q - p) * 6 * (2 / 3 - tg))
    Else
        g = p
    End If
    If tb < 1 / 6 Then
        b = p + ((q - p) * 6 * tb)
    ElseIf tb >= 1 / 6 And tb < 1 / 2 Then
        b = q
    ElseIf tb >= 1 / 2 And tb < 2 / 3 Then
        b = p + ((q - p) * 6 * (2 / 3 - tb))
    Else
        b = p
    End If
    r = r * 255: g = g * 255: b = b * 255
End Sub

'
' This subroutine converts a color value represented by VBA as a Long
' integer into its RGB components. The result is put in the last
' 3 arguments of the subroutine.
'
Private Sub colToRGB(col, r, g, b)
    r = col Mod 256
    g = (col \ 256) Mod 256
    b = (col \ 256 \ 256) Mod 256
End Sub

'
' This subroutine "rotates" the hue of a given color of the
' specified angle (in degrees).
'
Private Sub rotateColor(col As ColorFormat, rot)
    colToRGB col.RGB, r, g, b
    RGBtoHSL r, g, b, h, s, l
    h = (h + rot) Mod 360
    HSLtoRGB h, s, l, r, g, b
    col.RGB = RGB(r, g, b)
End Sub

'
' This subroutine alters the lightness of a given color.
' The amount should be between 0 and 1.
'
Private Sub changeLightness(col As ColorFormat, amount)
    colToRGB col.RGB, r, g, b
    RGBtoHSL r, g, b, h, s, l
    l = l + amount
    If l > 1 Then l = 1
    If l < 0 Then l = 0
    HSLtoRGB h, s, l, r, g, b
    col.RGB = RGB(r, g, b)
End Sub

'
' After a motion effect has been applied to a shape, the coordinates
' of all subsequent motion effects have been moved together with the
' shape. This subroutine applies a given shift to the arrival
' coordinates (indeed, arrival coordinates is all I need to update)
' of all the other motion effects for the same shape. Arguments
' effectSequence (the sequence of effects applied to the shape) and
' sh (the affected shape) do not need, and in general do not, refer
' to the same slide.
'
' A motion path is specified in VML. Information about the specification
' can be found here: http://www.w3.org/TR/NOTE-VML#_Toc416858391
'
Private Sub shiftAllMotions(effectSequence As Sequence, sh As Shape, shiftX, shiftY)
    Dim currentEffect As Effect, lastX As Double, lastY As Double
    For Each currentEffect In effectSequence
        ' The following variable is where I will put the reconstructed
        ' path with updated arrival coordinates
        motionPathString$ = ""
        ' Keep in mind that sh is a shape the effect is applied to (therefore
        ' it comes from a certain slide), while effectSequence is the sequence of effects
        ' under consideration (which comes from a different slide). Therefore,
        ' operator "Is" cannot be used to match the shapes whose motion effects
        ' should be updated.
        If isPathEffect(currentEffect) And currentEffect.Shape.Tags("shapeId") = sh.Tags("shapeId") Then
            ' This is a motion effect applied to the shape under consideration
            motionPathTokens = Split(currentEffect.Behaviors(1).MotionEffect.Path)
            ' The first character states this is a path motion, therefore I preserve it
            motionPathString$ = motionPathString$ + Trim(motionPathTokens(0)) + " "
            If currentEffect.Behaviors(1).Timing.Speed < 0 Then
                ' The path has been reversed: update origin coordinates instead
                lastX = localizeDecimalSeparators(motionPathTokens(1))
                lastY = localizeDecimalSeparators(motionPathTokens(2))
                lastX = lastX + shiftX
                lastY = lastY + shiftY
                motionPathString$ = motionPathString$ + Trim(Str$(lastX)) + " " + Trim(Str$(lastY)) + " "
                ' Append the rest of the motion string
                For i = 3 To UBound(motionPathTokens)
                    motionPathString$ = motionPathString$ + motionPathTokens(i) + " "
                Next i
            Else
                ' Update the last two (i.e., arrival) coordinates
                getLastCoordinates currentEffect.Behaviors(1).MotionEffect.Path, lastX, lastY, lastToken
                lastX = lastX + shiftX
                lastY = lastY + shiftY
                ' Copy everything but the last two coordinates from the original
                ' motion string
                For i = 0 To lastToken
                    motionPathString$ = motionPathString$ + motionPathTokens(i) + " "
                Next i
                ' Append the modified coordinates
                motionPathString$ = motionPathString$ + Trim(Str$(lastX)) + " " + Trim(Str$(lastY)) + " "
            End If
            ' Assign the new path
            currentEffect.Behaviors(1).MotionEffect.Path = motionPathString$
        End If
    Next currentEffect
End Sub

'
' This converts an angle from degrees to radians. At the
' same time, since shape rotation angles are computed in PowerPoint
' starting from the positive Y semiaxis and going in
' clockwise direction, it reverses the convention by returning
' an angle in radiants that starts from the positive X semiaxis
' and goes counterclockwise.
'
Private Function degToRad(degAngle) As Double
    degToRad = 3.14159265358979 * ((360 - degAngle) Mod 360) / 180
End Function

'
' This subroutine gets the last (i.e., arrival) coordinates from
' a string describing a motion path. Extracted coordinates are put
' in lastX and lastY, while lastTokenBeforeCoordinates will be
' updated with the index of the token in pathString$ that precedes
' the last coordinates.
'
Private Sub getLastCoordinates(pathString$, lastX As Double, lastY As Double, lastTokenBeforeCoordinates)
    pathStringTokens = Split(pathString$)
    tokenIndex = UBound(pathStringTokens)
    While tokenIndex > 0
        If pathStringTokens(tokenIndex) <> "" And _
            Not (Mid$(pathStringTokens(tokenIndex), 1, 1) >= "A" And _
            Mid$(pathStringTokens(tokenIndex), 1, 1) <= "Z") Then
            lastY = localizeDecimalSeparators(pathStringTokens(tokenIndex))
            lastX = localizeDecimalSeparators(pathStringTokens(tokenIndex - 1))
            lastTokenBeforeCoordinates = tokenIndex - 2
            Exit Sub
        End If
        tokenIndex = tokenIndex - 1
    Wend
End Sub


'
' This subroutine does what it says: it applies an emphasis
' (or motion) effect to a shape. Arguments are: the sequence of
' effects (which will only be used to update motion path coordinates),
' the emphasis effect to be applied, and the shape it applies to
'
Private Sub applyEmphasisEffect(seq As Sequence, e As Effect, sh As Shape)
    On Error GoTo recover
    ePar = getEffectParagraph(e)
    ' Here I should be supposed to check the value of
    ' e.Shape.HasTextFrame before attemping to access
    ' the sh.TextFrame.TextRange property. Guess what?
    ' In some cases PowerPoint returns false even if
    ' properties like sh.TextFrame.TextRange.Font.Size
    ' can be accessed. Is it me or could this be yet
    ' another bug?
    ' Worked around by attempting assignments anyway, and
    ' watching for errors during the process.
    On Error Resume Next
    shTextRange = Null
    If ePar > 0 Then
        ' This effect applies to a text paragraph
        Set shTextRange = sh.TextFrame.TextRange.Paragraphs(ePar)
    Else
        Set shTextRange = sh.TextFrame.TextRange
    End If
    On Error GoTo recover
    ' Note: if an effect acts both on a text element and on its container
    ' shape, then the effect must first be applied to the container shape,
    ' in order to avoid unpredictable automatic resizing.
    If e.EffectType = msoAnimEffectGrowShrink Then
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            ' I am not scaling a bitmap here, therefore I need to
            ' recompute map X and Y scaling in accordance with the shape
            ' rotation.
            rotCos = Cos(degToRad(sh.Rotation))
            rotSin = Sin(degToRad(sh.Rotation))
            scaleX = e.Behaviors(1).ScaleEffect.ByX / 100 * Abs(rotCos) + e.Behaviors(1).ScaleEffect.ByY / 100 * Abs(rotSin)
            scaleY = e.Behaviors(1).ScaleEffect.ByX / 100 * Abs(rotSin) + e.Behaviors(1).ScaleEffect.ByY / 100 * Abs(rotCos)
            ' Disable size autofitting for text frames and unlock
            ' aspect ratio
            sh.LockAspectRatio = msoFalse
            On Error Resume Next
            sh.TextFrame.AutoSize = ppAutoSizeNone
            On Error GoTo recover
            sh.ScaleWidth scaleX, msoFalse, msoScaleFromMiddle
            sh.ScaleHeight scaleY, msoFalse, msoScaleFromMiddle
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Size = shTextRange.Font.Size * (e.Behaviors(1).ScaleEffect.ByX / 100 + e.Behaviors(1).ScaleEffect.ByY / 100) / 2
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectChangeFontColor Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                assignColor shTextRange.Font.Color, e.EffectParameters.Color2
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectChangeFillColor Then
        If sh.Fill.Transparency < 1 Then
            sh.Fill.Solid
        End If
        assignColor sh.Fill.ForeColor, e.EffectParameters.Color2
    ElseIf e.EffectType = msoAnimEffectChangeFontStyle Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Italic = (e.Behaviors(1).SetEffect.To = 1)
                shTextRange.Font.Bold = (e.Behaviors(2).SetEffect.To = 1)
                shTextRange.Font.Underline = (e.Behaviors(3).SetEffect.To = 1)
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectTransparency Then
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Line.Transparency < 1 Then
                sh.Line.Transparency = e.EffectParameters.amount
            End If
            If sh.Fill.Transparency < 1 Then
                sh.Fill.Transparency = e.EffectParameters.amount
            End If
        End If
        ' Only Office 2007 or newer exposes text font transparency
        ' in VBA, therefore this piece of code has been removed.
    ElseIf e.EffectType = msoAnimEffectChangeFont Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Name = e.EffectParameters.FontName
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectChangeLineColor Then
        If Not sh.Line.Visible Then sh.Line.Visible = msoTrue
        assignColor sh.Line.ForeColor, e.EffectParameters.Color2
    ElseIf e.EffectType = msoAnimEffectChangeFontSize Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                ' Please leave the /1 alone: it is required for some strange internal
                ' type conversion, otherwise leading to improper font sizes :-(
                shTextRange.Font.Size = shTextRange.Font.Size * e.Behaviors(1).PropertyEffect.To / 1
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectSpin Then
        ' Rotating just the text is not supported
        sh.Rotation = sh.Rotation + e.Behaviors(1).RotationEffect.By
    ElseIf e.EffectType = msoAnimEffectDesaturate Then
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                With sh.Fill.ForeColor
                    colToRGB .RGB, r, g, b
                    .RGB = RGB((r + g + b) / 3, (r + g + b) / 3, (r + g + b) / 3)
                End With
                With sh.Fill.BackColor
                    colToRGB .RGB, r, g, b
                    .RGB = RGB((r + g + b) / 3, (r + g + b) / 3, (r + g + b) / 3)
                End With
            End If
            If sh.Line.Transparency < 1 Then
                With sh.Line.ForeColor
                    colToRGB .RGB, r, g, b
                    .RGB = RGB((r + g + b) / 3, (r + g + b) / 3, (r + g + b) / 3)
                End With
            End If
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                With shTextRange.Font.Color
                    colToRGB .RGB, r, g, b
                    .RGB = RGB((r + g + b) / 3, (r + g + b) / 3, (r + g + b) / 3)
                End With
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectColorWave Or e.EffectType = msoAnimEffectColorBlend Or _
            e.EffectType = msoAnimEffectBrushOnColor Or e.EffectType = msoAnimEffectTeeter Then
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                assignColor sh.Fill.ForeColor, e.EffectParameters.Color2
            End If
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                assignColor shTextRange.Font.Color, e.EffectParameters.Color2
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectComplementaryColor2 Then
        ' PowerPoint computes the complementary color in some other way.
        ' I feel pretty satisfied with this rotation in the HSL space
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                rotateColor sh.Fill.ForeColor, 180
            End If
            If sh.Line.Transparency < 1 Then
                rotateColor sh.Line.ForeColor, 180
            End If
        End If
    ElseIf e.EffectType = msoAnimEffectVerticalGrow Then
        ' Font scaling alone is not supported for this effect
            
        ' Disable size autofitting for text frames and unlock
        ' aspect ratio
        sh.LockAspectRatio = msoFalse
        On Error Resume Next
        sh.TextFrame.AutoSize = ppAutoSizeNone
        On Error GoTo recover
        sh.ScaleHeight 1.5, msoFalse
        shiftY = sh.Height / 4
        If sh.Fill.Transparency < 1 Then
            assignColor sh.Fill.ForeColor, e.EffectParameters.Color2
        End If
        sh.Top = sh.Top - shiftY
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                assignColor shTextRange.Font.Color, e.EffectParameters.Color2
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectLighten Then
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                changeLightness sh.Fill.ForeColor, 0.3
            End If
            If sh.Line.Transparency < 1 Then
                changeLightness sh.Line.ForeColor, 0.3
            End If
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                changeLightness shTextRange.Font.Color, 0.3
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectBrushOnUnderline Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Underline = msoTrue
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectComplementaryColor Then
        ' PowerPoint computes the complementary color in some other way.
        ' I feel pretty satisfied with this rotation in the HSL space
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                rotateColor sh.Fill.ForeColor, 120
            End If
            If sh.Line.Transparency < 1 Then
                rotateColor sh.Line.ForeColor, 120
            End If
        End If
    ElseIf e.EffectType = msoAnimEffectContrastingColor Then
        ' PowerPoint computes the contrasting color in some other way.
        ' I feel pretty satisfied with this rotation in the HSL space
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                rotateColor sh.Fill.ForeColor, 90
            End If
            If sh.Line.Transparency < 1 Then
                rotateColor sh.Line.ForeColor, 90
            End If
        End If
    ElseIf e.EffectType = msoAnimEffectBoldFlash Then
        ' msoAnimEffectBoldFlash is a non-permanent effect
    ElseIf e.EffectType = msoAnimEffectFlashBulb Then
        ' msoAnimEffectFlashBulb is a non-permanent effect
    ElseIf e.EffectType = msoAnimEffectDarken Then
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                changeLightness sh.Fill.ForeColor, -0.3
            End If
            If sh.Line.Transparency < 1 Then
                changeLightness sh.Line.ForeColor, -0.3
            End If
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                changeLightness shTextRange.Font.Color, -0.3
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectGrowWithColor Then
        If sh.Fill.Transparency < 1 Then
            sh.Fill.Solid
            assignColor sh.Fill.ForeColor, e.EffectParameters.Color2
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Size = shTextRange.Font.Size * 1.5
                assignColor shTextRange.Font.Color, e.EffectParameters.Color2
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectFlicker Then
        ' msoAnimEffectFlicker is a non-permanent effect
    ' *** WARNING: the shaking effect has no associated effecttype (PowerPoint bug :-((( )
    ElseIf e.EffectType = msoAnimEffectBoldReveal Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Bold = msoTrue
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ElseIf e.EffectType = msoAnimEffectWave Then
        ' msoAnimEffectWave is a non-permanent effect
    ElseIf e.EffectType = msoAnimEffectStyleEmphasis Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeId = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Italic = msoTrue
                shTextRange.Font.Bold = msoTrue
                shTextRange.Font.Underline = msoTrue
                assignColor shTextRange.Font.Color, e.EffectParameters.Color2
            End If
            shapeId = shapeId + 1
            If sh.Type = msoGroup Then
                If shapeId > sh.GroupItems.Count Then
                    shapeId = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeId).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeId = 0
            End If
        Loop Until shapeId = 0
    ' *** WARNING: the blinking effect has no associated effecttype (PowerPoint bug :-((( )
    ElseIf e.EffectType = msoAnimEffectBlast Then
        ' msoAnimEffectBlast has too vague a behavior to be implemented :-O
    Else
        If isEmphasisEffect(e) Then
            On Error GoTo 0
            ' Ok, this is neither an emphasis effect nor an entry effect:
            ' it must be a motion effect
            motionpath = Split(e.Behaviors(1).MotionEffect.Path)
            Dim lastX As Double, lastY As Double
            If e.Behaviors(1).Timing.Speed < 0 Then
                lastX = localizeDecimalSeparators(motionpath(1))
                lastY = localizeDecimalSeparators(motionpath(2))
            Else
                getLastCoordinates e.Behaviors(1).MotionEffect.Path, lastX, lastY, lastToken
            End If
            ' Coordinates are expressed in VML (see http://www.w3.org/TR/1998/NOTE-VML-19980513#_Toc416858391)
            ' as multiples of the slide width/height and are relative to the shape center
            shapeCenterX = (sh.Left + sh.Width / 2) / ActivePresentation.PageSetup.SlideWidth
            shapeCenterY = (sh.Top + sh.Height / 2) / ActivePresentation.PageSetup.SlideHeight
            newX = (shapeCenterX + lastX) * ActivePresentation.PageSetup.SlideWidth
            newY = (shapeCenterY + lastY) * ActivePresentation.PageSetup.SlideHeight
            sh.Left = newX - sh.Width / 2
            sh.Top = newY - sh.Height / 2
            shiftAllMotions seq, sh, -lastX, -lastY
        End If
    End If
    Exit Sub
recover:
    ' Ok, Powerpoint bug again: this is an emphasis effect that
    ' has no EffectType member. Let's pass it by.
End Sub

'
' This function returns true if (and only if) the effect given
' as argument is a motion (path) effect
'
Private Function isPathEffect(e As Effect) As Boolean
    On Error GoTo pathRecover
    isPathEffect = False
    ' The following conditions have been built starting from the page "Powerpoint
    ' constants" of the VBA documentation.
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPath5PointStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCrescentMoon
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSquare
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTrapezoid
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathHeart
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathOctagon
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPath6PointStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathFootball
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathEqualTriangle
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathParallelogram
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathPentagon
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPath4PointStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPath8PointStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTeardrop
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathPointyStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCurvedSquare
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCurvedX
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathVerticalFigure8
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCurvyStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathLoopdeLoop
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathBuzzsaw
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathHorizontalFigure8
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathPeanut
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathFigure8Four
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathNeutron
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSwoosh
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathBean
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathPlus
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathInvertedTriangle
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathInvertedSquare
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathLeft
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTurnRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathArcDown
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathZigzag
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSCurve2
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSineWave
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathBounceLeft
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathDown
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTurnUp
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathArcUp
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathHeartbeat
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSpiralRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathWave
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCurvyLeft
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathDiagonalDownRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTurnDown
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathArcLeft
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathFunnel
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSpring
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathBounceRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSpiralLeft
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathDiagonalUpRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTurnUpRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathArcRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSCurve1
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathDecayingWave
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCurvyRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathStairsDown
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathUp
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathRight

    ' 0 = msoAnimEffectCustom = Customized path
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectCustom
    Exit Function
    
pathRecover:
    ' Powerpoint bug: this effect has no EffectType property;
    ' I cannot either recognize or handle it. At the time of
    ' writing this code, there were no motion effects affected
    ' by this problem, therefore this is not a motion effect.
    isPathEffect = False
End Function


'
' This function returns true iff the given effect is either
' an emphasis effect or a motion effect.
'
Private Function isEmphasisEffect(e As Effect) As Boolean
    On Error GoTo recoverIsEmphasis
    isEmphasisEffect = False
    ' The following conditions have been built starting from the page "Powerpoint
    ' constants" of the VBA documentation.
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectGrowShrink
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeFontColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeFillColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeFontStyle
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectTransparency
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeFont
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeLineColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeFontSize
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectSpin
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectDesaturate
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectColorWave
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectComplementaryColor2
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectVerticalGrow
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectLighten
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectColorBlend
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectBrushOnUnderline
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectBrushOnColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectComplementaryColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectContrastingColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectBoldFlash
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectFlashBulb
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectDarken
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectGrowWithColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectTeeter
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectFlicker
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectBoldReveal
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectWave
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectStyleEmphasis
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectBlast
    
    isEmphasisEffect = isEmphasisEffect Or isPathEffect(e)

    ' If isEmphasisEffect is true at this point, then I have
    ' an emphasis or motion effect. But let's really make sure it is not
    ' an entry/exit effect.
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectAppear
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFly
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectBlinds
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectBox
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectCheckerboard
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectCircle
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectCrawl
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectDiamond
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectDissolve
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFade
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFlashOnce
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectPeek
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectPlus
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectRandomBars
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSpiral
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSplit
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectStretch
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectStrips
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSwivel
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectWedge
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectWheel
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectWipe
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectZoom
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectRandomEffects
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectBoomerang
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectBounce
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectColorReveal
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectCredits
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectEaseIn
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFloat
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectGrowAndTurn
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectLightSpeed
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectPinwheel
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectRiseUp
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSwish
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectThinLine
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectUnfold
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectWhip
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectAscend
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectCenterRevolve
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFadedSwivel
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectDescend
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSling
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSpinner
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectStretchy
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectZip
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectArcUp
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFadedZoom
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectGlide
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectExpand
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFlip
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFold
    Exit Function
recoverIsEmphasis:
    ' Powerpoint bug: this effect has no EffectType property;
    ' I cannot either recognize or handle it. Luckily enough,
    ' there is no need to process the affected effects because
    ' they are non-permanent (apart from the color that the
    ' shaking effect allows to apply to the shape). Here I
    ' assume that an unrecognizable effect is an emphasis effect.
    isEmphasisEffect = True
End Function
'
' This function takes an effect as argument. If the
' effect is applied to a text paragraph, it returns the
' index of that text paragraph (in its container shape).
' Otherwise, it returns -1.
'
Private Function getEffectParagraph(e As Effect)
    paragraph_idx = -1
    On Error Resume Next
    ' The following assignment may fail because the Paragraph property does not
    ' exist at all for those effects that are applied to shapes instead of text.
    ' But, was this truly expected by design? :-?
    paragraph_idx = e.Paragraph
    On Error GoTo 0
    getEffectParagraph = paragraph_idx
End Function

'
' This subroutine deletes all the shapes for which the first
' effect in the sequence is an entry effect. This is reasonable,
' because those shapes are expected to appear later on.
'
Private Sub purgeFutureShapes(s As Slide, textParagraphEffectsOnly As Boolean)
    Dim slide_timeline As Sequence
    Set slide_timeline = s.TimeLine.MainSequence
    ProgressForm.infoLabel = "Preprocessing slide effects..."
    ProgressForm.Repaint
    If doNotSplitMouseTriggered Then
        start_deleting_at = 1
    Else
        i = 1: start_deleting_at = 0
        While i <= slide_timeline.Count And start_deleting_at = 0
            If slide_timeline(i).Timing.TriggerType <> msoAnimTriggerAfterPrevious And _
               slide_timeline(i).Timing.TriggerType <> msoAnimTriggerWithPrevious Then
                ' Start deleting shapes from the next mouse-triggered event.
                ' Any preceding shapes will be deleted when their effects
                ' are individually considered
                start_deleting_at = i
            End If
            i = i + 1
        Wend
    End If
    
    If start_deleting_at > 0 Then
        For i = start_deleting_at To s.TimeLine.MainSequence.Count
            If i > s.TimeLine.MainSequence.Count Then Exit For
            delete_shape_idx = -1
            If Not slide_timeline(i).Exit And Not isEmphasisEffect(slide_timeline(i)) Then
                ' This is an entry effect applied in the future. Likely a candidate
                ' to justify shape deletion
                delete_shape_idx = i
            End If
            parI = getEffectParagraph(slide_timeline(i))
            For j = i - 1 To start_deleting_at Step -1
                If slide_timeline(i).Shape Is slide_timeline(j).Shape And _
                    (slide_timeline(j).Exit Or isEmphasisEffect(slide_timeline(j))) Then
                    ' Probably we need to abort deletion: there may
                    ' be an exit/emphasis effect for the same shape before the entry effect.
                    ' In that case, this means that the shape must be visible at the
                    ' beginning. However, first we need to check if this is a paragraph
                    ' effect and, in that case, if the exit/emphasis
                    ' effect applies to the very same paragraph.
                    parJ = getEffectParagraph(slide_timeline(j))
                    If parI = parJ Then
                        ' Either none of the effects is a paragraph effect (in which
                        ' case the match is ok because both effects work on the same shape)
                        ' or both effects are paragraph effects and work on the same paragraph
                        ' (in which case the match is still ok because they affect the
                        ' same graphical element). If the match is ok, then deletion
                        ' must be aborted.
                        delete_shape_idx = -1
                    End If
                End If
            Next j
            If delete_shape_idx > 0 Then
                ' Delete shapes for which a following entry effect exists.
                ' Restrict deletion to text paragraphs only if instructed to
                ' do so.
                If parI > 0 Or Not textParagraphEffectsOnly Then
                    ' Pay attention, because shape deletion (not paragraph deletion)
                    ' causes animation effects to disappear from the timeline, so we
                    ' need to decrease i in order to keep in sync with the currently
                    ' processed effect.
                    ' In general, deletion of a shape may cause several preceding
                    ' effects to also disappear: here we count how many in order to
                    ' understand how many positions should i go backward (note that
                    ' future effects for the same shapes should not be counted, because
                    ' they will safely disappear from the timeline without the need
                    ' to realign the value of i).
                    prevEffectsForThisShape = 0
                    For k = 1 To i
                        If slide_timeline(k).Shape Is slide_timeline(i).Shape Then
                            prevEffectsForThisShape = prevEffectsForThisShape + 1
                        End If
                    Next k
                    ' Assertion: at the end of the above iteration, prevEffectsForThisShape
                    ' should always be >0 (because at least the i'th effect affects that
                    ' shape)
                    If deleteShape(slide_timeline(i).Shape, slide_timeline, delete_shape_idx) Then
                        i = i - prevEffectsForThisShape
                    End If
                End If
            End If
        Next i
    End If
    ProgressForm.infoLabel = ""
    ProgressForm.Repaint
End Sub

'
' This function returns the sequential number of a shape in s
' that matches the id, or 0 if no such shape exists. The
' function relies on the values of the "shapeId" tag,
' which must have been set up in advance using the
' copyShapeIds subroutine.
'
Private Function findShape(s As Slide, id)
    Dim currentShape As Shape
    i = 1
    findShape = 0
    For Each currentShape In s.Shapes
        If currentShape.Tags("shapeId") = id Then
            findShape = i
            Exit Function
        End If
        i = i + 1
    Next currentShape
End Function

'
' This subroutine applies to slide s a generic animation effect that is
' on top of the timeline of seq_slide. At the same time, it also removes
' the effect from the timeline of seq_slide. Returns 0 if behaving normally.
' Returns 1 in the exceptional case when an animation effect is added by
' the function itself.
'
Private Function applyEffect(s As Slide, seq_slide As Slide)
    Dim current_effect As Effect, sh As Shape
    Set current_effect = seq_slide.TimeLine.MainSequence(1)
    Set sh = current_effect.Shape
    ' By default the applyEffect function only consumes effects, does not add them
    applyEffect = 0
    If current_effect.EffectInformation.AfterEffect = msoAnimAfterEffectHide Then
        ' This effect is set for hiding the shape after the animation, so it
        ' must be treated equivalently to an exit effect: simply delete the shape
        If findShape(s, sh.Tags("shapeId")) > 0 Then
            deleteShape s.Shapes(findShape(s, sh.Tags("shapeId"))), seq_slide.TimeLine.MainSequence, 1
        End If
        current_effect.Delete
    Else
        If current_effect.EffectInformation.AfterEffect = msoAnimAfterEffectHideOnNextClick Then
            ' This effect is set for hiding after the next click:
            ' insert a new exit animation that will be processed in the following
            found = False
            Set tl = seq_slide.TimeLine.MainSequence
            For i = 2 To tl.Count
                If tl(i).Timing.TriggerType = msoAnimTriggerOnPageClick Then
                    tl.AddEffect current_effect.Shape, msoAnimEffectDissolve, , msoAnimTriggerWithPrevious
                    ' Best thing would be to insert the exit effect right after the next click-triggered
                    ' effect, but this is not possible, guess why, due to a PowerPoint bug which causes
                    ' the Index argument of AddEffect to be handled unpredictably. So, we need to work this
                    ' around by inserting the effect at the end of the sequence and, only afterwards,
                    ' move it to the right location.
                    tl(tl.Count).MoveTo i + 1
                    tl(i + 1).Exit = msoTrue
                    found = True
                    Exit For
                End If
            Next i
            If Not found Then
                tl.AddEffect current_effect.Shape, msoAnimEffectDissolve, , msoAnimTriggerOnPageClick, i
                tl(i).Exit = msoTrue
            End If
            ' This is the only case when the applyEffect function adds an animation effect to the
            ' sequence: here we notify the calling routine about the fact that the animation sequence
            ' has lengthened.
            applyEffect = 1
        End If
        If current_effect.Timing.RewindAtEnd Then
            ' A rewound-after-the-end animation has no effect (unless it is set for
            ' being hidden after the animation, which has already been checked)
            current_effect.Delete
        Else
            If current_effect.Exit Then
                ' This is an exit effect: simply delete the shape (or the text
                ' paragraph) from the next slide
                If findShape(s, sh.Tags("shapeId")) > 0 Then
                    deleteShape s.Shapes(findShape(s, sh.Tags("shapeId"))), seq_slide.TimeLine.MainSequence, 1
                End If
                current_effect.Delete
            Else
                If isEmphasisEffect(current_effect) Then
                    ' This is an emphasis (or motion) effect. Note that an autoreversed emphasis
                    ' effect has no overall effect. Also, an emphasis effect can never be applied
                    ' to a single text paragraph
                    If Not current_effect.Timing.AutoReverse Then
                        If findShape(s, sh.Tags("shapeId")) > 0 Then
                            applyEmphasisEffect seq_slide.TimeLine.MainSequence, seq_slide.TimeLine.MainSequence(1), s.Shapes(findShape(s, sh.Tags("shapeId")))
                        End If
                    End If
                    current_effect.Delete
                Else
                    ' This is an entry effect.
                    If Not findShape(s, sh.Tags("shapeId")) > 0 Then
                        ' The shape is not already present
                        sh.Copy
                        ' Invoke purgeEffects to clear any subsequent entry
                        ' effects, which may interfere
                        ' with calls to purgeFutureShapes below in this same
                        ' subroutine.
                        ' (note that these subsequent calls may happen when
                        ' in the same slide multiple objects appear simultaneously,
                        ' and therefore applyEffect is invoked multiple times).
                        purgeEffects s
                        s.Shapes.Paste
                        Set newShape = s.Shapes(findShape(s, sh.Tags("shapeId")))
                        ' Coordinates of the pasted shape are sometimes
                        ' automatically adjusted (for example if the shape
                        ' overlaps with another one)
                        newShape.Left = sh.Left
                        newShape.Top = sh.Top
                        par = -1
                        On Error Resume Next
                        ' The following assignment may raise an error for missing
                        ' Paragraph property
                        par = current_effect.Paragraph
                        On Error GoTo 0
                        If par > 0 Then
                            ' Remove all the paragraphs that are supposed to appear later
                            For parIdx = 1 To newShape.TextFrame.TextRange.Paragraphs.Count
                                If parIdx <> par Then
                                    foundEntryAnim = False
                                    For k = 1 To seq_slide.TimeLine.MainSequence.Count
                                        If seq_slide.TimeLine.MainSequence(k).Shape Is sh And Not isEmphasisEffect(seq_slide.TimeLine.MainSequence(k)) _
                                            And Not seq_slide.TimeLine.MainSequence(k).Exit Then
                                            On Error Resume Next
                                            If seq_slide.TimeLine.MainSequence(k).Paragraph = parIdx Then
                                                foundEntryAnim = True
                                            End If
                                            On Error GoTo 0
                                        End If
                                    Next k
                                    If foundEntryAnim Then
                                        clearParagraph s.Shapes(findShape(s, sh.Tags("shapeId"))), parIdx
                                    End If
                                End If
                            Next parIdx
                        End If
                        ' Sometimes text auto-fitting does not seem to act
                        ' properly: this is an attempt to "awaken" it by
                        ' notifying of a change in the shape size
                        newShape.Width = sh.Width
                        newShape.Height = sh.Height
                        ' Now we have pasted the shape. Note that we paste
                        ' only one shape at a time, therefore it should carry
                        ' with itself its own entry effect. There is one
                        ' exception: a single text box shape may be associated with
                        ' several subsequent entry effects, that correspond
                        ' to paragraphs in the text appearing one after the
                        ' other (and after the text box itself has appeared).
                        ' We should get rid of paragraphs that are supposed
                        ' to appear later on, and this is why we call purgeFutureShapes
                        ' also here. Note that we should remove the entry effect
                        ' for the shape we have just added before invoking
                        ' purgeFutureShapes, or the shape itself will be
                        ' deleted!
                        s.TimeLine.MainSequence(1).Delete
                        purgeFutureShapes s, True
                    Else
                        ' The shape is already present: I only need to add a
                        ' paragraph to it, if required.
                        par = -1
                        ' The following assignment may raise an error for missing
                        ' Paragraph property
                        On Error Resume Next
                        par = current_effect.Paragraph
                        On Error GoTo 0
                        If par > 0 Then
                            Set newShape = s.Shapes(findShape(s, sh.Tags("shapeId")))
                            copyParagraph s.Shapes(findShape(s, sh.Tags("shapeId"))).TextFrame.TextRange.Paragraphs(par), sh.TextFrame.TextRange.Paragraphs(par)
                            
                            ' Attempt to preserve indentations and margins (these are not
                            ' part of paragraph information, but rather of a Ruler object).
                            ' In principle, the number of ruler levels (i.e., possible
                            ' indentation levels) is fixed. However, according to the documentation
                            ' it should be 5 whereas in practice I have seen cases where it
                            ' counts up to 9. To stay on the safe side, the number of
                            ' ruler levels here is parametric.
                            For ruler_level = 1 To sh.TextFrame.Ruler.Levels.Count
                                ' For some obscure reasons, out-of-range margins are sometimes
                                ' returned (for example, corresponding to the smallest possible
                                ' value in a Long variable). In this case, it's better to
                                ' refrain from copying the margin value, or an error would be
                                ' raised.
                                If Abs(sh.TextFrame.Ruler.Levels(ruler_level).FirstMargin) < 10000000 Then
                                    newShape.TextFrame.Ruler.Levels(ruler_level).FirstMargin = sh.TextFrame.Ruler.Levels(ruler_level).FirstMargin
                                End If
                                If Abs(sh.TextFrame.Ruler.Levels(ruler_level).LeftMargin) < 10000000 Then
                                    newShape.TextFrame.Ruler.Levels(ruler_level).LeftMargin = sh.TextFrame.Ruler.Levels(ruler_level).LeftMargin
                                End If
                            Next ruler_level
                            
                            ' Sometimes text auto-fitting does not seem to act
                            ' properly: this is an attempt to "awaken" it by
                            ' notifying of a change in the shape size
                            newShape.Width = sh.Width
                            newShape.Height = sh.Height
                        End If
                    End If
                    current_effect.Delete
                End If
            End If
        End If
    End If
End Function

'
' This subroutine removes all the animation effects from a slide. Useful
' to leave slides clean after processing
'
Private Sub purgeEffects(s As Slide)
    For i = 1 To s.TimeLine.MainSequence.Count
        s.TimeLine.MainSequence(1).Delete
    Next i
    s.SlideShowTransition.EntryEffect = ppEffectNone
End Sub

'
' This function copies shape Ids to a less volatile Tag. This is
' very useful to match different instances of the same shape in different
' slides, as the copy-and-paste process used to implement entry effects
' discards the shape id.
'
Private Sub copyShapeIds(s As Slide)
    Dim sh As Shape
    For Each sh In s.Shapes
        sh.Tags.Add "shapeId", Str$(sh.id)
    Next sh
End Sub

'
' This is a support function for the bakeSlideNumbers sub. What it does
' is to move a currently selected placeholder to the slides that are set
' to show it, and bake text that appears inside it.
'
Private Sub bakePlaceholder(footerElement, start_index, end_index, designIndex, titleMasterExists As Boolean, isTitleMaster As Boolean)
    Dim sh As Shape, currentSlide As Slide
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        ' Make the shape recognizeable as a placeholder before moving it
        ActiveWindow.Selection.ShapeRange.Tags.Add "placeholder", Right$(Str$(footerElement), 1)
        ActiveWindow.Selection.ShapeRange.Tags.Add "shapeId", "placeholder" + Right$(Str$(footerElement), 1)
        ' Remove shape from the slide master (will be pasted later in the slides
        ' where it is supposed to appear)
        ActiveWindow.Selection.Cut
        ActiveWindow.ViewType = ppViewNormal
        For Each currentSlide In ActivePresentation.Slides
            With currentSlide
                sameDesign = (.Design.Index = designIndex)
                If sameDesign Then
                    ' Layout can only be checked if the current slide uses the
                    ' design from which we took the placeholder. Otherwise, we risk
                    ' to seek a TitleMaster for a design that does not have it.
                    If titleMasterExists Then
                        If isTitleMaster Then
                            matchingLayout = (.Layout = ppLayoutTitle And .Design.TitleMaster.HeadersFooters.DisplayOnTitleSlide)
                        Else
                            matchingLayout = (.Layout <> ppLayoutTitle)
                        End If
                    Else
                        matchingLayout = ((.Layout = ppLayoutTitle And .Design.SlideMaster.HeadersFooters.DisplayOnTitleSlide) Or _
                                          (.Layout <> ppLayoutTitle))
                    End If
                End If
                If sameDesign And matchingLayout Then
                    If (footerElement = 1 And .HeadersFooters.DateAndTime.Visible) Or _
                       (footerElement = 2 And .HeadersFooters.Footer.Visible) Or _
                       (footerElement = 3 And .HeadersFooters.SlideNumber.Visible) Then
                        .Shapes.Paste
                        For Each sh In .Shapes
                            If sh.Tags("placeholder") <> "" Then
                                sh.ZOrder msoSendToBack
                                ' Text is baked character by character, in order to avoid losing formatting
                                For c = 1 To sh.TextFrame.TextRange.Characters.Count
                                     sh.TextFrame.TextRange.Characters(c) = sh.TextFrame.TextRange.Characters(c)
                                Next c
                                ' Shape names must be unique (a "Permission Denied" error is raised otherwise)
                                sh.Name = "slideNumberPlaceholder" & Str$(sh.id)
                            End If
                        Next sh
                    End If
                End If
            End With
        Next currentSlide
    End If
End Sub


'
' This function moves elements from slide masters to slides, in order to keep slide
' numbers fixed during the split. Note that slide numbers may occur in several shapes
' in a slide master, not just the "slide number" footer: slide numbers appearing in
' such extra shapes will not be processed.
'
Private Sub bakeSlideNumbers(start_index, end_index)
    Dim shs As Shapes, d As Design, sh As Shape


    ProgressForm.infoLabel = "Adjusting slide numbers. This may take some time..."
    
    ' Cycle through all slide masters, including title ones, and move relevant placeholders
    ' to all the slides that use them. When moving, text is reassigned so that any special
    ' <pagenumber> field is replaced by its actual value

    For footerElement = 1 To 3
        For d_index = 1 To ActivePresentation.Designs.Count
            ' PowerPoint requires a specific Design to be currently displayed in order
            ' to be able to select its shapes. Now, since the only way in PowerPoint 2003
            ' to switch to a Design view is to use ppViewTitleMaster or ppViewSlideMaster,
            ' both pointing to the first Design, here is a
            ' horrible hack to always keep the Design of interest first. Interestingly,
            ' reordering Designs does not have adverse effects on their usage in slides.
            ActivePresentation.Designs(d_index).MoveTo 1
            Set d = ActivePresentation.Designs(1)
            
            ProgressForm.SlideBar.value = (ActivePresentation.Designs.Count * (footerElement - 1) + d_index) / (ActivePresentation.Designs.Count * 3) * 100
            ProgressForm.Repaint
            ' Clear current selection (if any)
            ActiveWindow.Selection.Unselect
            If d.HasTitleMaster Then
                ' Must switch to an appropriate view in order to be able to select shapes.
                ' Note that we first switch to slide view because, in order to make sure that
                ' the first available title master is selected, we must come from a different
                ' view.
                ActiveWindow.ViewType = ppViewSlide
                ActiveWindow.ViewType = ppViewTitleMaster
                For Each sh In d.TitleMaster.Shapes
                    If sh.Type = msoPlaceholder Then
                        If (footerElement = 1 And sh.PlaceholderFormat.Type = ppPlaceholderDate) Or _
                           (footerElement = 2 And sh.PlaceholderFormat.Type = ppPlaceholderFooter) Or _
                           (footerElement = 3 And sh.PlaceholderFormat.Type = ppPlaceholderSlideNumber) Then sh.Select msoTrue
                    End If
                Next sh
                bakePlaceholder footerElement, start_index, end_index, d.Index, True, True
            End If
            ' Must switch to an appropriate view in order to be able to select shapes.
            ' Note that we first switch to slide view because, in order to make sure that
            ' the first available slide master is selected, we must come from a different
            ' view.
            ActiveWindow.ViewType = ppViewSlide
            ActiveWindow.ViewType = ppViewSlideMaster
            For Each sh In d.SlideMaster.Shapes
                If sh.Type = msoPlaceholder Then
                    If (footerElement = 1 And sh.PlaceholderFormat.Type = ppPlaceholderDate) Or _
                       (footerElement = 2 And sh.PlaceholderFormat.Type = ppPlaceholderFooter) Or _
                       (footerElement = 3 And sh.PlaceholderFormat.Type = ppPlaceholderSlideNumber) Then sh.Select msoTrue
                End If
            Next sh
            bakePlaceholder footerElement, start_index, end_index, d.Index, d.HasTitleMaster, False
        Next d_index
    Next footerElement
                
    ActiveWindow.ViewType = ppViewNormal
    
End Sub

'
' This function enriches existing slide numbers with a subindex, namely a progressive
' number assigned anew to each slide resulting from splitting a single original one.
' It works in close conjunction with bakeSlideNumbers, with a main difference:
' - bakeSlideNumbers is invoked once on all the slide deck to make slide numbers
'   persistent
' - augmentSlideNumbers is invoked once for each split slide, strictly after processing
'   of that slide has finished and, possibly, after a duplicate of that slide is
'   generated (modified slide numbers would otherwise be inherited in all subsequent
'   slides)
'
Private Sub augmentSlideNumbers(slide_number, progressive_slide_count)
    Dim sh As Shape

    For Each sh In ActivePresentation.Slides(slide_number).Shapes
        If slideNumbersAdjustMode = SLIDENUMBER_SUBINDEX And Left$(sh.Name, 22) = "slideNumberPlaceholder" Then
            sh.TextFrame.TextRange.InsertAfter "." + Right$(Str$(progressive_slide_count), Len(Str$(progressive_slide_count)) - 1)
        End If
    Next sh
End Sub

Sub PPspliT_main()
    On Error GoTo error_handler

    If Application.Presentations.Count = 0 Then
        Exit Sub
    End If

    Dim slide_timeline As Sequence
    cancelStatus = False
    
    ' Non-contiguous ranges of slides are NOT supported: they are assumed to
    ' start at the lowest numbered selected slide and end at the highest numbered
    ' selected slide.
    If ActiveWindow.Selection.Type = ppSelectionSlides Then
        min_slide_index = 32767
        max_slide_index = 0
        For Each s In ActiveWindow.Selection.SlideRange
            If s.SlideIndex < min_slide_index Then min_slide_index = s.SlideIndex
            If s.SlideIndex > max_slide_index Then max_slide_index = s.SlideIndex
        Next s
        slide_number = min_slide_index
        tot_slides = max_slide_index
        split_selected_slides = MsgBox(prompt:="It seems that a set of slides is currently selected. " + _
             "By proceeding, you will only be splitting slides in the range" + Str$(min_slide_index) + "-" + Right$(Str$(max_slide_index), Len(Str$(max_slide_index)) - 1) + "." + Chr$(13) + _
             "(non-contiguous sets of slides are not supported, therefore all the slides between the first and last selected ones will be affected by the split process)." + Chr$(13) + _
             "Click " + Chr$(34) + "Yes" + Chr$(34) + " if this is what you want." + Chr$(13) + _
             "Click " + Chr$(34) + "No" + Chr$(34) + " if you want to split ALL the slides in the presentation instead." + Chr$(13) + _
             "Click " + Chr$(34) + "Cancel" + Chr$(34) + " to simply cancel the operation.", buttons:=vbYesNoCancel, Title:="PPspliT - Information request")
        If split_selected_slides = vbNo Then
            slide_number = 1
            tot_slides = ActivePresentation.Slides.Count
        ElseIf split_selected_slides = vbCancel Then
            Exit Sub
        End If
    Else
        slide_number = 1
        tot_slides = ActivePresentation.Slides.Count
    End If
    
    ProgressForm.SlideBar.value = 0
    ProgressForm.OverallBar.value = 0
    ProgressForm.Show
    
    If ActiveWindow.ViewType <> ppViewSlide And ActiveWindow.ViewType <> ppViewNormal Then
        ActiveWindow.ViewType = ppViewNormal
    End If
      
    ' Bake slide numbers (and other footers that may contain slide numbers) into the
    ' presentation, if requested.
    If slideNumbersAdjustMode <> SLIDENUMBER_DONOTHING Then bakeSlideNumbers slide_number, tot_slides
    
    ' Since lots of duplicate slides will be created in the process, I must
    ' keep note of:
    ' orig_tot_slides, which is the total number of slides in the selected
    ' range before creating duplicate slides
    orig_tot_slides = tot_slides
    ' actual_slide, which is the number of slides in the originally selected range
    ' that have been processed until now
    actual_slide = slide_number
    '
    ' Iterate over all the slides in the presentation
    '
    While actual_slide <= tot_slides
        additional_slide_present = False
        ProgressForm.SlideNumber = "Slide " + Str$(actual_slide) + " of " + Str$(orig_tot_slides)
        alreadyPurged = False
        ' Count of slides generated from splitting a single original one
        split_slides = 0
        If ActivePresentation.Slides(slide_number).TimeLine.MainSequence.Count > 0 Then
            '
            ' There are effects to be processed in the current slide
            '
            
            copyShapeIds ActivePresentation.Slides(slide_number)
            
            '
            ' First of all, take care of effects that start without a click
            ' (and, therefore, have an immediate effect on the rendered slide)
            '
            cont = (ActivePresentation.Slides(slide_number).TimeLine.MainSequence(1).Timing.TriggerType = msoAnimTriggerWithPrevious _
                    Or ActivePresentation.Slides(slide_number).TimeLine.MainSequence(1).Timing.TriggerType = msoAnimTriggerAfterPrevious)
            If cont And Not doNotSplitMouseTriggered Then
                ' Keep a copy of the original slide, which I will use to track the animation
                ' sequence. I always proceed in this way: I carry the original slide
                ' unaltered and grab the list of effects to be applied from it, while
                ' shapes are actually modified on copies of that original slide
                ActivePresentation.Slides(slide_number).Duplicate
                ' Remember to remove the duplicated slide later on
                additional_slide_present = True
                Set slide_timeline = ActivePresentation.Slides(slide_number + 1).TimeLine.MainSequence
                ' Remove all the shapes that will appear after a future entry effect
                purgeFutureShapes ActivePresentation.Slides(slide_number), False
                purgeEffects ActivePresentation.Slides(slide_number)
                alreadyPurged = True
            End If
            While cont And Not doNotSplitMouseTriggered
                ' Actually, there are animations that start without a click
                applyEffect ActivePresentation.Slides(slide_number), ActivePresentation.Slides(slide_number + 1)
                ' Some effects have disappeared: check whether I still have
                ' effects that start without a click
                If slide_timeline.Count = 0 Then
                    cont = False
                Else
                    ' Go on until I encounter a mouse-triggered effect
                    cont = (slide_timeline(1).Timing.TriggerType = msoAnimTriggerWithPrevious _
                            Or slide_timeline(1).Timing.TriggerType = msoAnimTriggerAfterPrevious)
                End If
            Wend
            If additional_slide_present Then
                ' Match the Z order of shapes between the original slide and its
                ' duplicate.
                matchZOrder ActivePresentation.Slides(slide_number), ActivePresentation.Slides(slide_number + 1)
            End If
        Else
            actual_slide = actual_slide + 1
        End If
            
        '
        ' Now, take care of mouse-triggered effects
        '
        ' Get the number of animation effects from the correct slide.
        If additional_slide_present Then
            tot_anims = ActivePresentation.Slides(slide_number + 1).TimeLine.MainSequence.Count
        Else
            tot_anims = ActivePresentation.Slides(slide_number).TimeLine.MainSequence.Count
        End If
        If tot_anims > 0 Then
            processed_anims = 0
            If Not alreadyPurged Then
                ActivePresentation.Slides(slide_number).Duplicate
                purgeFutureShapes ActivePresentation.Slides(slide_number), False
                purgeEffects ActivePresentation.Slides(slide_number)
                alreadyPurged = True
                
            End If
            ActivePresentation.Slides(slide_number).Duplicate
            split_slides = split_slides + 1
            augmentSlideNumbers slide_number, split_slides
            slide_number = slide_number + 1
            While ActivePresentation.Slides(slide_number + 1).TimeLine.MainSequence.Count > 0
                ' Mouse-triggered effects need to be split on two different slides
                ' Now iterate over all non-mouse-triggered effects starting with the current one
                cont = True
                While cont
                    ' The applyEffect method eats an animation effect for each call,
                    ' unless it returns 1.
                    addedEffects = applyEffect(ActivePresentation.Slides(slide_number), ActivePresentation.Slides(slide_number + 1))
                    
                    '
                    ' Ok, the current effect has been processed. Keep staying on the same slide
                    ' as long as there are other non-mouse-triggered effects.
                    '
                    Set slide_timeline = ActivePresentation.Slides(slide_number + 1).TimeLine.MainSequence
                    If slide_timeline.Count = 0 Then
                        ' No more effects to process (this must be checked on the next slide,
                        ' as several effects and shapes may have been removed in the current
                        ' one)
                        cont = False
                    Else
                        cont = (slide_timeline(1).Timing.TriggerType = msoAnimTriggerWithPrevious _
                                Or slide_timeline(1).Timing.TriggerType = msoAnimTriggerAfterPrevious) And Not doNotSplitMouseTriggered
                    End If
                    processed_anims = processed_anims + 1 - addedEffects
                    anims_percentage = Int(processed_anims / tot_anims * 100)
                    
                    ProgressForm.SlideLabel = Str$(anims_percentage) + " %"
                    ProgressForm.SlideBar.value = anims_percentage
                    ProgressForm.Repaint
                    DoEvents
                    If cancelStatus Then
                        Unload ProgressForm
                        Exit Sub
                    End If
                Wend
                matchZOrder ActivePresentation.Slides(slide_number), ActivePresentation.Slides(slide_number + 1)
                If slide_timeline.Count > 0 Then
                    ActivePresentation.Slides(slide_number).Duplicate
                    split_slides = split_slides + 1
                    augmentSlideNumbers slide_number, split_slides                    ' Try (hard) to trigger text auto-fit in PowerPoint <= 2003
                    For Each shape_object In ActivePresentation.Slides(slide_number + 1).Shapes
                        shape_object.Height = shape_object.Height + 0.01
                        shape_object.Height = shape_object.Height - 0.01
                    Next shape_object
                    purgeEffects ActivePresentation.Slides(slide_number)
                    slide_number = slide_number + 1
                Else
                    ' No more animations to process, but the last slide might still need some
                    ' touching of slide numbers
                    split_slides = split_slides + 1
                    augmentSlideNumbers slide_number, split_slides
                End If
            Wend
            ActivePresentation.Slides(slide_number + 1).Delete
            additional_slide_present = False
            ' All the animations for the current slide have been processed
            purgeEffects ActivePresentation.Slides(slide_number)
            actual_slide = actual_slide + 1
        End If      ' tot_anims > 0
        If additional_slide_present Then
            ActivePresentation.Slides(slide_number + 1).Delete
            purgeEffects ActivePresentation.Slides(slide_number)
            actual_slide = actual_slide + 1
        End If
        
        slide_number = slide_number + 1
        
        overall_percentage = Int((actual_slide - 1) / orig_tot_slides * 100)
        ProgressForm.OverallLabel = Str$(overall_percentage) + " %"
        ProgressForm.OverallBar = overall_percentage
        ProgressForm.SlideLabel = ""
        ProgressForm.SlideBar = 0
        ProgressForm.Repaint
        DoEvents
        If cancelStatus Then
            Unload ProgressForm
            Exit Sub
        End If
    Wend        ' actual_slide <= tot_slides
    
    Unload ProgressForm
    Exit Sub
    
error_handler:
    resp = MsgBox("Sorry, but despite the efforts in foreseeing and catching possible anomalies, I have incurred an unrecoverable error." & vbCrLf & _
                  "Error number: " & Str$(Err.Number) & vbCrLf & _
                  "Error description: " & Err.Description & vbCrLf & _
                  "Slide number: " & slide_number & vbCrLf & "Would you like to try continuing anyway (discouraged)?", vbYesNo, "Fatal error")
    If resp = vbYes Then
        Resume Next
    Else
        On Error GoTo 0
        Resume
    End If
End Sub

' The status of the "Slide numbers" dropdown
' has been changed
Sub changeAdjustSlideNumbersStatus()
    Dim myBar As CommandBar, myDropdown As CommandBarComboBox
    For Each b In CommandBars
        If b.Name = "PPspliT" Then
            Set myBar = b
        End If
    Next b
    Set myDropdown = myBar.Controls(4)
    slideNumbersAdjustMode = myDropdown.ListIndex
End Sub

' The status of the "split on mouse-triggered animations" button
' has been changed
Sub changeMouseSplitStatus()
    Dim myBar As CommandBar, myButton As CommandBarButton
    For Each b In CommandBars
        If b.Name = "PPspliT" Then
            Set myBar = b
        End If
    Next b
    Set myButton = myBar.Controls(2)
    If myButton.State = msoButtonDown Then
        myButton.State = msoButtonUp
    Else
        myButton.State = msoButtonDown
    End If
    doNotSplitMouseTriggered = (myButton.State = msoButtonUp)
End Sub

' Add the PPspliT toolbar, if not present
Sub auto_open()
    Dim a As AddIn
    For Each a In AddIns
        If a.Name = "PPspliT" Then
            aPath = a.Path
        End If
    Next a
    slideNumbersAdjustMode = SLIDENUMBER_BAKE
    splitMouseTriggered = False
    
    ' Any possibly existing command bar is replaced, to allow add-in
    ' upgrades that change its structure
    For Each b In CommandBars
        If b.Name = "PPspliT" Then b.Delete
    Next b
    
    ' (Re-)create the command bar
    Dim newBar As CommandBar, newButton As CommandBarButton, newDropdown As CommandBarComboBox
    Set newBar = CommandBars.Add(Name:="PPspliT", Position:=msoBarTop, temporary:=True)
    
    Set newButton = newBar.Controls.Add(msoControlButton)
    newButton.OnAction = "PPspliT_main"
    newButton.TooltipText = "Split animations"
    newButton.Caption = "Split animations"
    newButton.Style = msoButtonIconAndCaption
    If aPath <> "" Then
        newButton.Picture = LoadPicture(aPath + "\ppsplit-button.gif")
    End If
    
    Set newButton = newBar.Controls.Add(msoControlButton)
    newButton.OnAction = "changeMouseSplitStatus"
    newButton.TooltipText = "Split on click-triggered animation effects"
    newButton.Caption = "Split on click-triggered effects"
    If aPath <> "" Then
        newButton.Picture = LoadPicture(aPath + "\mouse-button.gif")
    End If
    newButton.State = msoButtonDown
    newButton.BeginGroup = True
    
    Set newButton = newBar.Controls.Add(msoControlButton)
    newButton.TooltipText = "How to handle slide numbers in slide footers"
    newButton.Caption = "Slide numbers:"
    newButton.Style = msoButtonCaption
    newButton.Enabled = False
    newButton.BeginGroup = True
    
    Set newDropdown = newBar.Controls.Add(msoControlDropdown)
    newDropdown.AddItem "Do nothing", 1
    newDropdown.AddItem "Preserve original", 2
    newDropdown.AddItem "Preserve, and add subindex", 3
    newDropdown.OnAction = "changeAdjustSlideNumbersStatus"
    newDropdown.TooltipText = "How to handle slide numbers in slide footers"
    newDropdown.Caption = "Slide number adjustment setting"
    newDropdown.Width = 160
    newDropdown.ListIndex = SLIDENUMBER_BAKE
    
    Set newButton = newBar.Controls.Add(msoControlButton)
    newButton.OnAction = "displayAboutForm"
    newButton.TooltipText = "About"
    newButton.Caption = "About"
    newButton.Style = msoButtonIconAndCaption
    If aPath <> "" Then
        newButton.Picture = LoadPicture(aPath + "\about-button.gif")
    End If
    newButton.BeginGroup = True
    
    newBar.Visible = True
End Sub

' Remove the PPspliT toolbar, if existing
Sub auto_close()
    Dim myBar As CommandBar
    foundBar = False
    For Each b In CommandBars
        If b.Name = "PPspliT" Then
            foundBar = True
            Set myBar = b
        End If
    Next b
    If foundBar Then
        myBar.Delete
    End If
End Sub

' Display the about form
Sub displayAboutForm()
    AboutForm.Show
End Sub
