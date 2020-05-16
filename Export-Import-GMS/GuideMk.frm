VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GuideMk 
   Caption         =   "Guide Maker v2.2"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2535
   OleObjectBlob   =   "GuideMk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GuideMk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chk As Boolean
Dim gml As Layer 'guide-master layer


Private Sub ChkBoxBleeds_Click()
    If chk Then Exit Sub '
    Dim w As Double, h As Double
    Dim b As Double
    Dim x1 As Double, y1 As Double
    
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter
    
    b = Txtbleeds.Value
    
    w = ActivePage.SizeWidth
    h = ActivePage.SizeHeight

    x1 = ActivePage.CenterX - ActivePage.SizeWidth / 2
    y1 = ActivePage.CenterY - ActivePage.SizeHeight / 2
    
    If ChkBoxBleeds.Value = True Then
        gml.CreateGuide x1 - b, y1, x1 - b, y1 + h 'x1, y1, x2, y2
        gml.CreateGuide x1 + w + b, y1, x1 + w + b, y1 + h
        gml.CreateGuide x1, y1 - b, x1 + w, y1 - b
        gml.CreateGuide x1, y1 + h + b, x1 + w, y1 + h + b
    Else
        'left, right, top, bottom, facing, operation
        AddGuidesToRange x1 - b, x1 + w + b, y1 + h + b, y1 - b, "del"
    End If
    
End Sub

Private Sub ChkBoxDoc_Click()
    If chk Then Exit Sub
    Dim h As Double
    Dim w As Double
    Dim x1 As Double, y1 As Double
    
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter
    
    w = ActivePage.SizeWidth
    h = ActivePage.SizeHeight

    x1 = ActivePage.CenterX - ActivePage.SizeWidth / 2
    y1 = ActivePage.CenterY - ActivePage.SizeHeight / 2
    
    If ChkBoxDoc.Value = True Then
        gml.CreateGuide x1, y1, x1, y1 + h 'x1, y1, x2, y2
        gml.CreateGuide x1 + w, y1, x1 + w, y1 + h
        gml.CreateGuide x1, y1, x1 + w, y1
        gml.CreateGuide x1, y1 + h, x1 + w, y1 + h
    Else
        'left, right, top, bottom, operation
        AddGuidesToRange x1, x1 + w, y1 + h, y1, "del"
    End If

End Sub

Private Sub ChkBoxFields_Click()
    If chk Then Exit Sub
    Dim h As Double
    Dim w As Double
    Dim f As Double
    Dim x1 As Double, y1 As Double
    
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter
    
    f = Txtfilds.Value
    
    w = ActivePage.SizeWidth
    h = ActivePage.SizeHeight

    x1 = ActivePage.CenterX - ActivePage.SizeWidth / 2
    y1 = ActivePage.CenterY - ActivePage.SizeHeight / 2
    
    If ChkBoxFields.Value = True Then
        gml.CreateGuide x1 + f, y1, x1 + f, y1 + h 'x1, y1, x2, y2
        gml.CreateGuide x1 + w - f, y1, x1 + w - f, y1 + h
        gml.CreateGuide x1, y1 + f, x1 + w, y1 + f
        gml.CreateGuide x1, y1 + h - f, x1 + w, y1 + h - f
    Else
        AddGuidesToRange x1 + f, x1 + w - f, y1 + h - f, y1 + f, "del"
    End If

End Sub

Private Sub CmbDelAll_Click()
    Dim sr As New ShapeRange
    For i = 1 To ActiveDocument.Pages.Count
        sr.AddRange ActiveDocument.Pages(i).FindShapes(Type:=cdrGuidelineShape)
    Next i
    sr.Delete
    ChkBoxDoc.Value = False
    ChkBoxFields.Value = False
    ChkBoxBleeds.Value = False
End Sub

Private Sub SpinBtnBleed_SpinDown()
    Txtbleeds.Value = SpinBtnBleed.Value
    CheckGuides
End Sub

Private Sub SpinBtnBleed_SpinUp()
    Txtbleeds.Value = SpinBtnBleed.Value
    CheckGuides
End Sub

Private Sub SpinBtnField_SpinDown()
    Txtfilds.Value = SpinBtnField.Value
    CheckGuides
End Sub

Private Sub SpinBtnField_SpinUp()
    Txtfilds.Value = SpinBtnField.Value
    CheckGuides
End Sub

Private Sub UserForm_Initialize()
    For i = 1 To ActiveDocument.MasterPage.AllLayers.Count
        If ActiveDocument.MasterPage.AllLayers.Item(i).IsGuidesLayer Then
            Set gml = ActiveDocument.MasterPage.AllLayers.Item(i)
            Exit For
        End If
    Next
    
    If ActiveDocument.FacingPages = False Then
        If ActivePage.BoundingBox.Left <> 0 Then
            MsgBox ("Rulers are shifted!" & vbCrLf)
        End If
    End If

    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter
    
    CheckGuides

End Sub

Private Sub CheckGuides()
    Dim h As Double
    Dim w As Double
    Dim f As Double
    Dim b As Double
    Dim inf As String
    Dim x1 As Double, y1 As Double
    
    f = Val(Txtfilds.Value)
    b = Val(Txtbleeds.Value)

    w = ActivePage.SizeWidth
    h = ActivePage.SizeHeight
    
    x1 = ActivePage.CenterX - ActivePage.SizeWidth / 2
    y1 = ActivePage.CenterY - ActivePage.SizeHeight / 2
    
    SpinBtnField.Value = f
    SpinBtnBleed.Value = b
    
    chk = True
    'check doc
    inf = "con" 'con - contain
    AddGuidesToRange x1, x1 + w, y1 + h, y1, inf
    If inf = "yes" Then
        ChkBoxDoc.Value = True
    Else
        ChkBoxDoc.Value = False
    End If
    
    'check fields
    inf = "con"
    AddGuidesToRange x1 + f, x1 + w - f, y1 + h - f, y1 + f, inf
    If inf = "yes" Then
        ChkBoxFields.Value = True
    Else
        ChkBoxFields.Value = False
    End If
    
     'check bleeds
    inf = "con"
    AddGuidesToRange x1 - b, x1 + w + b, y1 + h + b, y1 - b, inf
    If inf = "yes" Then
        ChkBoxBleeds.Value = True
    Else
        ChkBoxBleeds.Value = False
    End If

    chk = False

End Sub

Private Sub AddGuidesToRange(lf As Double, rt As Double, tp As Double, _
                                bt As Double, ByRef inf As String)
    Dim sr As New ShapeRange
    Dim lf_b As Boolean, rt_b As Boolean, tp_b As Boolean, bt_b As Boolean
    
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter
    
    For i = 1 To gml.Shapes.Count
        If gml.Shapes(i).BoundingBox.Width < 0.001 Then
            If gml.Shapes(i).BoundingBox.CenterX < lf + 0.1 And _
                gml.Shapes(i).BoundingBox.CenterX > lf - 0.1 Then
                If inf = "del" Then sr.Add gml.Shapes(i)
                lf_b = True
            End If
            If gml.Shapes(i).BoundingBox.CenterX < rt + 0.1 And _
                gml.Shapes(i).BoundingBox.CenterX > rt - 0.1 Then
                If inf = "del" Then sr.Add gml.Shapes(i)
                rt_b = True
            End If
        End If
        If gml.Shapes(i).BoundingBox.Height < 0.001 Then
            If gml.Shapes(i).BoundingBox.CenterY < tp + 0.1 And _
                gml.Shapes(i).BoundingBox.CenterY > tp - 0.1 Then
                If inf = "del" Then sr.Add gml.Shapes(i)
                tp_b = True
            End If
            If gml.Shapes(i).BoundingBox.CenterY < bt + 0.1 And _
                gml.Shapes(i).BoundingBox.CenterY > bt - 0.1 Then
                If inf = "del" Then sr.Add gml.Shapes(i)
                bt_b = True
            End If
        End If
    Next
    If inf = "del" Then sr.Delete
    If inf = "con" Then
        If (lf_b = True And rt_b = True And tp_b = True And bt_b = True) Then
            inf = "yes"
        Else
            inf = "no"
        End If
    End If
End Sub
