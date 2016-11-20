Attribute VB_Name = "Module1"
'I was asked to write this small piece of code which
'halves the size of the selected image in excel.
'No error handling
'SizePerc is size factor





Sub ShrinkImage()

    Dim MyPic As Shape
    Dim UserSelection As Variant
    Dim SizePerc As Double
    SizePerc = 0.5

    Set UserSelection = ActiveWindow.Selection

    Set MyPic = ActiveSheet.Shapes(UserSelection.Name)

    MyPic.LockAspectRatio = msoTrue

    MyPic.Width = MyPic.Width * SizePerc
    
End Sub



