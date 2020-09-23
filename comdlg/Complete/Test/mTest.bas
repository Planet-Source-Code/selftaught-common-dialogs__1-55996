Attribute VB_Name = "mTest"
Option Explicit

Public Function OrItemData(ByVal lst As ListBox) As Long
    Dim i As Long
    For i = 0& To lst.ListCount - 1&
        If lst.Selected(i) Then OrItemData = OrItemData Or lst.ItemData(i)
    Next
End Function
