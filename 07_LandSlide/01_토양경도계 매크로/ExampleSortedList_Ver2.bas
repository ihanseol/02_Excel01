Attribute VB_Name = "ExampleSortedList_Ver2"
Option Explicit


'http://www.eurus.dti.ne.jp/~yoneyama/Excel/vba/vba_sortedlist.html

Private Sub test_A1()

    Dim DataList As Object
    Dim x, i As Long
    Set DataList = CreateObject("System.Collections.SortedList")

    x = Range("B2:C8").Value
    For i = LBound(x) To UBound(x)
        If DataList.Contains(x(i, 1)) = False Then
            DataList.Add x(i, 1), x(i, 2)
        End If
    Next i
    For i = 0 To DataList.Count - 1
        Cells(i + 2, 5).Value = DataList.GetKey(i)
        Cells(i + 2, 6).Value = DataList.GetByIndex(i)
    Next i

    Set DataList = Nothing

End Sub

Private Sub test_B1()
    Dim i As Long
    Dim DataList As Object
    Dim x
        Set DataList = CreateObject("System.Collections.SortedList")
        Randomize

        For i = 1 To 10
            DataList.item(Rnd()) = i
        Next i

        For i = 0 To DataList.Count - 1
            Cells(i + 1, 1).Value = DataList.GetByIndex(i)
            Cells(i + 1, 2).Value = DataList.GetKey(i)
        Next i

        Set DataList = Nothing

End Sub

Private Sub test_B2()
    Dim i As Long
    Dim DataList As Object
    Dim x
        Set DataList = CreateObject("System.Collections.SortedList")

        x = Range("A1��A10").Value
        For i = LBound(x) To UBound(x)
            DataList.item(x(i, 1)) = ""
        Next i

        For i = 0 To DataList.Count - 1
            Cells(i + 1, 3).Value = DataList.GetKey(i)
        Next i

        Set DataList = Nothing

End Sub


Private Sub test_C11()
    Dim i As Long
    Dim DataList As Object
    Dim x
        Set DataList = CreateObject("System.Collections.SortedList")

        x = Range("A1��B5").Value
        For i = LBound(x) To UBound(x)
            If DataList.Contains(x(i, 1)) = False Then
                DataList.Add x(i, 1), x(i, 2)
            End If
        Next i

        '��?��������ʥ����
        DataList.Add "�ޫ�?", 6

        '������
        Range("D:E").ClearContents
        For i = 0 To DataList.Count - 1
            Cells(i + 1, 4).Value = DataList.GetKey(i)
            Cells(i + 1, 5).Value = DataList.GetByIndex(i)
        Next i

        Set DataList = Nothing

End Sub



Private Sub test_C12()
    Dim i As Long
    Dim DataList As Object
    Dim x
        Set DataList = CreateObject("System.Collections.SortedList")

        x = Range("A1��B5").Value
        For i = LBound(x) To UBound(x)
            If DataList.Contains(x(i, 1)) = False Then
                DataList.Add x(i, 1), x(i, 2)
            End If
        Next i

         '���ꫢ����
        DataList.Clear

        '��?��������ʥ����
        DataList.Add "�ޫ�?", 6

        '������
        Range("D:E").ClearContents
        For i = 0 To DataList.Count - 1
            Cells(i + 1, 4).Value = DataList.GetKey(i)
            Cells(i + 1, 5).Value = DataList.GetByIndex(i)
        Next i

        Set DataList = Nothing

End Sub

'Test Japn Code page ...

Private Sub test_C13()
    Dim i As Long
    Dim DataList As Object
    Dim x
        Set DataList = CreateObject("System.Collections.SortedList")

        x = Range("A1��B5").Value
        For i = LBound(x) To UBound(x)
            If DataList.Contains(x(i, 1)) = False Then
                DataList.Add x(i, 1), x(i, 2)
            End If
        Next i

      '��Ȫ���𶪹��
        DataList.Remove ("���")

        '������
        Range("D:E").ClearContents
        For i = 0 To DataList.Count - 1
            Cells(i + 1, 4).Value = DataList.GetKey(i)
            Cells(i + 1, 5).Value = DataList.GetByIndex(i)
        Next i

        Set DataList = Nothing

End Sub





