Attribute VB_Name = "�߰����"
Sub Auto_Open()
 
    Auto_Close
    
    On Error Resume Next
    
    With Application.CommandBars("Tools").Controls
    
        With .Add(Type:=msoControlButton)
            .FaceId = 8
            .Caption = "�޷� �����"
            .OnAction = "Calendar"
        End With
    
        With .Add(Type:=msoControlButton)
            .FaceId = 9644
            .Caption = "���� �����ϱ�"
            .OnAction = "ExportSchedule"
        End With
    
        With .Add(Type:=msoControlButton)
            .FaceId = 47
            .Caption = "�޷� �ʱ�ȭ"
            .OnAction = "ClearCalendar"
        End With
        
        With .Add(Type:=msoControlButton)
            .FaceId = 3
            .Caption = "���� ��Ʈ ����"
            .OnAction = "mkFile"
        End With
    
    End With
    
    On Error GoTo 0
    
End Sub
 
Sub Auto_Close()
 
    Application.CommandBars("Tools").Reset
    
End Sub
