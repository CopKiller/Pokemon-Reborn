Attribute VB_Name = "modIncubator"
Option Explicit

Public Sub RequestIncubator()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestIncubator
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub
