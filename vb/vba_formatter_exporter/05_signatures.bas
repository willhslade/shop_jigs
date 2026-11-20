Public Sub NoParams()
End Sub

Public Sub OneParam(ByVal a As Long)
End Sub

Public Sub TwoParams(ByVal a As Long, ByVal b As String)
End Sub

Public Function WithReturn(ByVal a As Long, Optional ByVal b As String = "x,y(z)", ParamArray rest() As Variant) As Boolean
End Function

Private Function VeryLongSignature(ByVal FirstParameterName As String, ByVal SecondParameterName As Long, ByRef ThirdParameterName As Variant, Optional ByVal FourthParameter As Boolean = False) As Collection
End Function

Public Sub AlreadyWrapped( _
    ByVal a As Long, _
    ByVal b As Long, _
    ByVal c As Long)
End Sub

Public Property Get Item(ByVal Index As Long) As Variant
End Property

Public Property Let Item(ByVal Index As Long, ByVal Value As Variant)
End Property

Public Property Set Obj(ByVal Value As Object)
End Property

Friend Sub FriendSub(ByVal a As Long)
End Sub

Static Sub StaticSub(ByVal a As Long)
End Sub

Public Sub WithComment(ByVal a As Long, ByVal b As Long) ' trailing comment on signature
End Sub

Public Function ManyParams(ByVal p01 As Long, ByVal p02 As Long, ByVal p03 As Long, ByVal p04 As Long, ByVal p05 As Long, ByVal p06 As Long, ByVal p07 As Long, ByVal p08 As Long, ByVal p09 As Long, ByVal p10 As Long, ByVal p11 As Long, ByVal p12 As Long, ByVal p13 As Long, ByVal p14 As Long, ByVal p15 As Long, ByVal p16 As Long, ByVal p17 As Long, ByVal p18 As Long, ByVal p19 As Long, ByVal p20 As Long, ByVal p21 As Long, ByVal p22 As Long, ByVal p23 As Long, ByVal p24 As Long, ByVal p25 As Long, ByVal p26 As Long, ByVal p27 As Long, ByVal p28 As Long, ByVal p29 As Long, ByVal p30 As Long, ByVal p31 As Long, ByVal p32 As Long, ByVal p33 As Long, ByVal p34 As Long, ByVal p35 As Long, ByVal p36 As Long) As Long
End Function

Public Sub ArrayParams(ByRef a() As Long, ByRef b(1 To 3) As String, Optional ByVal c As Variant)
End Sub
