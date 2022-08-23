VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "joueur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private une_main As hand
Private name As String
Private number As Integer
Private position As String
Private statut As Boolean
Private stack As Long
Private action As String
Private mise As Long

Property Let let_une_main(main_ As hand)
    Set une_main = main_
End Property

Property Let let_name(name_ As String)
    name = name_
End Property

Property Let let_number(number_ As Integer)
    number = number_
End Property

Property Let let_position(position_ As String)
    position = position_
End Property

Property Let let_statut(statut_ As Boolean)
    statut = statut_
End Property

Property Let let_action(action_ As String)
    action = action_
End Property

Property Let let_stack(stack_ As Long)
    stack = stack_
End Property

Property Let let_mise(mise_ As Long)
    mise = mise_
End Property

Property Get get_une_main()
    Set get_une_main = une_main
End Property

Property Get get_name()
    get_name = name
End Property

Property Get get_number()
    get_number = number
End Property

Property Get get_position()
    get_position = position
End Property

Property Get get_statut()
    If stack > 0 Then
        get_statut = 1
    Else
        get_statut = 0
    End If
End Property

Property Get get_action()
    get_action = action
End Property

Property Get get_stack()
    get_stack = stack
End Property

Property Get get_mise()
    get_mise = mise
End Property

Function miser(mise_ As Long)

Dim relance As Long
relance = mise_ - mise

If stack - relance <= 0 Then
    mise = stack + mise
    stack = 0
Else
    stack = stack - relance
    mise = mise_
End If

End Function

