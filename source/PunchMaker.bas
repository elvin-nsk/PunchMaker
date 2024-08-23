Attribute VB_Name = "PunchMaker"
'===============================================================================
'   Макрос          : PunchMaker
'   Версия          : 2024.08.23
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "PunchMaker"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2024.08.23"

'===============================================================================
' # Globals

Public Const GROOVE_SIZE As Double = 4 'диаметр выреза, мм
Public Const GROOVE_PUNCH_LENGTH As Double = 20
Public Const MAX_ANGLE = 80 'максимальный угол для выреза, градусов
Public Const VALID_OUTLINE_COLOR As String = "CMYK,USER,0,0,0,100"

'===============================================================================
' # Entry points

Sub Start()
    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
       
    Dim Shapes As ShapeRange
    With InputData.RequestShapes
        If .IsError Then Exit Sub
        Set Shapes = .Shapes
    End With
    
    BoostStart APP_DISPLAYNAME
    
    Simplify Shapes
    FilterValid Shapes
    
    If Shapes.Count = 0 Then
        VBA.MsgBox "Не найдены фигуры с подходящим цветом обводки"
        GoTo Finally
    End If
    If Not CheckShapesHasCurves(Shapes) Then
        VBA.MsgBox "Объекты должны быть простыми кривыми"
        GoTo Finally
    End If
    
    With New Grooves
        Set .Shapes = Shapes
        .GrooveSize = GROOVE_SIZE
        .PunchLength = GROOVE_PUNCH_LENGTH
        .ProbeRadius = GROOVE_SIZE / 10
        .ProbeSteps = 36
        .ConcavityMult = 1 - (1 / 360 * MAX_ANGLE)
        .MakeGrooves
    End With
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

'===============================================================================
' # Helpers

Private Function CheckShapesHasCurves(ByVal Shapes As ShapeRange) As Boolean
    CheckShapesHasCurves = True
    Dim Shape As Shape
    For Each Shape In Shapes
        If Not HasCurve(Shape) Then
            CheckShapesHasCurves = False
            Exit Function
        End If
    Next Shape
End Function

Private Function FilterValid(ByRef Shapes As ShapeRange)
    Dim Result As New ShapeRange
    Dim ValidColor As Color: Set ValidColor = CreateColor(VALID_OUTLINE_COLOR)
    Dim Shape As Shape
    For Each Shape In Shapes
        If IsSome(Shape.Outline) Then
            If Shape.Outline.Color.IsSame(ValidColor) Then Result.Add Shape
        End If
    Next Shape
    Set Shapes = Result
End Function

'===============================================================================
' # Tests

Private Sub testSimplify()
    Simplify ActivePage.Shapes.All
End Sub
