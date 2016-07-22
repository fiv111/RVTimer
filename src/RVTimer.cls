VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RVTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' set sleep function
#If VBA7 Then
  ' 64 Bit
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr)
#Else
  ' 32 Bit
  Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#End If


' ---
' Const
' ---
Private Const kMillsec = 1000
Private Const kTimeoutMsg = "Out of Range. Bye."


' ---
' Property
' ---
' from time
Private startTime_ As Long
' to time
Private endTime_ As Long
' Call out the `callback` function when, nowTime as same as setTime
Private setTime_ As Long
' Current Time. Use `Timer`
Private nowTime_ As Long
' If your wait to repeat the job, set `true`.
Private repeat_ As Boolean
' Call `callback` within a certain period of time.
Private intervalTime_ As Long
' fps
Private fps_ As Double

' A pointer to the object.
Private cbObject_ As Object
' The name of the property or method on the object.
Private cbFuncName_ As String
' A type which is representing the type of procedure being called.
Private cbType_ As Integer
' A parameter ArrayList which is passed to the property or method being called.
Private cbArgList_ As Object


Public Property Let startTime(ByVal v As String)
  startTime_ = strtime2sec(v)
End Property

Public Property Let endTime(ByVal v As String)
  endTime_ = strtime2sec(v)
End Property

Public Property Let setTime(ByVal v As String)
  setTime_ = strtime2sec(v)
End Property

Private Property Let nowTime(ByVal v As Long)
  nowTime_ = Math.Round(v)
End Property

Public Property Let intervalTime(ByVal v As String)
  intervalTime_ = strtime2sec(v)
End Property

Public Property Let repeat(ByVal v As Boolean)
  repeat_ = v
End Property

Public Property Let fps(ByVal v As Double)
  fps_ = v
End Property

Public Property Let cbObject(ByVal v As Object)
  Set cbObject_ = v
End Property

Public Property Let cbFuncName(ByVal v As String)
  cbFuncName_ = v
End Property

Public Property Let cbType(ByVal v As Integer)
  cbType_ = v
End Property


Public Property Get startTime() As String
  startTime = sec2time(startTime_)
End Property

Public Property Get endTime() As String
  endTime = sec2time(endTime_)
End Property

Public Property Get setTime() As String
  setTime = sec2time(setTime_)
End Property

Public Property Get nowTime() As Long
  nowTime = nowTime_
End Property

Public Property Get intervalTime() As String
  intervalTime = sec2time(intervalTime_)
End Property

Public Property Get repeat() As Boolean
  repeat = repeat_
End Property

Public Property Get fps() As Double
  fps = fps_
End Property

Public Property Get cbObject() As Object
  Set cbObject = cbObject_
End Property

Public Property Get cbFuncName() As String
  cbFuncName = cbFuncName_
End Property

Public Property Get cbType() As Integer
  cbType = cbType_
End Property

Public Property Get cbArgList() As Variant
  Set cbArgList = cbArgList_
End Property


' ---
' Init
' ---
Private Sub Class_Initialize()
  Set cbArgList_ = CreateObject("System.Collections.ArrayList")
  cbType = VbMethod
  nowTime = Timer
  repeat = False
  intervalTime = "00:01:00"
  fps = 1
End Sub

Private Sub Class_Terminate()
  Set cbArgList_ = Nothing
End Sub


' ---
' Method
' ---
Private Function timeObject(Optional ByVal val As String = "00:00:00") As Object
  Dim o As Object
  Dim tmp() As String
  
  Set o = CreateObject("Scripting.Dictionary")
  
  tmp = Split(val, ":")
  
  o.Add "h", CDbl(tmp(0))
  o.Add "m", CDbl(tmp(1))
  o.Add "s", CDbl(tmp(2))
  
  Set timeObject = o
  Set o = Nothing
End Function

' Convert time value to second. (Long)
Private Function time2sec(ByVal h As Double, ByVal m As Double, ByVal s As Double) As Double
  Dim hourPerSec As Double
  Dim minPerSec As Double
  hourPerSec = 60 * 60
  minPerSec = 60
  time2sec = (h * hourPerSec) + (m * minPerSec) + s
End Function

' Convert string time value to second.
Public Function strtime2sec(ByVal strtime As String) As Double
  Dim t As Object
  Set t = timeObject(strtime)
  strtime2sec = time2sec(t("h"), t("m"), t("s"))
  Set t = Nothing
End Function

' Convert second to time string.
Public Function sec2time(ByVal s As Double) As String
  sec2time = Application.WorksheetFunction.Text(CDate(s / 86400#), "[h]:mm:ss")
End Function


' Callback function.
Private Sub callback()
  Dim size As Integer
  size = cbArgList.Count

  Select Case size
  Case 1
    CallByName cbObject, cbFuncName, cbType, cbArgList.Item(0)
  Case 2
    CallByName cbObject, cbFuncName, cbType, cbArgList.Item(0), cbArgList.Item(1)
  Case 3
    CallByName cbObject, cbFuncName, cbType, cbArgList.Item(0), cbArgList.Item(1), cbArgList.Item(2)
  Case 4
    CallByName cbObject, cbFuncName, cbType, cbArgList.Item(0), cbArgList.Item(1), cbArgList.Item(2), cbArgList.Item(3)
  Case 5
    CallByName cbObject, cbFuncName, cbType, cbArgList.Item(0), cbArgList.Item(1), cbArgList.Item(2), cbArgList.Item(3), cbArgList.Item(4)
  Case 6
    CallByName cbObject, cbFuncName, cbType, cbArgList.Item(0), cbArgList.Item(1), cbArgList.Item(2), cbArgList.Item(3), cbArgList.Item(4), cbArgList.Item(5)
  Case 7
    CallByName cbObject, cbFuncName, cbType, cbArgList.Item(0), cbArgList.Item(1), cbArgList.Item(2), cbArgList.Item(3), cbArgList.Item(4), cbArgList.Item(5), cbArgList.Item(6)
  Case 8
    CallByName cbObject, cbFuncName, cbType, cbArgList.Item(0), cbArgList.Item(1), cbArgList.Item(2), cbArgList.Item(3), cbArgList.Item(4), cbArgList.Item(5), cbArgList.Item(6), cbArgList.Item(7)
  Case 9
    CallByName cbObject, cbFuncName, cbType, cbArgList.Item(0), cbArgList.Item(1), cbArgList.Item(2), cbArgList.Item(3), cbArgList.Item(4), cbArgList.Item(5), cbArgList.Item(6), cbArgList.Item(7), cbArgList.Item(8)
  Case 10
    CallByName cbObject, cbFuncName, cbType, cbArgList.Item(0), cbArgList.Item(1), cbArgList.Item(2), cbArgList.Item(3), cbArgList.Item(4), cbArgList.Item(5), cbArgList.Item(6), cbArgList.Item(7), cbArgList.Item(8), cbArgList.Item(9)
  Case Default
    CallByName cbObject, cbFuncName, cbType
  End Select
End Sub


' run
Public Sub run()
  ' timeout
  If nowTime < startTime_ Or nowTime >= endTime_ Then
    'MsgBox kTimeoutMsg
    Exit Sub
  End If
  
  Do While 1
    nowTime = Math.Round(Timer)
    
    If nowTime < startTime_ Or nowTime >= endTime_ Then
      'MsgBox kTimeoutMsg
      Exit Do
    End If
    
    If nowTime = setTime_ Then
      Call callback
      
      If repeat Then
        setTime_ = setTime_ + intervalTime_
      End If
    End If
    
    Sleep (kMillsec / fps)
    DoEvents
  Loop
End Sub
