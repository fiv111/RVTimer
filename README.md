# RTTimer
A simple Timer.

## Quick start
```
Public Sub test()
  Dim t As RVTimer
  Set t = New RVTimer

  ' time setting.
  t.startTime    = "11:10:00"
  t.endTime      = "11:25:00"
  t.setTime      = "11:24:00"
  t.intervalTime = "00:00:10"
  t.repeat       = True

  ' callback function setting.
  t.cbObject   = someObject
  t.cbFuncName = "someMethodName"
  t.cbType     = VbMethod
  t.cbArgList.Add "argument0"

  ' run script.
  t.run

  Set t = Nothing
End Sub
```

## Usage

### RVTimer Class

#### time
```
' from time
startTime = "00:00:00"

' to time
endTime = "00:10:00"

' Call out the `callback` function when, nowTime as same as setTime
setTime = "00:04:00"

' Call `callback` within a certain period of time.
intervalTime = "00:01:00"

' If your wait to repeat the job, set `true`.
repeat = True

' A pointer to the object.
cbObject = myObject

' The name of the property or method on the object.
cbFuncName = "myFunc"

' A type which is representing the type of procedure being called.
cbType = VbMethod

' A parameter ArrayList which is passed to the property or method being called.
cbArgList.add "arg1"
cbArgList.add "arg2"
```

#### Property

##### Let
```
Public Property Let startTime(ByVal v As String)
Public Property Let endTime(ByVal v As String)
Public Property Let setTime(ByVal v As String)
Public Property Let intervalTime(ByVal v As String)
Public Property Let repeat(ByVal v As Boolean)
Public Property Let cbObject(ByVal v As Object)
Public Property Let cbFuncName(ByVal v As String)
Public Property Let cbType(ByVal v As Integer)
```

##### Get
```
Public Property Get startTime() As String
Public Property Get endTime() As String
Public Property Get setTime() As String
Public Property Get nowTime() As Long
Public Property Get intervalTime() As String
Public Property Get repeat() As Boolean
Public Property Get cbObject() As Object
Public Property Get cbFuncName() As String
Public Property Get cbType() As Integer
Public Property Get cbArgList() As Variant
```

#### Method
```
Public Function strtime2sec(ByVal strtime As String) As Double
Public Function sec2time(ByVal s As Double) As String
public Sub run()
```