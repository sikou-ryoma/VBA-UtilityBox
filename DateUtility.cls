VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'日時設定のユーティリティクラス

Private tgtY As Long 'targetYear
Private tgtM As Long 'targetMonth
Private tgtD As Long 'targetDay

'インスタンス生成時に現在の日付の情報が入る
Private Sub class_initialize()
    tgtY = Year(Date)
    tgtM = Month(Date)
    tgtD = Day(Date)
End Sub

'年(西暦4桁)の設定・取得
Public Property Let targetYear(ByVal inputYear As Long)
    tgtY = inputYear
End Property

Public Property Get targetYear() As Long
    targetYear = tgtY
End Property

'月の設定・取得
Public Property Let targetMonth(ByVal inputMonth As Long)
    tgtM = inputMonth
End Property

Public Property Get targetMonth() As Long
    targetMonth = tgtM
End Property

'日の設定・取得
Public Property Let targetDay(ByVal inputDay As Long)
    tgtD = inputDay
End Property

Public Property Get targetDay() As Long
    targetDay = tgtD
End Property

'設定した年月より月初日のシリアル値を取得
Public Property Get FirstDayOfMonth() As Date
    FirstDayOfMonth = DateSerial(tgtY, tgtM, 1)
End Property

'設定した年月より月末日のシリアル値を取得
Public Property Get LastDayOfMonth() As Date
    LastDayOfMonth = DateSerial(tgtY, tgtM + 1, 1) - 1
End Property

'引数に指定したシリアル値の年月より月初日のシリアル値を取得
Public Function FirstDayOfMonthFromVal(ByVal dt As Date) As Date
    FirstDayOfMonthFromVal = DateSerial(Year(dt), Month(dt), 1)
End Function

'引数に指定したシリアル値の年月より月末日のシリアル値を取得
Public Property Get LastDayOfMonthFromVal(ByVal dt As Date) As Date
    LastDayOfMonthFromVal = DateSerial(Year(dt), Month(dt) + 1, 1) - 1
End Property

'設定した年月日よりその日のシリアル値を取得
Public Property Get targetDate() As Date
    targetDate = DateSerial(tgtY, tgtM, tgtD)
End Property


