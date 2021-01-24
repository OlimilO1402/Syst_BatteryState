VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PowerCaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SystemPowerCaps As SystemPowerCapabilities

Private Sub Class_Initialize()
    m_SystemPowerCaps = MPower.GetSystemPowerCapabilities
End Sub

Public Function ToStr() As String
    Dim s As String
    With m_SystemPowerCaps
        s = "PowerCapabilities" & vbCrLf & "{" & vbCrLf
        s = s & "    PowerButtonPresent     : " & CBool(.boolPowerButtonPresent) & ";" & vbCrLf
        s = s & "    SleepButtonPresent     : " & CBool(.boolSleepButtonPresent) & ";" & vbCrLf
        s = s & "    LidPresent             : " & CBool(.boolLidPresent) & ";" & vbCrLf
        s = s & "    SystemS1               : " & CBool(.boolSystemS1) & ";" & vbCrLf
        s = s & "    SystemS2               : " & CBool(.boolSystemS2) & ";" & vbCrLf
        s = s & "    SystemS3               : " & CBool(.boolSystemS3) & ";" & vbCrLf
        s = s & "    SystemS4               : " & CBool(.boolSystemS4) & ";" & vbCrLf
        s = s & "    SystemS5               : " & CBool(.boolSystemS5) & ";" & vbCrLf
        s = s & "    HiberFilePresent       : " & CBool(.boolHiberFilePresent) & ";" & vbCrLf
        s = s & "    FullWake               : " & CBool(.boolFullWake) & ";" & vbCrLf
        s = s & "    VideoDimPresent        : " & CBool(.boolVideoDimPresent) & ";" & vbCrLf
        s = s & "    ApmPresent             : " & CBool(.boolApmPresent) & ";" & vbCrLf
        s = s & "    UpsPresent             : " & CBool(.boolUpsPresent) & ";" & vbCrLf
        s = s & "    ThermalControl         : " & CBool(.boolThermalControl) & ";" & vbCrLf
        s = s & "    ProcessorThrottle      : " & CBool(.boolProcessorThrottle) & ";" & vbCrLf

        s = s & "    ProcessorMinimumThrottle : " & .ProcessorMinimumThrottle & ";" & vbCrLf
        s = s & "    ProcessorMaximumThrottle : " & .ProcessorMaximumThrottle & ";" & vbCrLf
        s = s & "    FastSystemS4           : " & CBool(.boolFastSystemS4) & ";" & vbCrLf
        'spare2(1 To 3)
        s = s & "    boolDiskSpinDown       : " & CBool(.boolDiskSpinDown) & ";" & vbCrLf
        'spare3(1 To 8)

        s = s & "    SystemBatteriesPresent : " & CBool(.boolSystemBatteriesPresent) & ";" & vbCrLf
        s = s & "    BatteriesAreShortTerm  : " & CBool(.boolBatteriesAreShortTerm) & ";" & vbCrLf
        'BatteryScale(1 To 3)

        s = s & "    AcOnlineWake           : " & SystemPowerState_ToStr(.AcOnlineWake) & ";" & vbCrLf
        s = s & "    SoftLidWake            : " & SystemPowerState_ToStr(.SoftLidWake) & ";" & vbCrLf
        s = s & "    RtcWake                : " & SystemPowerState_ToStr(.RtcWake) & ";" & vbCrLf
        s = s & "    MinimumDeviceWakeState : " & SystemPowerState_ToStr(.MinimumDeviceWakeState) & ";" & vbCrLf
        s = s & "    DefaultLowLatencyWake  : " & SystemPowerState_ToStr(.DefaultLowLatencyWake) & ";" & vbCrLf & "}"
    End With
    ToStr = s
End Function
