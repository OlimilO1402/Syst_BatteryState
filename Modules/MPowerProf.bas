Attribute VB_Name = "MPower"
Option Explicit
'Microsoft.WindowsAPICodePack.ApplicationServices
'Microsoft.WindowsAPICodePack.ApplicationServices.PowerManagementNativeMethods
'internal static class Power
'internal static class PowerManagementNativeMethods

Public Enum PowerInformationLevel
    SystemPowerPolicyAc = 0
    SystemPowerPolicyDc
    VerifySystemPolicyAc
    VerifySystemPolicyDc
    SystemPowerCapabilities
    SystemBatteryState
    SystemPowerStateHandler
    ProcessorStateHandler
    SystemPowerPolicyCurrent
    AdministratorPowerPolicy
    SystemReserveHiberFile
    ProcessorInformation
    SystemPowerInformation
    ProcessorStateHandler2
    LastWakeTime
    LastSleepTime
    SystemExecutionState
    SystemPowerStateNotifyHandler
    ProcessorPowerPolicyAc
    ProcessorPowerPolicyDc
    VerifyProcessorPowerPolicyAc
    VerifyProcessorPowerPolicyDc
    ProcessorPowerPolicyCurrent
    SystemPowerStateLogging
    SystemPowerLoggingEntry
    SetPowerSettingValue
    NotifyUserPowerSetting
    PowerInformationLevelUnused0
    PowerInformationLevelUnused1
    SystemVideoState
    TraceApplicationPowerMessage
    TraceApplicationPowerMessageEnd
    ProcessorPerfStates
    ProcessorIdleStates
    ProcessorCap
    SystemWakeSource
    SystemHiberFileInformation
    TraceServicePowerMessage
    ProcessorLoad
    PowerShutdownNotification
    MonitorCapabilities
    SessionPowerInit
    SessionDisplayState
    PowerRequestCreate
    PowerRequestAction
    GetPowerRequestList
    ProcessorInformationEx
    NotifyUserModeLegacyPowerEvent
    GroupPark
    ProcessorIdleDomains
    WakeTimerList
    SystemHiberFileSize
    PowerInformationLevelMaximum
End Enum

Public Type BatteryReportingScale
    Granularity As Long
    Capacity    As Long
End Type

Public Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Type PowerBroadcastSetting
    powerSetting As Guid
    DataLength   As Long
End Type

Public Type SystemBatteryState
    boolAcOnLine       As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolBatteryPresent As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolCharging       As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolDischarging    As Byte ' [MarshalAs(UnmanagedType.I1)]
    Spare1 As Byte
    spare2 As Byte
    spare3 As Byte
    Spare4 As Byte
    MaxCapacity        As Long ' public uint
    RemainingCapacity  As Long ' public uint
    Rate               As Long ' public uint
    EstimatedTime      As Long ' public uint
    DefaultAlert1      As Long ' public uint
    DefaultAlert2      As Long ' public uint
End Type

Public Enum SystemPowerState
    Unspecified
    Working
    Sleeping1
    Sleeping2
    Sleeping3
    Hibernate
    Shutdown
    Maximum
End Enum

Public Type SystemPowerCapabilities
    boolPowerButtonPresent   As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolSleepButtonPresent   As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolLidPresent           As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolSystemS1             As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolSystemS2             As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolSystemS3             As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolSystemS4             As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolSystemS5             As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolHiberFilePresent     As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolFullWake             As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolVideoDimPresent      As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolApmPresent           As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolUpsPresent           As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolThermalControl       As Byte ' [MarshalAs(UnmanagedType.I1)]
    boolProcessorThrottle    As Byte ' [MarshalAs(UnmanagedType.I1)]

    ProcessorMinimumThrottle As Byte
    ProcessorMaximumThrottle As Byte
    boolFastSystemS4         As Byte ' [MarshalAs(UnmanagedType.I1)]
    spare2(1 To 3)           As Byte ' [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
    boolDiskSpinDown         As Byte ' [MarshalAs(UnmanagedType.I1)]
    spare3(1 To 8)           As Byte ' public byte[] [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]

    boolSystemBatteriesPresent As Byte '[MarshalAs(UnmanagedType.I1)]
    boolBatteriesAreShortTerm  As Byte '[MarshalAs(UnmanagedType.I1)]
    BatteryScale(1 To 3)       As BatteryReportingScale '[MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]

    AcOnlineWake               As SystemPowerState
    SoftLidWake                As SystemPowerState
    RtcWake                    As SystemPowerState
    MinimumDeviceWakeState     As SystemPowerState
    DefaultLowLatencyWake      As SystemPowerState
End Type

Private Const PowerBroadcastMessage     As Long = 536& ' u
Private Const PowerSettingChangeMessage As Long = 32787  'u
Private Const ScreenSaverSetActive      As Long = 17& 'u
Private Const SendChange                As Long = 2& 'u
Private Const UpdateInFile              As Long = 1& 'u

Public Enum PowerPersonality
    PPUnknown       '/// <summary>The power personality Guid does not match a known value.</summary>
    HighPerformance '/// <summary>Power settings designed to deliver maximum performance at the expense of power consumption savings.</summary>
    PowerSaver      '/// <summary>Power settings designed consume minimum power at the expense of system performance and responsiveness.</summary>
    Automatic       '/// <summary>Power settings designed to balance performance and power consumption.</summary>
End Enum

'[DllImport("powrprof.dll")]
Public Declare Function CallNtPowerInformationSPC Lib "powrprof" Alias "CallNtPowerInformation" ( _
    ByVal informationLevel As PowerInformationLevel, _
    ByVal pinputBuffer As Long, _
    ByVal inputBufferSize As Long, _
    ByRef outputBuffer As SystemPowerCapabilities, _
    ByVal outputBufferSize As Long) As Long 'uint

'[DllImport("powrprof.dll")]
Public Declare Function CallNtPowerInformationSBS Lib "powrprof" Alias "CallNtPowerInformation" ( _
    ByVal informationLevel As PowerInformationLevel, _
    ByVal pinputBuffer As Long, _
    ByVal inputBufferSize As Long, _
    ByRef outputBuffer As SystemBatteryState, _
    ByVal outputBufferSize As Long) As Long

'[DllImport("powrprof.dll")]
Public Declare Sub PowerGetActiveScheme Lib "powrprof" ( _
    ByVal prootPowerKey As Long, _
    ByRef out_activePolicy As Guid) '[MarshalAs(UnmanagedType.LPStruct)]

'[DllImport("User32", CallingConvention = CallingConvention.StdCall, SetLastError = true)]
Private Declare Function privRegisterPowerSettingNotification Lib "user32" Alias "RegisterPowerSettingNotification" ( _
    ByVal hRecipient As Long, _
    ByRef GuidPowerSettingGuid As Guid, _
    ByVal Flags As Long) As Long


Public Enum ExecutionStates
    None = &H0&            '/// <summary>No state configured.</summary>
    SystemRequired = &H1&  '/// <summary>Forces the system to be in the working state by resetting the system idle timer.</summary>
    DisplayRequired = &H2& '/// <summary>Forces the display to be on by resetting the display idle timer.</summary>
    
    '/// <summary>
    '/// Enables away mode. This value must be specified with ES_CONTINUOUS. Away mode should be used only by media-recording and
    '/// media-distribution applications that must perform critical background processing on desktop computers while the computer appears
    '/// to be sleeping. See Remarks.
    '///
    '/// Windows Server 2003 and Windows XP/2000: ES_AWAYMODE_REQUIRED is not supported.
    '/// </summary>
    AwayModeRequired = &H40&
    
    '/// <summary>
    '/// Informs the system that the state being set should remain in effect until the next call that uses ES_CONTINUOUS and one of the
    '/// other state flags is cleared.
    '/// </summary>
    Continuous = -2147483647 'int.MinValue
End Enum

'[DllImport("kernel32.dll", SetLastError = true)]
Public Declare Function SetThreadExecutionState Lib "kernel32" ( _
    ByVal esFlags As ExecutionStates) As ExecutionStates

Public Guid_All             As Guid ' = new Guid(1755441502, 5098, 16865, 128, 17, 12, 73, 108, 164, 144, 176);
Public Guid_Automatic       As Guid ' = new Guid(941310498u, 63124, 16880, 150, 133, byte.MaxValue, 91, 178, 96, 223, 46);
Public Guid_HighPerformance As Guid ' = new Guid(2355003354u, 59583, 19094, 154, 133, 166, 226, 58, 140, 99, 92);
Public Guid_PowerSaver      As Guid ' = new Guid(2709787400u, 13633, 20395, 188, 129, 247, 21, 86, 242, 11, 74);

Public Enum PowerSource
    AC      '/// <summary> The computer is powered by an AC power source or a similar device, such as a laptop powered by a 12V automotive adapter.</summary>
    Battery '/// <summary> The computer is powered by a built-in battery. A battery has a limited amount of power; applications should conserve resources where possible. </summary>
    Ups     '/// <summary>The computer is powered by a short-term power source such as a UPS device.</summary>
End Enum
Public Enum RestartRestrictions
    None = &H0&        '/// <summary>Always restart the application.</summary>
    NotOnCrash = &H1&  '/// <summary>Do not restart when the application has crashed.</summary>
    NotOnHang = &H2&   '/// <summary>Do not restart when the application is hung.</summary>
    NotOnPatch = &H4&  '/// <summary>Do not restart when the application is terminated due to a system update.</summary>
    NotOnReboot = &H8& '/// <summary>Do not restart when the application is terminated because of a system reboot.</summary>
End Enum


'internal static class Power
'{
'internal static PowerManagementNativeMethods.SystemBatteryState GetSystemBatteryState()
'{
'    if (PowerManagementNativeMethods.CallNtPowerInformation(PowerManagementNativeMethods.PowerInformationLevel.SystemBatteryState, IntPtr.Zero, 0u, out PowerManagementNativeMethods.SystemBatteryState outputBuffer, (uint)Marshal.SizeOf(typeof(PowerManagementNativeMethods.SystemBatteryState))) == 3221225506u)
'    {
'        throw new UnauthorizedAccessException(LocalizedMessages.PowerInsufficientAccessBatteryState);
'    }
'    return outputBuffer;
'}
Public Function GetSystemBatteryState() As SystemBatteryState
    If CallNtPowerInformationSBS(PowerInformationLevel.SystemBatteryState, 0, 0, GetSystemBatteryState, LenB(GetSystemBatteryState)) = &HC0000022 Then
        Err.Raise &HC0000022, , "UnauthorizedAccessException(LocalizedMessages.PowerInsufficientAccessBatteryState)"
    End If
End Function

'internal static PowerManagementNativeMethods.SystemPowerCapabilities GetSystemPowerCapabilities()
'{
'    if (PowerManagementNativeMethods.CallNtPowerInformation(PowerManagementNativeMethods.PowerInformationLevel.SystemPowerCapabilities, IntPtr.Zero, 0u, out PowerManagementNativeMethods.SystemPowerCapabilities outputBuffer, (uint)Marshal.SizeOf(typeof(PowerManagementNativeMethods.SystemPowerCapabilities))) == 3221225506u)
'    {
'        throw new UnauthorizedAccessException(LocalizedMessages.PowerInsufficientAccessCapabilities);
'    }
'    return outputBuffer;
'}
Public Function GetSystemPowerCapabilities() As SystemPowerCapabilities
    If CallNtPowerInformationSPC(PowerInformationLevel.SystemPowerCapabilities, 0, 0, GetSystemPowerCapabilities, LenB(GetSystemPowerCapabilities)) = &HC0000022 Then
        Err.Raise &HC0000022, , "UnauthorizedAccessException(LocalizedMessages.PowerInsufficientAccessCapabilities)"
    End If
End Function

'/// <summary>Registers the application to receive power setting notifications for the specific power setting event.</summary>
'/// <param name="handle">Handle indicating where the power setting notifications are to be sent.</param>
'/// <param name="powerSetting">The GUID of the power setting for which notifications are to be sent.</param>
'/// <returns>Returns a notification handle for unregistering power notifications.</returns>
'internal static int RegisterPowerSettingNotification(IntPtr handle, Guid powerSetting)
'{
'    return PowerManagementNativeMethods.RegisterPowerSettingNotification(handle, ref powerSetting, 0);
'}
'}
Public Function RegisterPowerSettingNotification(ByVal handle As Long, powerSetting As Guid) As Long
    RegisterPowerSettingNotification = privRegisterPowerSettingNotification(handle, powerSetting, 0)
End Function

