Attribute VB_Name = "MPowerManager"
Option Explicit
'internal static class PowerPersonalityGuids
'{
'    internal static readonly Guid All             = new Guid(1755441502, 5098, 16865, 128, 17, 12, 73, 108, 164, 144, 176);
'    internal static readonly Guid Automatic       = new Guid(941310498u, 63124, 16880, 150, 133, byte.MaxValue, 91, 178, 96, 223, 46);
'    internal static readonly Guid HighPerformance = new Guid(2355003354u, 59583, 19094, 154, 133, 166, 226, 58, 140, 99, 92);
'    internal static readonly Guid PowerSaver      = new Guid(2709787400u, 13633, 20395, 188, 129, 247, 21, 86, 242, 11, 74);
'    internal static PowerPersonality GuidToEnum(Guid guid)
'    {
'        if (guid == HighPerformance) { return PowerPersonality.HighPerformance; }
'        if (guid == PowerSaver)      { return PowerPersonality.PowerSaver;      }
'        if (guid == Automatic)       { return PowerPersonality.Automatic;       }
'        return PowerPersonality.Unknown;
'    }
'}
Private m_All             As Guid
Private m_Automatic       As Guid
Private m_HighPerformance As Guid
Private m_PowerSaver      As Guid

Private m_monitoronlock     As Object
Private m_isMonitorOn       As Boolean '? Nullable
Private m_monitorRequired   As Boolean
Private m_requestBlockSleep As Boolean

Public Sub InitGuids()
                m_All = New_Guid(&H68A1E95E, &H13EA, &H41E1, 128, 17, 12, 73, 108, 164, 144, 176)
          m_Automatic = New_Guid(&H381B4222, &HF694, &H41F0, 150, 133, 255, 91, 178, 96, 223, 46)
    m_HighPerformance = New_Guid(&H8C5E7FDA, &HE8BF, &H4A96, 154, 133, 166, 226, 58, 140, 99, 92)
         m_PowerSaver = New_Guid(&HA1841308, &H3541, &H4FAB, 188, 129, 247, 21, 86, 242, 11, 74)
End Sub
Public Function New_Guid(ByVal Data1 As Long, ByVal Data2 As Integer, ByVal Data3 As Integer, _
                         ByVal Data40 As Byte, ByVal Data41 As Byte, ByVal Data42 As Byte, ByVal Data43 As Byte, _
                         ByVal Data44 As Byte, ByVal Data45 As Byte, ByVal Data46 As Byte, ByVal Data47 As Byte) As Guid
    With New_Guid
        .Data1 = Data1: .Data2 = Data2: .Data3 = Data3
        .Data4(0) = Data40: .Data4(1) = Data41: .Data4(2) = Data42: .Data4(3) = Data43
        .Data4(4) = Data44: .Data4(5) = Data45: .Data4(6) = Data46: .Data4(7) = Data47
    End With
End Function
Public Function Guid_IsEqual(g1 As Guid, g2 As Guid) As Boolean
    If g1.Data1 <> g2.Data1 Then Exit Function
    If g1.Data2 <> g2.Data2 Then Exit Function
    If g1.Data3 <> g2.Data3 Then Exit Function
    Dim i As Long
    For i = 0 To 7
        If g1.Data4(i) <> g2.Data4(i) Then Exit Function
    Next
End Function
Public Function GuidToEnum(aGuid As Guid) As PowerPersonality
    If Guid_IsEqual(aGuid, m_HighPerformance) Then _
        GuidToEnum = MPower.PowerPersonality.HighPerformance: Exit Function
    If Guid_IsEqual(aGuid, m_PowerSaver) Then _
        GuidToEnum = MPower.PowerPersonality.PowerSaver: Exit Function
    If Guid_IsEqual(aGuid, m_Automatic) Then _
        GuidToEnum = MPower.PowerPersonality.Automatic: Exit Function
    GuidToEnum = MPower.PowerPersonality.PPUnknown
End Function


'public static class PowerManager
'{
'    private static readonly object monitoronlock = new object();
'
'    private static bool? isMonitorOn;
'
'    private static bool monitorRequired;
'
'    private static bool requestBlockSleep;
'
'    /// <summary>
'    /// Gets a value that indicates the remaining battery life (as a percentage of the full battery charge). This value is in the range
'    /// 0-100, where 0 is not charged and 100 is fully charged.
'    /// </summary>
'    /// <exception cref="T:System.InvalidOperationException">The system does not have a battery.</exception>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires Vista/Windows Server 2008.</exception>
'    /// <value>An <see cref="T:System.Int32" /> value.</value>
'    public static int BatteryLifePercent
'    {
'        get
'        {
'            CoreHelpers.ThrowIfNotVista();
'            if (!Power.GetSystemBatteryState().BatteryPresent)
'            {
'                throw new InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent);
'            }
'            PowerManagementNativeMethods.SystemBatteryState systemBatteryState = Power.GetSystemBatteryState();
'            return (int)Math.Round((double)systemBatteryState.RemainingCapacity / (double)systemBatteryState.MaxCapacity * 100.0, 0);
'        }
'    }
Public Property Get BatteryLifePercent() As Long
    Dim aSystemBatteryState As SystemBatteryState: aSystemBatteryState = MPower.GetSystemBatteryState
    If Not aSystemBatteryState.boolBatteryPresent Then
        'Err.Raise -1, , "InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent)"
        MsgBox "InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent)"
    End If
    BatteryLifePercent = aSystemBatteryState.RemainingCapacity / aSystemBatteryState.MaxCapacity * 100
End Property
'
'    /// <summary>Gets a value that indicates whether a battery is present. The battery can be a short term battery.</summary>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires XP/Windows Server 2003 or higher.</exception>
'    /// <value>A <see cref="T:System.Boolean" /> value.</value>
'    public static bool IsBatteryPresent
'    {
'        get
'        {
'            CoreHelpers.ThrowIfNotXP();
'            return Power.GetSystemBatteryState().BatteryPresent;
'        }
'    }
Public Property Get IsBatteryPresent() As Boolean
    IsBatteryPresent = MPower.GetSystemBatteryState.boolBatteryPresent
End Property

'    /// <summary>Gets a value that indicates whether the battery is a short term battery.</summary>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires XP/Windows Server 2003 or higher.</exception>
'    /// <value>A <see cref="T:System.Boolean" /> value.</value>
'    public static bool IsBatteryShortTerm
'    {
'        get
'        {
'            CoreHelpers.ThrowIfNotXP();
'            return Power.GetSystemPowerCapabilities().BatteriesAreShortTerm;
'        }
'    }
Public Property Get IsBatteryShortTerm() As Boolean
    IsBatteryShortTerm = MPower.GetSystemPowerCapabilities.boolBatteriesAreShortTerm
End Property

'    /// <summary>Gets a value that indictates whether the monitor is on.</summary>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires Vista/Windows Server 2008.</exception>
'    /// <value>A <see cref="T:System.Boolean" /> value.</value>
'    public static bool IsMonitorOn
'    {
'        get
'        {
'            CoreHelpers.ThrowIfNotVista();
'            Lock (monitoronlock)
'            {
'                if (!isMonitorOn.HasValue)
'                {
'                    IsMonitorOnChanged += delegate
'                    {
'                    };
'                    EventManager.monitorOnReset.WaitOne();
'                }
'            }
'            return isMonitorOn.Value;
'        }
'        internal set
'        {
'            isMonitorOn = value;
'        }
'    }
Public Property Get IsMonitorOn() As Boolean
    IsMonitorOn = m_isMonitorOn
End Property
Private Property Let IsMonitorOn(ByVal Value As Boolean)
    m_isMonitorOn = Value
End Property

'
'    /// <summary>Gets a value that indicates a UPS is present to prevent sudden loss of power.</summary>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires XP/Windows Server 2003 or higher.</exception>
'    /// <value>A <see cref="T:System.Boolean" /> value.</value>
'    public static bool IsUpsPresent
'    {
'        get
'        {
'            CoreHelpers.ThrowIfNotXP();
'            PowerManagementNativeMethods.SystemPowerCapabilities systemPowerCapabilities = Power.GetSystemPowerCapabilities();
'            if (systemPowerCapabilities.BatteriesAreShortTerm)
'            {
'                return systemPowerCapabilities.SystemBatteriesPresent;
'            }
'            return false;
'        }
'    }
Public Property Get IsUpsPresent() As Boolean
    Dim aSystemPowerCapabilities As SystemPowerCapabilities: aSystemPowerCapabilities = MPower.GetSystemPowerCapabilities
    If aSystemPowerCapabilities.boolBatteriesAreShortTerm Then
        IsUpsPresent = CBool(aSystemPowerCapabilities.boolSleepButtonPresent)
        'Exit Property
    End If
End Property

'    /// <summary>Gets or sets a value that indicates whether the monitor is set to remain active.</summary>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires XP/Windows Server 2003 or higher.</exception>
'    /// <exception cref="T:System.Security.SecurityException">The caller does not have sufficient privileges to set this property.</exception>
'    /// <remarks>
'    /// This information is typically used by applications that display information but do not require user interaction. For example,
'    /// video playback applications.
'    /// </remarks>
'    /// <permission cref="T:System.Security.Permissions.SecurityPermission">
'    /// to set this property. Demand value: <see cref="F:System.Security.Permissions.SecurityAction.Demand" />; Named Permission Sets: <b>FullTrust</b>.
'    /// </permission>
'    /// <value>A <see cref="T:System.Boolean" /> value. <b>True</b> if the monitor is required to remain on.</value>
'    public static bool MonitorRequired
'    {
'        get
'        {
'            CoreHelpers.ThrowIfNotXP();
'            return monitorRequired;
'        }
'        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
'        set
'        {
'            CoreHelpers.ThrowIfNotXP();
'            if (value)
'            {
'                SetThreadExecutionState(ExecutionStates.DisplayRequired | ExecutionStates.Continuous);
'            }
'            Else
'            {
'                SetThreadExecutionState(ExecutionStates.Continuous);
'            }
'            monitorRequired = value;
'        }
'    }
Public Property Get MonitorRequired() As Boolean
    MonitorRequired = m_monitorRequired
End Property
Public Property Let MonitorRequired(ByVal Value As Boolean)
    m_monitorRequired = Value
End Property

'    /// <summary>Gets a value that indicates the current power scheme.</summary>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires Vista/Windows Server 2008.</exception>
'    /// <value>A <see cref="P:Microsoft.WindowsAPICodePack.ApplicationServices.PowerManager.PowerPersonality" /> value.</value>
'    public static PowerPersonality PowerPersonality
'    {
'        get
'        {
'            PowerManagementNativeMethods.PowerGetActiveScheme(IntPtr.Zero, out Guid activePolicy);
'            try
'            {
'                return PowerPersonalityGuids.GuidToEnum(activePolicy);
'            }
'            finally
'            {
'                CoreNativeMethods.LocalFree(ref activePolicy);
'            }
'        }
'    }
Public Property Get PowerPersonality() As PowerPersonality
    Dim aactivePolicy As Guid
    MPower.PowerGetActiveScheme 0, aactivePolicy
Try: On Error GoTo Catch
    
Catch:
End Property

'    /// <summary>Gets the current power source.</summary>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires Vista/Windows Server 2008.</exception>
'    /// <value>A <see cref="P:Microsoft.WindowsAPICodePack.ApplicationServices.PowerManager.PowerSource" /> value.</value>
'    public static PowerSource PowerSource
'    {
'        get
'        {
'            CoreHelpers.ThrowIfNotVista();
'            if (IsUpsPresent)
'            {
'                return PowerSource.Ups;
'            }
'            if (!IsBatteryPresent || GetCurrentBatteryState().ACOnline)
'            {
'                return PowerSource.AC;
'            }
'            return PowerSource.Battery;
'        }
'    }
Public Property Get PowerSource() As PowerSource
    If IsUpsPresent Then
        PowerSource = MPower.PowerSource.Ups
    ElseIf Not IsBatteryPresent And GetCurrentBatteryState.ACOnline Then
        PowerSource = MPower.PowerSource.AC
    Else
        PowerSource = MPower.PowerSource.Battery
    End If
End Property
Public Function PowerSource_ToStr(aps As PowerSource) As String
    Dim s As String
    Select Case aps
    Case MPower.PowerSource.Ups:      s = "PowerSource.Ups"
    Case MPower.PowerSource.AC:       s = "PowerSource.AC"
    Case MPower.PowerSource.Battery:  s = "PowerSource.Battery"
    End Select
    PowerSource_ToStr = s
End Function
'
'    /// <summary>Gets or sets a value that indicates whether the system is required to be in the working state.</summary>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires XP/Windows Server 2003 or higher.</exception>
'    /// <exception cref="T:System.Security.SecurityException">The caller does not have sufficient privileges to set this property.</exception>
'    /// <permission cref="T:System.Security.Permissions.SecurityPermission">
'    /// to set this property. Demand value: <see cref="F:System.Security.Permissions.SecurityAction.Demand" />; Named Permission Sets: <b>FullTrust</b>.
'    /// </permission>
'    /// <value>A <see cref="T:System.Boolean" /> value.</value>
'    public static bool RequestBlockSleep
'    {
'        get
'        {
'            CoreHelpers.ThrowIfNotXP();
'            return requestBlockSleep;
'        }
'        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
'        set
'        {
'            CoreHelpers.ThrowIfNotXP();
'            if (value)
'            {
'                SetThreadExecutionState(ExecutionStates.SystemRequired | ExecutionStates.Continuous);
'            }
'            Else
'            {
'                SetThreadExecutionState(ExecutionStates.Continuous);
'            }
'            requestBlockSleep = value;
'        }
'    }
'
'    /// <summary>Raised when the remaining battery life changes.</summary>
'    /// <exception cref="T:System.InvalidOperationException">The event handler specified for removal was not registered.</exception>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires Vista/Windows Server 2008.</exception>
'    public static event EventHandler BatteryLifePercentChanged
'    {
'        Add
'        {
'            CoreHelpers.ThrowIfNotVista();
'            MessageManager.RegisterPowerEvent(EventManager.BatteryCapacityChange, value);
'        }
'        Remove
'        {
'            CoreHelpers.ThrowIfNotVista();
'            MessageManager.UnregisterPowerEvent(EventManager.BatteryCapacityChange, value);
'        }
'    }
'
'    /// <summary>Raised when the monitor status changes.</summary>
'    /// <exception cref="T:System.InvalidOperationException">The event handler specified for removal was not registered.</exception>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires Vista/Windows Server 2008.</exception>
'    public static event EventHandler IsMonitorOnChanged
'    {
'        Add
'        {
'            CoreHelpers.ThrowIfNotVista();
'            MessageManager.RegisterPowerEvent(EventManager.MonitorPowerStatus, value);
'        }
'        Remove
'        {
'            CoreHelpers.ThrowIfNotVista();
'            MessageManager.UnregisterPowerEvent(EventManager.MonitorPowerStatus, value);
'        }
'    }
'
'    /// <summary>Raised each time the active power scheme changes.</summary>
'    /// <exception cref="T:System.InvalidOperationException">The event handler specified for removal was not registered.</exception>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires Vista/Windows Server 2008.</exception>
'    public static event EventHandler PowerPersonalityChanged
'    {
'        Add
'        {
'            MessageManager.RegisterPowerEvent(EventManager.PowerPersonalityChange, value);
'        }
'        Remove
'        {
'            CoreHelpers.ThrowIfNotVista();
'            MessageManager.UnregisterPowerEvent(EventManager.PowerPersonalityChange, value);
'        }
'    }
'
'    /// <summary>Raised when the power source changes.</summary>
'    /// <exception cref="T:System.InvalidOperationException">The event handler specified for removal was not registered.</exception>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires Vista/Windows Server 2008.</exception>
'    public static event EventHandler PowerSourceChanged
'    {
'        Add
'        {
'            CoreHelpers.ThrowIfNotVista();
'            MessageManager.RegisterPowerEvent(EventManager.PowerSourceChange, value);
'        }
'        Remove
'        {
'            CoreHelpers.ThrowIfNotVista();
'            MessageManager.UnregisterPowerEvent(EventManager.PowerSourceChange, value);
'        }
'    }
'
'    /// <summary>
'    /// Raised when the system will not be moving into an idle state in the near future so applications should perform any tasks that
'    /// would otherwise prevent the computer from entering an idle state.
'    /// </summary>
'    /// <exception cref="T:System.InvalidOperationException">The event handler specified for removal was not registered.</exception>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires Vista/Windows Server 2008.</exception>
'    public static event EventHandler SystemBusyChanged
'    {
'        Add
'        {
'            CoreHelpers.ThrowIfNotVista();
'            MessageManager.RegisterPowerEvent(EventManager.BackgroundTaskNotification, value);
'        }
'        Remove
'        {
'            CoreHelpers.ThrowIfNotVista();
'            MessageManager.UnregisterPowerEvent(EventManager.BackgroundTaskNotification, value);
'        }
'    }
Public Sub SystemBusyChanged()
    'Add
    'MessageManager.RegisterPowerEvent EventManager.BackgroundTaskNotification, Value
    'Remove
    'MessageManager.UnregisterPowerEvent EventManager.BackgroundTaskNotification, Value
End Sub

'    /// <summary>Gets a snapshot of the current battery state.</summary>
'    /// <returns>A <see cref="T:Microsoft.WindowsAPICodePack.ApplicationServices.BatteryState" /> instance that represents the state of the battery at the time this method was called.</returns>
'    /// <exception cref="T:System.InvalidOperationException">The system does not have a battery.</exception>
'    /// <exception cref="T:System.PlatformNotSupportedException">Requires XP/Windows Server 2003 or higher.</exception>
'    public static BatteryState GetCurrentBatteryState()
'    {
'        CoreHelpers.ThrowIfNotXP();
'        return new BatteryState();
'    }
Public Property Get GetCurrentBatteryState() As BatteryState
    Set GetCurrentBatteryState = New BatteryState
End Property
'    /// <summary>
'    /// Allows an application to inform the system that it is in use, thereby preventing the system from entering the sleeping power
'    /// state or turning off the display while the application is running.
'    /// </summary>
'    /// <param name="executionStateOptions">The thread's execution requirements.</param>
'    /// <exception cref="T:System.ComponentModel.Win32Exception">Thrown if the SetThreadExecutionState call fails.</exception>
'    public static void SetThreadExecutionState(ExecutionStates executionStateOptions)
'    {
'        if (PowerManagementNativeMethods.SetThreadExecutionState(executionStateOptions) == ExecutionStates.None)
'        {
'            throw new Win32Exception(LocalizedMessages.PowerExecutionStateFailed);
'        }
'    }
Public Sub SetThreadExecutionState(executionStateOptions As ExecutionStates)
    If MPower.SetThreadExecutionState(executionStateOptions) = ExecutionStates.None Then
        'Err.Raise
        MsgBox "Win32Exception(LocalizedMessages.PowerExecutionStateFailed)"
    End If
End Sub

'}
'
