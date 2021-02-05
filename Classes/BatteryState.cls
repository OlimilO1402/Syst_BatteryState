VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatteryState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'// Microsoft.WindowsAPICodePack.ApplicationServices.BatteryState
Private m_SystemBatteryState As SystemBatteryState

Private Sub Class_Initialize()
'        PowerManagementNativeMethods.SystemBatteryState systemBatteryState = Power.GetSystemBatteryState();
'        if (!systemBatteryState.BatteryPresent)
'        {
'            throw new InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent);
'        }
    Recall False
'    m_SystemBatteryState = MPower.GetSystemBatteryState
'    If Not CBool(m_SystemBatteryState.boolBatteryPresent) Then
'        'Err.Description = "InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent)"
'        'Err.Raise 0
'        MsgBox "InvalidOperationException( LocalizedMessages.PowerManagerBatteryNotPresent )"
'    End If
'        ACOnline = systemBatteryState.AcOnLine;
'        MaxCharge = (int)systemBatteryState.MaxCapacity;
'        CurrentCharge = (int)systemBatteryState.RemainingCapacity;
'        ChargeRate = (int)systemBatteryState.Rate;
'        uint estimatedTime = systemBatteryState.EstimatedTime;
'        if (estimatedTime != uint.MaxValue)
'        {
'            EstimatedTimeRemaining = new TimeSpan(0, 0, (int)estimatedTime);
'        }
'        Else
'        {
'            EstimatedTimeRemaining = TimeSpan.MaxValue;
'        }
'        SuggestedCriticalBatteryCharge = (int)systemBatteryState.DefaultAlert1;
'        SuggestedBatteryWarningCharge = (int)systemBatteryState.DefaultAlert2;

End Sub
Public Sub Recall(Optional ByVal loud As Boolean = False)
    m_SystemBatteryState = MPower.GetSystemBatteryState
    If Not CBool(m_SystemBatteryState.boolBatteryPresent) Then
        'Err.Description = "InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent)"
        'Err.Raise 0
        If loud Then MsgBox "InvalidOperationException( LocalizedMessages.PowerManagerBatteryNotPresent )"
    End If
End Sub
Public Property Get ACOnline() As Boolean
    ACOnline = m_SystemBatteryState.boolAcOnLine
End Property
Public Property Get ChargeRate() As Long
    ChargeRate = m_SystemBatteryState.Rate
End Property
Public Property Get CurrentCharge() As Long
    CurrentCharge = m_SystemBatteryState.RemainingCapacity
End Property
Public Property Get EstimatedTimeRemaining() As Long 'Date
    EstimatedTimeRemaining = m_SystemBatteryState.EstimatedTime
End Property
Public Property Get MaxCharge() As Long
    MaxCharge = m_SystemBatteryState.MaxCapacity
End Property
Public Property Get SuggestedBatteryWarningCharge() As Long
    SuggestedBatteryWarningCharge = m_SystemBatteryState.DefaultAlert2
End Property
Public Property Get SuggestedCriticalBatteryCharge() As Long
    SuggestedCriticalBatteryCharge = m_SystemBatteryState.DefaultAlert1
End Property
Private Function Secs_ToStr(ByVal secs As Long) As String
    Dim est As Long: est = secs
    Dim sec As Long: sec = est Mod 60
    Dim min As Long: min = (est \ 60) Mod 60
    Dim hor As Long: hor = Int(est \ 3600)
    Secs_ToStr = CStr(hor) & ":" & Format(min, "00") & ":" & Format(sec, "00")
End Function
'[BatteryStateStringRepresentation,
'ACOnline: {1}
'{0}Max Charge: {2} mWh
'{0}Current Charge: {3} mWh
'{0}Discharge Rate: {4} mWh
'{0}Estimated Time Remaining: {5}
'{0}Suggested Critical Battery Charge: {6} mWh
'{0}Suggested Battery Warning Charge: {7} mWh
'{0}]
Friend Function ToStr() As String
'return string.Format( CultureInfo.InvariantCulture, _
'                      LocalizedMessages.BatteryStateStringRepresentation, _
'                      Environment.NewLine, _
'                      ACOnline, _
'                      MaxCharge, _
'                      CurrentCharge, _
'                      ChargeRate, _
'                      EstimatedTimeRemaining, _
'                      SuggestedCriticalBatteryCharge, _
'                      SuggestedBatteryWarningCharge)
    Dim s As String
        s = "BatteryState" & vbCrLf & _
            "{" & vbCrLf & _
            "    ACOnline                       = " & Me.ACOnline & ";" & vbCrLf & _
            "    Max Charge                     = " & Me.MaxCharge & "mWh;" & vbCrLf
If Me.MaxCharge <> 0 Then
    s = s & "    CurrentCharge                  = " & Me.CurrentCharge & " (" & CLng(Me.CurrentCharge / Me.MaxCharge * 100) & "%)" & "mWh;" & vbCrLf
End If
    s = s & "    ChargeRate                     = " & Me.ChargeRate & "mWh;" & vbCrLf & _
            "    EstimatedTimeRemaining         = " & Secs_ToStr(Me.EstimatedTimeRemaining) & ";" & vbCrLf & _
            "    SuggestedCriticalBatteryCharge = " & Me.SuggestedCriticalBatteryCharge & "mWh;" & vbCrLf & _
            "    SuggestedBatteryWarningCharge  = " & Me.SuggestedBatteryWarningCharge & "mWh;" & vbCrLf & _
            "}"
    ToStr = s
End Function
'/// <summary>A snapshot of the state of the battery.</summary>
'public class BatteryState
'{
'    /// <summary>Gets a value that indicates whether the battery charger is operating on external power.</summary>
'    /// <value>A <see cref="T:System.Boolean" /> value. <b>True</b> indicates the battery charger is operating on AC power.</value>
'    public bool ACOnline
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the rate of discharge for the battery (in mW).</summary>
'    /// <remarks>
'    /// If plugged in, fully charged: DischargeRate = 0. If plugged in, charging: DischargeRate = positive mW per hour. If unplugged:
'    /// DischargeRate = negative mW per hour.
'    /// </remarks>
'    /// <value>An <see cref="T:System.Int32" /> value.</value>
'    public int ChargeRate
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the current charge of the battery (in mW).</summary>
'    /// <value>An <see cref="T:System.Int32" /> value.</value>
'    public int CurrentCharge
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the estimated time remaining until the battery is empty.</summary>
'    /// <value>A <see cref="T:System.TimeSpan" /> object.</value>
'    public TimeSpan EstimatedTimeRemaining
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the maximum charge of the battery (in mW).</summary>
'    /// <value>An <see cref="T:System.Int32" /> value.</value>
'    public int MaxCharge
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the manufacturer's suggested battery charge level that should cause a warning to be sent to the user.</summary>
'    /// <value>An <see cref="T:System.Int32" /> value.</value>
'    public int SuggestedBatteryWarningCharge
'    {
'        get;
'        private set;
'    }
'
'    /// <summary>Gets the manufacturer's suggested battery charge level that should cause a critical alert to be sent to the user.</summary>
'    /// <value>An <see cref="T:System.Int32" /> value.</value>
'    public int SuggestedCriticalBatteryCharge
'    {
'        get;
'        private set;
'    }
'
'    internal BatteryState()
'    {
'        PowerManagementNativeMethods.SystemBatteryState systemBatteryState = Power.GetSystemBatteryState();
'        if (!systemBatteryState.BatteryPresent)
'        {
'            throw new InvalidOperationException(LocalizedMessages.PowerManagerBatteryNotPresent);
'        }
'        ACOnline = systemBatteryState.AcOnLine;
'        MaxCharge = (int)systemBatteryState.MaxCapacity;
'        CurrentCharge = (int)systemBatteryState.RemainingCapacity;
'        ChargeRate = (int)systemBatteryState.Rate;
'        uint estimatedTime = systemBatteryState.EstimatedTime;
'        if (estimatedTime != uint.MaxValue)
'        {
'            EstimatedTimeRemaining = new TimeSpan(0, 0, (int)estimatedTime);
'        }
'        Else
'        {
'            EstimatedTimeRemaining = TimeSpan.MaxValue;
'        }
'        SuggestedCriticalBatteryCharge = (int)systemBatteryState.DefaultAlert1;
'        SuggestedBatteryWarningCharge = (int)systemBatteryState.DefaultAlert2;
'    }
'
'    /// <summary>Generates a string that represents this <b>BatteryState</b> object.</summary>
'    /// <returns>A <see cref="T:System.String" /> representation of this object's current state.</returns>
'    public override string ToString()
'    {
'        return string.Format(CultureInfo.InvariantCulture, LocalizedMessages.BatteryStateStringRepresentation, Environment.NewLine, ACOnline, MaxCharge, CurrentCharge, ChargeRate, EstimatedTimeRemaining, SuggestedCriticalBatteryCharge, SuggestedBatteryWarningCharge);
'    }
'}
