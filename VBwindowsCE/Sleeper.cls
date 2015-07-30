VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sleeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_evtMgr As IEventManager2
Private WithEvents m_timer As CTimer
Attribute m_timer.VB_VarHelpID = -1
Private m_sleeping As Boolean

Private Sub Class_Initialize()
    ' create a timer
    '
    Dim shell As New CShell
    Set m_timer = shell.CreateTimer()

    ' get the event manager interface from cshell
    '
    Set m_evtMgr = shell
End Sub

' -- Sleep will give up the CPU for ms milliseconds
Public Sub Sleep(ByVal ms As Long)
    ' this class can't really handle re-entrant sleeping
    '
    If (m_sleeping) Then Exit Sub
    m_sleeping = True

    ' start the timer that will tell us when to stop sleeping
    '
    m_timer.Period = ms
    m_timer.Start

    ' run the event loop (this will exit when our timer event
    '   handler calls Terminate())
    '
    m_evtMgr.Loop

    ' calling DoEvents() once will reset the "terminated" state
    '   of the event loop, so that the outer one won't quit as well
    '
    DoEvents
    
    m_sleeping = False
End Sub

Private Sub m_timer_Trigger(ByVal timer As AFCoreLib.ITimer)
    ' stop our instance of the event loop, and our timer
    '
    m_evtMgr.Terminate
    m_timer.Stop
End Sub

