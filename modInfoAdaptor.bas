Attribute VB_Name = "modInfoAdaptor"
Public m_AdaptorCnt As Long    'count the adaptors.
Public m_MaxNewForm As Byte   'Count the max frmMains
Public cntMainForm(1 To 100) As Byte  'index off frmMains
Public Y As Integer    ' uses to place new windows..

Public Xmenu() As frmMain
Public Max_Adapters As Long
Public AdaptorPopup As Byte

