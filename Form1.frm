VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fmMain 
   Caption         =   "OPCSAVEDATA-SubVersion Application By Mister.T"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   14655
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "Read"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7695
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20188
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "7/22/2011"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LvListView 
      Height          =   5895
      Left            =   960
      TabIndex        =   7
      Top             =   840
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   10398
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�������"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CheckBox DataChgChk 
      Caption         =   "ʹ�ö������ݲ�ȡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12360
      TabIndex        =   3
      Top             =   840
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Timer tmUpdate 
      Left            =   13440
      Top             =   4680
   End
   Begin VB.CommandButton btnAddItem 
      Caption         =   "��ȡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton btnQuit 
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame OPC 
      Caption         =   "OPC���ݲɼ�"
      Height          =   7335
      Left            =   360
      TabIndex        =   8
      Top             =   360
      Width           =   14055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   6600
      Width           =   2055
   End
End
Attribute VB_Name = "fmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit

' OPC���������
Dim WithEvents objserver As OPCServer
Attribute objserver.VB_VarHelpID = -1
Dim objGroups As OPCGroups
Dim WithEvents objtestgrp As OPCGroup '�¼��Ķ�Ӧ
Attribute objtestgrp.VB_VarHelpID = -1
Dim objItems As OPCItems
Dim LServerHandles() As Long
    Dim x As Integer
    Dim y As Integer
    Dim mon As Integer
    Dim d As Integer
    Dim h As Integer
    Dim m As Integer
    Dim s As Integer
    Dim ss  As Boolean
Dim lTransID_Rd As Long
Dim lCancelID_Rd As Long
Dim lTransID_Wt As Long
Dim lCancelID_Wt As Long

Sub Connect(strProgID As String, Optional strNode As String)
    
    If objserver Is Nothing Then
        ' ����һ��OPC����������
        Set objserver = New OPCServer
    End If
    
    If objserver.ServerState = OPCDisconnected Then
        ' ����OPC������
        objserver.Connect strProgID, strNode
    End If
    
    If objGroups Is Nothing Then
        ' ����һ��OPC�鼯��
        Set objGroups = objserver.OPCGroups
    End If
    
    If objtestgrp Is Nothing Then
        ' ���һ��OPC��
        Set objtestgrp = objGroups.Add("TestGrp")
    End If
    
End Sub

Sub Disconnect()
    Dim lErrors() As Long

    If Not objItems Is Nothing Then
        If objItems.Count > 0 Then
            ' ���OPC��
            objItems.Remove 114, LServerHandles, lErrors
        End If
        Set objItems = Nothing
    End If
    
    If Not objtestgrp Is Nothing Then
        ' ���OPC��
        objGroups.Remove "TestGrp"
        Set objtestgrp = Nothing
    End If
    
    If Not objGroups Is Nothing Then
        Set objGroups = Nothing
    End If
    
    If Not objserver Is Nothing Then
        If objserver.ServerState <> OPCDisconnected Then
            ' �Ͽ�OPC������.
            objserver.Disconnect
        End If
        
        Set objserver = Nothing
    End If
        
End Sub

Sub AddItem()
    Dim strItemIDs(114) As String
    Dim lClientHandles(114) As Long
    Dim lErrors() As Long
    Dim i As Integer
    Dim ExcelApp As Excel.Application
    Dim ExcelBook As Excel.Workbook
    Dim ExcelSheet As Excel.Worksheet
    Dim Strfilename As String
    
    If objtestgrp Is Nothing Then
        Exit Sub
    End If
    
    If Not objItems Is Nothing Then
        If objItems.Count > 0 Then
            Exit Sub
        End If
    End If
    
    ' ������״̬
    If DataChgChk.Value = vbChecked Then
    
        objtestgrp.IsActive = True
    Else
        objtestgrp.IsActive = False
    End If
    ' �������ͬ��֪ͨ
    objtestgrp.IsSubscribed = True
    
    ' ����OPC���
    Set objItems = objtestgrp.OPCItems
    '��ѯ���ݿ⣬�õ�TAG��
     Dim TagName As String, MyTable As String
     TagName = "TAGNAME"
     MyTable = "TAG"
     Dim testg As Variant
   
    ' testg = ReadData(MyTable, Tagname)
   For i = 1 To LvListView.ListItems.Count

   lClientHandles(i) = i
   strItemIDs(i) = LvListView.ListItems(i).SubItems(1)
   Next i

'     I = 1
' Do While Not rs.EOF
 ' strItemIDs(I) = rs("TAGNAME")
  '              lClientHandles(I) = I
   '             I = I + 1
    '            rs.MoveNext
'Loop


    
'������ʾ��Ϣ
   ' MsgBox "XXXXXXXXXX", vbInformation + vbOKOnly
    '�ر����ݼ��������ݿ�����ӣ����ͷű���


    StatusBar1.Panels(1).Text = " ���ݼ��سɹ�"
    
    ' ���ɴ�TAG1��TAG8�����ʶ��
   ' For I = 1 To 5
  '      strItemIDs(I) = "Simulation Items.Integer.Int_0" & I
 '       lClientHandles(I) = I
 '   Next
 '   For I = 6 To 8
 '            strItemIDs(I) = "Simulation Items.Real.Real_0" & I - 5
'        lClientHandles(I) = I
'    Next
    ' ���OPC��
    Call objItems.AddItems(114, strItemIDs, _
        lClientHandles, LServerHandles, lErrors)

'         ss = False
'
'Label1:
'Call AsyncRead
'Dim Savetime As Double
'timeBeginPeriod 1
'Savetime = timeGetTime
'While timeGetTime < Savetime + 1000
'     If ss = True Then
'    timeEndPeriod 1
'     Call Disconnect
''     Set OPCServer = Nothing
'        Exit Sub
'    End If
'DoEvents
'Wend
'GoTo Label1
    
End Sub

Sub AsyncRead()
    Dim lErrors() As Long
   StatusBar1.Panels(1).Text = "Data Reading..."
    If objtestgrp Is Nothing Then
        Exit Sub
    End If
    
    If objtestgrp.OPCItems.Count > 0 Then
        ' ��ͬ�ڶ�ȡ
        lTransID_Rd = lTransID_Rd + 1
        objtestgrp.AsyncRead 114, LServerHandles, _
            lErrors, lTransID_Rd, lCancelID_Rd
    End If
 
End Sub

Sub AsyncWrite(nIndex As Integer, ByRef vtItemValues() As Variant, _
    ByRef lErrors() As Long)
Dim lHandle(1) As Long
    
    If objtestgrp Is Nothing Then
        Exit Sub
    End If
    
    If objtestgrp.OPCItems.Count > 0 Then
        lHandle(1) = LServerHandles(nIndex)
        
        ' ��ͬ��д��
        lTransID_Wt = lTransID_Wt + 1
        objtestgrp.AsyncWrite 1, lHandle(), vtItemValues, _
                lErrors, lTransID_Wt, lCancelID_Wt
    End If

End Sub

Private Sub Command1_Click()

    Dim i As Integer
    Dim ExcelApp As Excel.Application
    Dim ExcelBook As Excel.Workbook
    Dim ExcelSheet As Excel.Worksheet
    Dim Strfilename As String
    Dim xxx(8) As Variant
    Dim yyy(8) As Variant

    Strfilename = "D:\My Documents\Desktop\ASYNC.xls"
    Set ExcelApp = New Excel.Application
    Set ExcelBook = ExcelApp.Workbooks.Open(Strfilename)
    Set ExcelSheet = ExcelBook.Sheets(1)
    ExcelApp.Visible = False
    
    With Worksheets("Sheet1")
            For i = 1 To 8
            ' �ӹ������еõ�TAG1��TAG8�����ʶ��
                xxx(i) = .Cells(i + 1, 3).Text
                yyy(i) = .Cells(i + 1, 2).Text
            Next i
  '       .Range("A2:i65").ClearContents
        End With
    ExcelApp.Quit
    Set ExcelBook = Nothing
    Set ExcelSheet = Nothing
    Set ExcelApp = Nothing
    

           
           '����ACCESS

           
    Dim MyData As String, MyTable As String
    Dim Cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim x As Integer
    Dim myArray As Variant
    myArray = Array("00013", "����ó�ײ�", "��ΰ��", "��", "��Ŀ����", _
        "����ʦ", "˶ʿ", "�ӱ�ʡ", #3/28/1982#, 24, #8/16/2006#, 0, #8/16/2006#, 0)
   MyData = "D:\My Documents\Desktop\OPC1.mdb"
MyTable = "ְ��������Ϣ"
'���������ݿ������
    Set Cnn = New ADODB.Connection
    With Cnn
        .Provider = "microsoft.jet.oledb.4.0"
        .Open MyData
End With
'��ѯ���ݱ�
    Set rs = New ADODB.Recordset
    rs.Open "TAG", Cnn, 1, 3
    '��Ӽ�¼
With rs
       
        For x = 1 To 8
         .AddNew      '��Ӹ����ֶε�����
            rs("TAGName") = xxx(x)
            rs("TAGDIS") = yyy(x)
         .Update
        Next x
             '�������ݱ�
End With
'������ʾ��Ϣ
    MsgBox "�Ѿ��ɹ�����ְ����������ӵ����ݿ��У�", vbInformation + vbOKOnly
    '�ر����ݼ��������ݿ�����ӣ����ͷű���
rs.Close
    Cnn.Close
    Set rs = Nothing
    Set Cnn = Nothing

Dim mycat As New adox.Catalog  '����ADOX��catalog�������
Dim mytbl As New Table  '����table�������



'�������ݿ����ƣ���������·����

'����Ҫ���������ݱ�����
MyTable = "TAGVALUE"
'���������ݿ������
mycat.ActiveConnection = "provider=microsoft.jet.oledb.4.0;" & "data source=" & MyData
'ɾ�����ݿ����Ѿ����ڵ����ݱ�
'mycat.Tables.Delete mytable

'�������ݱ�������ֶ�
With mytbl
    .Name = MyTable
    .Columns.Append "ѧ��", adVarWChar, 10
    For x = 1 To 10
        .Columns.Append yyy(x), adDouble
    Next x
'    .Columns.Append "����", adVarWChar, 6
'    .Columns.Append "�Ա�", adVarWChar, 1
'    .Columns.Append "�༶", adVarWChar, 10
'    .Columns.Append "��ѧ", adDouble
 '   .Columns.Append "����", adSingle
'    .Columns.Append "����", adSingle
'    .Columns.Append "��ѧ", adSingle
'    .Columns.Append "Ӣ��", adSingle
'    .Columns.Append "DATE", adDBTimeStamp
End With
'�����������ݱ���ӵ�ADOX��tables������
mycat.Tables.Append mytbl
'�ͷű���
Set mycat = Nothing
Set mytbl = Nothing
'������Ϣ
MsgBox "���ݱ�<" & MyTable & ">�����ɹ���", vbInformation, "�������ݱ�"


End Sub

Private Sub Command2_Click()
Call AsyncRead
End Sub

Private Sub DataChgChk_Click()

    If DataChgChk.Value = vbChecked Then
        tmUpdate.Enabled = False
        If Not objtestgrp Is Nothing Then
            objtestgrp.IsActive = True
        End If
    Else
         tmUpdate.Enabled = True
        If Not objtestgrp Is Nothing Then
            objtestgrp.IsActive = False
        End If
    End If
    
End Sub


Private Sub Form_Load()
    ' ��ʼ��ȫ�ֱ���
    DataChgChk.Value = vbUnchecked

    tmUpdate.Enabled = True
  '  tmUpdate.Enabled = False
    tmUpdate.Interval = 1000
    lTransID_Rd = 0
    lTransID_Wt = 0
    LvListView.ColumnHeaders.Add 1, , "���", LvListView.Width / 4
    LvListView.ColumnHeaders.Add 2, , "��ǩ����", LvListView.Width / 4
    LvListView.ColumnHeaders.Add 3, , "��ǩ����", LvListView.Width / 4
    LvListView.ColumnHeaders.Add 4, , "��ǩֵ", LvListView.Width / 4
     LvListView.ColumnHeaders.Add 5, , "TEST", LvListView.Width / 5
    
     Dim TagName As String, MyTable As String
     TagName = "TAGNAME"
     MyTable = "TAG"
     Dim testg() As Report_Data
     ReDim testg(114) As Report_Data
     
     testg = ReadData(MyTable, TagName)
     
   For i = LBound(testg) To UBound(testg)
 'For i = 1 To 8
   
Set itx = LvListView.ListItems.Add(, , i)
itx.SubItems(1) = testg(i).TagName
itx.SubItems(2) = testg(i).TagDIS
   Next i
' Set itx = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' ����Disconnect�ӳ���
    Call Disconnect

End Sub

Private Sub btnConnect_Click()
    '����Connect�ӳ���
  ' Call Connect("DSxPOpcSimulator.TSxOpcSimulator.1")
     ' Call Connect("DSxPOpcSimulator.TSxOpcSimulator.1", "127.0.0.1")
       Call Connect("Freelance2000OPCServer.49.1", "172.16.1.48")

End Sub

Private Sub btnAddItem_Click()
    ' ����AddItem�ӳ���
    Call AddItem
    
End Sub

Private Sub btnQuit_Click()
    ' ж�ش���
    ss = True
    Unload fmMain

End Sub





Private Sub tmUpdate_Timer()
    ' ��ͬ�ڶ�ȡ
    
    Call AsyncRead
    
End Sub



Private Sub objtestgrp_AsyncReadComplete( _
    ByVal TransactionID As Long, ByVal NumItems As Long, _
    ClientHandles() As Long, ItemValues() As Variant, _
    Qualities() As Long, TimeStamps() As Date, Errors() As Long)
'    Dim strBuf As String
'    Dim nWidth As Integer
'    Dim nHeight As Integer
'    Dim nDrawHeight As Integer
'    Dim sglScale As Single
 '   Dim i As Integer
'    Dim Index As Integer


'Dim x1 As Integer
'    For x1 = 1 To UBound(ClientHandles)
'    LvListView.ListItems(x1).SubItems(4) = ClientHandles(x1)
'
'    Next x1
        For i = 1 To UBound(ItemValues)

'          Set itx = LvListView.ListItems(i)
'         itx.SubItems(3) = ItemValues(i)
'         Next i

          LvListView.ListItems(i).SubItems(3) = ItemValues(i)
          Next i
          
 StatusBar1.Panels(1).Text = "Read Complete......"
'
'
'   ReDim LastData(CInt(UBound(ItemValues))) As Variant
    Text1.Text = Time
    Label1.Caption = s
    StatusBar1.Panels(3).Text = Time
'
    s = Second(Time)
    m = Minute(Time)
    h = Hour(Time)


   If (s = 0 And m = 0 And h = 8) Or (h = 0 And m = 0 And s = 0) Or (h = 16 And m = 0 And s = 0) Then

    Dim MyData As String
    Dim MyTable(4) As String
    Dim xx As Integer
    Dim xxx(50) As Variant
    Dim myArray As Variant
    Dim j As Integer
    Dim DB_ep() As Variant
    Dim ModNub As Date
    ModNub = Now
    xx = 8
    MyData = "database\OPC1.mdb"
    MyTable(1) = "��ĥϵͳ"
    MyTable(2) = "Ҥϵͳ"
    MyTable(3) = "ˮ��Aĥ"
    MyTable(4) = "ˮ��Bĥ"
    DB_ep = CheckData(MyTable(1), h)
Call SaveData(MyTable(1), 1, 19, -3, DB_ep, ModNub)
    DB_ep = CheckData(MyTable(2), h)
Call SaveData(MyTable(2), 22, 57, 18, DB_ep, ModNub)
    DB_ep = CheckData(MyTable(3), h)
Call SaveData(MyTable(3), 60, 84, 56, DB_ep, ModNub)
    DB_ep = CheckData(MyTable(4), h)
Call SaveData(MyTable(4), 87, 111, 83, DB_ep, ModNub)
'Call SaveData(MyTable, 1, 21, -1,h)
'Call SaveData("tagvalue1", 22, 59, 20,h)
'Call SaveData("", 60, 86, 58,h)
'Call SaveData("", 87, 113, 85,h)
'Call SaveData("tagvalue1", 9, 16, 8, DB_ep, 16)
End If
End Sub

Private Sub objtestgrp_AsyncWriteComplete( _
    ByVal TransactionID As Long, ByVal NumItems As Long, _
    ClientHandles() As Long, Errors() As Long)

End Sub

Private Sub objTestGrp_DataChange( _
    ByVal TransactionID As Long, ByVal NumItems As Long, _
    ClientHandles() As Long, ItemValues() As Variant, _
    Qualities() As Long, TimeStamps() As Date)
    Dim strBuf As String
    Dim nWidth As Integer
    Dim nHeight As Integer
    Dim nDrawHeight As Integer
    Dim sglScale As Single
    Dim x1 As Integer
    Dim x2 As Integer
    Dim Index As Integer
'    ReDim LastData(CInt(UBound(ItemValues))) As Variant
    Text1.Text = Time
    Label1.Caption = s
    StatusBar1.Panels(3).Text = Time
   ' x = x + 1
    s = Second(Time)
    m = Minute(Time)
    h = Hour(Time)

'    For x1 = 1 To UBound(ClientHandles)
'    x2 = CInt(ClientHandles(x1))
'    LvListView.ListItems(x2).SubItems(4) = ClientHandles(x1)
'
'    Next x1
    
        For i = 1 To UBound(ItemValues)
         x2 = CInt(ClientHandles(i))
          Set itx = LvListView.ListItems(x2)
         itx.SubItems(3) = ItemValues(i)
         Next i

        

  
   If (s = 0 And m = 0 And h = 8) Or (h = 0 And m = 0 And s = 0) Or (h = 16 And m = 0 And s = 0) Then

    Dim MyData As String
    Dim MyTable(4) As String
    Dim xx As Integer
    Dim xxx(50) As Variant
    Dim myArray As Variant
    Dim j As Integer
    Dim DB_ep() As Variant
    Dim ModNub As Date
    ModNub = Now
    xx = 8
    MyData = "database\OPC1.mdb"
    MyTable(1) = "��ĥϵͳ"
    MyTable(2) = "Ҥϵͳ"
    MyTable(3) = "ˮ��Aĥ"
    MyTable(4) = "ˮ��Bĥ"
    DB_ep = CheckData(MyTable(1), h)
Call SaveData(MyTable(1), 1, 19, -3, DB_ep, ModNub)
    DB_ep = CheckData(MyTable(2), h)
Call SaveData(MyTable(2), 22, 57, 18, DB_ep, ModNub)
    DB_ep = CheckData(MyTable(3), h)
Call SaveData(MyTable(3), 60, 84, 56, DB_ep, ModNub)
    DB_ep = CheckData(MyTable(4), h)
Call SaveData(MyTable(4), 87, 111, 83, DB_ep, ModNub)
'Call SaveData(MyTable, 1, 21, -1,h)
'Call SaveData("tagvalue1", 22, 59, 20,h)
'Call SaveData("", 60, 86, 58,h)
'Call SaveData("", 87, 113, 85,h)
'Call SaveData("tagvalue1", 9, 16, 8, DB_ep, 16)
End If
End Sub

