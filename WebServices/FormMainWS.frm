VERSION 5.00
Begin VB.Form FormMainWS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB6 Mercury - WS"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   13155
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCmdStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   4215
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "FormMainWS.frx":0000
      Top             =   480
      Width           =   5415
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "FormMainWS.frx":000E
      Top             =   480
      Width           =   5415
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Credit"
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "CmdStatus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Request"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Response"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "FormMainWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload FormMainWS
End Sub
Private Sub cmdTest_Click()
  Call ProcessCredit
End Sub
Public Sub ProcessCredit()
    Dim sURL As String
    Dim sEnv As String
    Dim sResp As String
    Dim xmlHtp As New MSXML2.XMLHTTP40
    
    sURL = "https://w1.mercurydev.net/ws/ws.asmx"
    
    sEnv = "<?xml version=""1.0"" encoding=""utf-8""?>"
    sEnv = sEnv & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:mer=""http://www.mercurypay.com"">"
    sEnv = sEnv & "<soapenv:Header/>"
    sEnv = sEnv & "<soapenv:Body>"
    sEnv = sEnv & "  <CreditTransaction xmlns=""http://www.mercurypay.com"">"
    sEnv = sEnv & "    <tran>&lt;?xml version=""1.0""?&gt;"
    sEnv = sEnv & "          &lt;TStream&gt;"
    sEnv = sEnv & "            &lt;Transaction&gt;"
    sEnv = sEnv & "                &lt;MerchantID&gt;595901&lt;/MerchantID&gt;"
    sEnv = sEnv & "                &lt;TranType&gt;Credit&lt;/TranType&gt;"
    sEnv = sEnv & "                &lt;TranCode&gt;Sale&lt;/TranCode&gt;"
    sEnv = sEnv & "                &lt;InvoiceNo&gt;123&lt;/InvoiceNo&gt;"
    sEnv = sEnv & "                &lt;RefNo&gt;123&lt;/RefNo&gt;"
    sEnv = sEnv & "                &lt;Memo&gt;WS-Test&lt;/Memo&gt;"
    sEnv = sEnv & "                &lt;Frequency&gt;OneTime&lt;/Frequency&gt;"
    sEnv = sEnv & "                &lt;PartialAuth&gt;Allow&lt;/PartialAuth&gt;"
    sEnv = sEnv & "                &lt;Amount&gt;"
    sEnv = sEnv & "                    &lt;Purchase&gt;5.00&lt;/Purchase&gt;"
    sEnv = sEnv & "                &lt;/Amount&gt;"
    sEnv = sEnv & "                &lt;Account&gt;"
    sEnv = sEnv & "                    &lt;AcctNo&gt;4003000123456781&lt;/AcctNo&gt;"
    sEnv = sEnv & "                    &lt;Name&gt;CJennings&lt;/Name&gt;"
    sEnv = sEnv & "                    &lt;ExpDate&gt;1215&lt;/ExpDate&gt;"
    sEnv = sEnv & "                &lt;/Account&gt;"
    sEnv = sEnv & "            &lt;/Transaction&gt;"
    sEnv = sEnv & "        &lt;/TStream&gt;"
    sEnv = sEnv & "    </tran>"
    sEnv = sEnv & "    <pw>xyz</pw>"
    sEnv = sEnv & "  </CreditTransaction>"
    sEnv = sEnv & "</soapenv:Body>"
    sEnv = sEnv & "</soapenv:Envelope>"
    
    With xmlHtp
        .Open "post", sURL, False
        .setRequestHeader "Host", "w1.mercurypay.com"
        .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        .setRequestHeader "soapAction", "http://www.mercurypay.com/CreditTransaction"
        Text2.Text = sEnv
        .send sEnv
        sResp = .responseText
        Text1.Text = .responseText
        found1 = InStr(1, sResp, "CmdStatus", vbTextCompare)
        If found1 > 0 Then 'Found
            found1 = found1 + 13
            found2 = InStr(found1, sResp, "/CmdStatus", vbTextCompare)
            If found2 > found1 Then 'Found
                found2 = found2 - 4
                flength = found2 - found1
                txtCmdStatus.Text = Mid$(sResp, found1, flength)
            End If
        End If
    End With
End Sub
