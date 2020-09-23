VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "sp_Interface Generator"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   13830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdCopyToClip 
      Caption         =   "Copy to Clipboard"
      Height          =   375
      Index           =   1
      Left            =   11880
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox ReturnRecordset 
      Caption         =   "Return A Recordset"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox textArea 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   5160
      Width           =   13215
   End
   Begin VB.TextBox ProcName 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Text            =   "pr_webcfg_mid_get_DBConnStr"
      Top             =   1200
      Width           =   7215
   End
   Begin VB.TextBox ConnectionString 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   11535
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2895
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   13215
      ExtentX         =   23310
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "Procedure Name"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   12975
   End
   Begin VB.Label Label1 
      Caption         =   "Connection String"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_strProcName
Public m_strConnectionString
Public m_blnRecordset
Public m_blnGenerate



Private Sub cmdCopyToClip_Click(Index As Integer)
    Clipboard.Clear
    Clipboard.SetText (textArea.Text)
End Sub



Private Sub cmdGenerate_Click()

    m_strProcName = ConnectionString.Text
    m_strConnectionString = ProcName.Text
    m_blnGenerate = True
    WebBrowser1.Navigate2 "about:blank"

End Sub

Private Sub Form_Load()
    ConnectionString.Text = "Provider=SQLOLEDB; Data Source=THENAMEOFYOURSQLSERVERHERE; Initial Catalog=DATABASENAMEHERE; User ID=USERIDHERE; Password=PASSWORDHERE"
    ProcName.Text = "usp_s_ADDRESSES"
    m_blnGenerate = False
    WebBrowser1.Navigate2 "about:blank"
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    'this is how you place the string in a webbrowser control
    'without createing a file
    'I did this in DocumentComplete event, but it could be done
    'in any module
'    If URL = "about:blank" Then
    If m_blnGenerate Then
        m_blnGenerate = False
        WebBrowser1.Document.Body.innerhtml = ShowProcInfo()
    End If
End Sub


Function ShowProcInfo()
  
    'Declare Formatted Data String
    Dim strHTML
    Dim strTextArea
    
    strHTML = ""
    strTextArea = "'<!--#include Virtual=""ADOVBS.inc""-->" & vbCrLf
    strTextArea = strTextArea & "'<!--#include Virtual=""PHPVBLib.asp""-->" & vbCrLf
    
    'Declare and create the command object
    Dim cmd
    Set cmd = CreateObject("ADODB.Command")
    
    'Open the connection on the command by assigning the
    'connection string to the ActiveConnection property
    m_strConnectionString = ConnectionString.Text
    m_strProcName = ProcName.Text
    
    cmd.ActiveConnection = m_strConnectionString
    cmd.CommandType = 4   'Stored Procedure Command Type
    
    'Set the CommandText to the proc name
    cmd.CommandText = m_strProcName
    
    'Call refresh to retrieve the values
    cmd.Parameters.Refresh

    strHTML = strHTML & "<HTML>" & vbCrLf
    strHTML = strHTML & "<HEAD></HEAD>" & vbCrLf
    strHTML = strHTML & "<BODY bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0>" & vbCrLf
    strHTML = strHTML & "<STYLE>" & vbCrLf
    strHTML = strHTML & "TD{" & vbCrLf
    strHTML = strHTML & "  FONT-SIZE: 8pt;" & vbCrLf
    strHTML = strHTML & "}" & vbCrLf
    strHTML = strHTML & "</STYLE>" & vbCrLf

    strHTML = strHTML & "<table width=""100%"" border=""1"">" & vbCrLf
    strHTML = strHTML & "<tr style=""background-color:lightblue;"">" & vbCrLf
    strHTML = strHTML & "<th colspan=""6"">Proc Name = " & m_strProcName & "</th>" & vbCrLf
    strHTML = strHTML & "</tr>" & vbCrLf
    strHTML = strHTML & "<tr style=""background-color:lightblue;"">" & vbCrLf
    strHTML = strHTML & "<th>Parameter Name</th>" & vbCrLf
    strHTML = strHTML & "<th>Direction</th>" & vbCrLf
    strHTML = strHTML & "<th>Type</th>" & vbCrLf
    strHTML = strHTML & "<th>Precision</th>" & vbCrLf
    strHTML = strHTML & "<th>Size</th>" & vbCrLf
    strHTML = strHTML & "<th>Value</th>" & vbCrLf
    strHTML = strHTML & "</tr>" & vbCrLf
    
    Dim blnTR1
    Dim param
    
    For Each param In cmd.Parameters
        If blnTR1 Then
            strHTML = strHTML & "<TR style=""background-color:silver;"">" & vbCrLf
        Else
            strHTML = strHTML & "<TR>" & vbCrLf
        End If
        
        blnTR1 = Not blnTR1
        
        strHTML = strHTML & "" & _
            "<TD align=""left"">&nbsp;" & param.Name & "&nbsp;</TD>" & _
            "<TD align=""center"">&nbsp;" & GetParameterDirectionEnum(param.Direction) & _
                "&nbsp;(" & param.Direction & ")&nbsp;</TD>" & _
            "<TD align=""center"">&nbsp;" & GetDataTypeEnum(param.Type) & _
                "&nbsp;(" & param.Type & ")&nbsp;</TD>" & _
            "<TD align=""center"">&nbsp;" & param.Precision & "&nbsp;</TD>" & _
            "<TD align=""center"">&nbsp;" & param.Size & "&nbsp;</TD>" & _
            "<TD align=""center"">&nbsp;" & param.Value & "&nbsp;</TD>" & vbCrLf & _
            "</TR>" & vbCrLf
    Next

    'strHTML = strHTML & "<tr>" & vbCrLf
    'strHTML = strHTML & "<td colspan=""6"">" & vbCrLf
    'strHTML = strHTML & "<TEXTAREA style=""width:100%;height:100%"" id=textarea1 name=textarea1>" & vbCrLf

    Dim blnFirstParameter   'Is this is the first parameter
    Dim strDeclaration      'Function declaration
    Dim strCommandParameters    'Parameters for the command
    Dim strOutputParameters 'Retrieving of output parameters
    Dim strPrecisionParameters  'Setting of precision for Decimal and Numeric
    Dim strTempParamVarName 'The variable name for the parameter
    
    'Default the Function name to the proc name
    strDeclaration = "Function " & m_strProcName & "("
    
    blnFirstParameter = True
    
    m_blnRecordset = ReturnRecordset.Value = 1
    
    If m_blnRecordset = True Then
        strDeclaration = strDeclaration & "rst"
        blnFirstParameter = False
    End If
    
    For Each param In cmd.Parameters

        If Left(param.Name, 1) = "@" Then
            strTempParamVarName = Mid(param.Name, 2)
        Else
            strTempParamVarName = param.Name
        End If

        If Not param.Direction = 4 Then
        
            If Not blnFirstParameter = True Then
                strDeclaration = strDeclaration & ", "
            Else
                blnFirstParameter = False
            End If
            
            strDeclaration = strDeclaration & strTempParamVarName
            
            If param.Direction = 3 Then
                strOutputParameters = strOutputParameters & vbTab & strTempParamVarName & _
                        " = cmd.Parameters(""" & param.Name & """).Value" & vbCrLf
            End If
        End If
        
        strCommandParameters = strCommandParameters & _
                    vbTab & "cmd.Parameters.Append cmd.CreateParameter(""" & param.Name _
                    & """, " & GetDataTypeEnum(param.Type) & ", " & GetParameterDirectionEnum(param.Direction) & _
                    ", " & param.Size & ", " & strTempParamVarName & ")" & vbCrLf
        
        If param.Type = 14 Or param.Type = 131 Then
            strPrecisionParameters = strPrecisionParameters & "cmd.Parameters(""" & _
                param.Name & """).Precision = " & param.Precision & vbCrLf
        End If
        
    Next
    
    strDeclaration = strDeclaration & ")"

    strTextArea = strTextArea & strDeclaration & vbCrLf & vbCrLf
    strTextArea = strTextArea & vbTab & "Dim cmd                 '- Command Object" & vbCrLf
    strTextArea = strTextArea & vbTab & "Dim RETURN_VALUE        '- Return Value" & vbCrLf
    strTextArea = strTextArea & vbCrLf
    
    strTextArea = strTextArea & vbTab & "RETURN_VALUE         = Null" & vbCrLf
    strTextArea = strTextArea & vbTab & "Set cmd              = Server.CreateObject(""ADODB.Command"")" & vbCrLf
    
    If m_blnRecordset = True Then
        strTextArea = strTextArea & vbTab & "Set rst              = Server.CreateObject(""ADODB.Recordset"")" & vbCrLf
    End If
    
    strTextArea = strTextArea & vbTab & "cmd.ActiveConnection = """ & m_strConnectionString & """" & vbCrLf
    strTextArea = strTextArea & vbTab & "cmd.CommandType      = &H0004      '- adCmdStoredProc" & vbCrLf
    strTextArea = strTextArea & vbTab & "cmd.CommandText      = """ & m_strProcName & """" & vbCrLf & vbCrLf
    
    strTextArea = strTextArea & strCommandParameters
    strTextArea = strTextArea & strPrecisionParameters
        
    If m_blnRecordset = True Then
        strTextArea = strTextArea & vbTab & "rst.CursorLocation = 3  'adUseClient" & vbCrLf
        strTextArea = strTextArea & vbTab & "rst.Open cmd, , 3, 1    'adOpenStatic, adLockReadOnly" & vbCrLf & vbCrLf
        strTextArea = strTextArea & vbTab & "Set rst.ActiveConnection = Nothing  'disconnect the recordset" & vbCrLf
    Else
        strTextArea = strTextArea & vbTab & "cmd.Execute" & vbCrLf & vbCrLf
    End If
    
    strTextArea = strTextArea & strOutputParameters
    strTextArea = strTextArea & vbTab & m_strProcName & " = cmd.Parameters(""RETURN_VALUE"").Value" & _
                                        vbCrLf & vbCrLf
    
    strTextArea = strTextArea & vbTab & "Set cmd = Nothing" & vbCrLf & vbCrLf
    strTextArea = strTextArea & "End Function" & vbCrLf
            
    'strHTML = strHTML & "</TEXTAREA>" & vbCrLf
    'strHTML = strHTML & "    </td>" & vbCrLf
    'strHTML = strHTML & "</tr>" & vbCrLf
    strHTML = strHTML & "</table>" & vbCrLf

    Set cmd = Nothing

    textArea.Text = strTextArea
    ShowProcInfo = strHTML

End Function



Function GetParameterDirectionEnum(lngDirection)
    Select Case lngDirection
        Case 0  'adParamUnknown
            GetParameterDirectionEnum = "adParamUnknown"
        Case 1  'adParamInput
            GetParameterDirectionEnum = "adParamInput"
        Case 2  'adParamOutput
            GetParameterDirectionEnum = "adParamOutput"
        Case 3  'adParamInputOutput
            GetParameterDirectionEnum = "adParamInputOutput"
        Case 4  'adParamReturnValue
            GetParameterDirectionEnum = "adParamReturnValue"
        Case Else
                        GetParameterDirectionEnum = "<B>Direction Not Found</B>"
    End Select
End Function


Function GetDataTypeEnum(lngType)
    Select Case lngType
        Case 0
            GetDataTypeEnum = "adEmpty"
        Case 2
            GetDataTypeEnum = "adSmallInt"
        Case 3
            GetDataTypeEnum = "adInteger"
        Case 4
            GetDataTypeEnum = "adSingle"
        Case 5
            GetDataTypeEnum = "adDouble"
        Case 6
            GetDataTypeEnum = "adCurrency"
        Case 7
            GetDataTypeEnum = "adDate"
        Case 8
            GetDataTypeEnum = "adBSTR"
        Case 9
            GetDataTypeEnum = "adIDispatch"
        Case 10
            GetDataTypeEnum = "adError"
        Case 11
            GetDataTypeEnum = "adBoolean"
        Case 12
            GetDataTypeEnum = "adVariant"
        Case 13
            GetDataTypeEnum = "adIUnknown"
        Case 14
            GetDataTypeEnum = "adDecimal"
        Case 16
            GetDataTypeEnum = "adTinyInt"
        Case 17
            GetDataTypeEnum = "adUnsignedTinyInt"
        Case 18
            GetDataTypeEnum = "adUnsignedSmallInt"
        Case 19
            GetDataTypeEnum = "adUnsignedInt"
        Case 20
            GetDataTypeEnum = "adBigInt"
        Case 21
            GetDataTypeEnum = "adUnsignedBigInt"
        Case 64
            GetDataTypeEnum = "adFileTime"
        Case 72
            GetDataTypeEnum = "adGUID"
        Case 128
            GetDataTypeEnum = "adBinary"
        Case 129
            GetDataTypeEnum = "adChar"
        Case 130
            GetDataTypeEnum = "adWChar"
        Case 131
            GetDataTypeEnum = "adNumeric"
        Case 132
            GetDataTypeEnum = "adUserDefined"
        Case 133
            GetDataTypeEnum = "adDBDate"
        Case 134
            GetDataTypeEnum = "adDBTime"
        Case 135
            GetDataTypeEnum = "adDBTimeStamp"
        Case 136
            GetDataTypeEnum = "adChapter"
        Case 138
            GetDataTypeEnum = "adPropVariant"
        Case 139
            GetDataTypeEnum = "adVarNumeric"
        Case 200
            GetDataTypeEnum = "adVarChar"
        Case 201
            GetDataTypeEnum = "adLongVarChar"
        Case 202
            GetDataTypeEnum = "adVarWChar"
        Case 203
            GetDataTypeEnum = "adLongVarWChar"
        Case 204
            GetDataTypeEnum = "adVarBinary"
        Case 205
            GetDataTypeEnum = "adLongVarBinary"
        Case 8192
            GetDataTypeEnum = "adArray"
        Case Else
            GetDataTypeEnum = "<B>Type Constant Not Found</B>"
    End Select
End Function

