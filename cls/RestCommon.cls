VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RestCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'�����t���R���p�C��
#Const Release = 1 '0:�u�Q�Ɛݒ�v���g�����C���e���Z���X���g��
#Const DebugLevel = 0

'
' REST API ���ʃN���X
'
'  �C���X�^���X�̍쐬���@
'    Dim comRest As RestCommon
'    Set comRest = new RestCommon
'
'    Dictionary.cls �� JsonConverter.vba ���ʓr�K�v(GitHub)
'

'���N���X�� static �ϐ����`
'  �P���ɓǂݏ����\�ɂ���Ȃ� Public
'  �����̐�������肽���Ȃ�AProperty �Ō��J����
'
'  HTTP �w�b�_�[
'  �Œ�w�b�_�[���Q��ޗp�ӂ���
'  �ς̃w�b�_�[�͎��s���Ƀp�����[�^�œn���Ă��ǂ�
Private myHeaders As Dictionary '�A�v���Œ�̃w�b�_�[
Private myAuthHeaders As Dictionary '�F�؂Ɋւ��w�b�_�[

Private httpReqStatus As Integer 'http ���N�G�X�g�̉���(�Ō�̒l)

'���v���p�e�B�[
'  {Set|Let}/Get���g��
'  �I�u�W�F�N�g�Ȃ� Set �ϐ��Ȃ� Let

'���߂� http request status �̎擾
' 200�Ȃ�ʏ�OK
Public Property Get status() As Integer
    status = httpReqStatus
End Property


'���N���X�̃R���X�g���N�^
Private Sub Class_Initialize()
    '
End Sub

'���N���X�̃f�X�g���N�^
Private Sub Class_Terminate()
    '
End Sub

'�����\�b�h
'�l��Ԃ��ꍇ�� Function ���g��

'http�w�b�_�[�̒ǉ�
Public Sub AddHeader(key As String, value As String)
    If myHeaders Is Nothing Then
        Set myHeaders = New Dictionary
    End If
    myHeaders.Add key, value
End Sub

'http�w�b�_�[�N���A
Public Sub ClearHeaders()
    Set myHeaders = Nothing
End Sub

'Auth http�w�b�_�[�̒ǉ�
Public Sub AddAuthHeader(key As String, value As String)
    If myAuthHeaders Is Nothing Then
        Set myAuthHeaders = New Dictionary
    End If
    myAuthHeaders.Add key, value
End Sub

'Auth http�w�b�_�[�N���A
Public Sub ClearAuthHeaders()
    Set myAuthHeaders = Nothing
End Sub

'
'�� Get���s(JSON)
' ������JSON �̓I�u�W�F�N�g(Dictionary�^)�B
' JSON ���s�v�Ȃ� execGet() ���g��
' url �G���R�[�h���K�v�Ȃ�A���O�� encodeurl(url) �ŕϊ����Ă���(Excel2013�ȍ~�̕W���֐�)
' ���ʂ́Arul �p�����[�^�������g���ꍇ���������AJSON �ł��n����Ƃ��������(ex. Kintone)
Public Function ExecGetJson(url As String, objJson As Object) As String
    
    ExecGetJson = execMethodJson(url, "GET", objJson)

End Function

'
'�� POST ���s(JSON)
'   ������ JSON �̓I�u�W�F�N�g
'
Public Function ExecPostJson(url As String, objJson As Object) As String
    
    ExecPostJson = execMethodJson(url, "POST", objJson)

End Function

'
'�� PUT ���s(JSON)
'   ������ JSON �̓I�u�W�F�N�g
'
Public Function ExecPutJson(url As String, objJson As Object) As String
    
    ExecPutJson = execMethodJson(url, "PUT", objJson)
    
End Function

'
'�� DELETE ���s(JSON)
'   ������ JSON �̓I�u�W�F�N�g
'
Public Function ExecDeleteJson(url As String, objJson As Object) As String
    
    ExecDeleteJson = execMethodJson(url, "DELETE", objJson)

End Function

'
'�� JSON ���M����
'   ������ json �̓I�u�W�F�N�g
'   http �̃��X�|���X�������Ԃ�
'   POST/PUT/DELETE�ŋ���
'
Private Function execMethodJson(url As String, method As String, objJson As Object) As String
    Dim strJson As String
    Dim resBody As String
    Dim headers As Dictionary
    
    Set headers = New Dictionary
    headers.Add "Content-Type", "application/json"
    'Json�I�u�W�F�N�g���當����ɕϊ�
    strJson = ConvertToJson(objJson)
    resBody = http_request(url, method, headers, strJson)
    execMethodJson = resBody

End Function


'�� Put���s(��JSON)
'  �]��g�����͖������낤
Public Sub ExecPut(url As String)

    Call http_request(url, "PUT")

End Sub

'�� Delete���s(��JSON)
'  �]��g�����͖������낤
Public Sub ExecDelete(url As String)

    Call http_request(url, "DELETE")

End Sub

'�� Get���s(��JSON)
' url �G���R�[�h���K�v�Ȃ�A���O�� encodeurl(url) �ŕϊ����Ă���(Excel2013�ȍ~�̕W���֐�)
' JSON ��n���K�v������Ȃ�AexecGetJson() ���g��
'
Public Function ExecGet(url As String) As String
    Dim resBody As String
    
    resBody = http_request(url, "GET")
    ExecGet = resBody

End Function

'
' �� httpclient �ł� http ���N�G�X�g
'  ���� header ���ݒ�
'  body �������Ԃ�
Private Function http_request(url As String, method As String, Optional headers As Dictionary, Optional request As String) As String
#If Release = 1 Then
    Dim httpReq As Object    'XMLHTTP �I�u�W�F�N�g
#Else
    Dim httpReq As MSXML2.xmlhttp    'XMLHTTP �I�u�W�F�N�g
#End If
    Dim item As Variant
    Dim reqBody As String
    Dim resBody As String
    Dim i As Integer
    Dim key As String

    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    httpReq.Open method, url, False

    '���ʃw�b�_�[
    If Not (myHeaders Is Nothing) Then
        For i = 0 To myHeaders.Count - 1
            key = myHeaders.Keys(i)
            httpReq.setRequestHeader key, myHeaders(key)
        Next
    End If
    
    '�F�؃w�b�_�[
    If Not (myAuthHeaders Is Nothing) Then
        For i = 0 To myAuthHeaders.Count - 1
            key = myAuthHeaders.Keys(i)
            httpReq.setRequestHeader key, myAuthHeaders(key)
        Next
    End If
    
    '�ʐݒ�w�b�_�[
    If Not (headers Is Nothing) Then
        For i = 0 To headers.Count - 1
            key = headers.Keys(i)
            httpReq.setRequestHeader key, headers(key)
        Next
    End If
    
    '���N�G�X�gbody
    If Not IsMissing(request) Then
        reqBody = request
    End If
    
    '���N�G�X�g���M
    httpReq.send (reqBody)

    '���X�|���X�擾
    resBody = httpReq.responseText

    httpReqStatus = httpReq.status '�v���p�e�B�ŎQ�Ɖ\�ɂ��Ă���
    If httpReq.status <> 200 Then
        ' �G���[����
    End If

    Set httpReq = Nothing
    
    http_request = resBody

End Function

'
' �� Base64 �G���R�[�h
'    http�w�b�_�̃p�X���[�h�̃G���R�[�h�Ɏg�p
'    �ėp�I�Ɏg���ɂ́A�o�C�i������̕ϊ����o����悤�ɕ������ق����ǂ�
'
Public Function EncodeBase64(str As String) As String
#If Release = 1 Then
    Dim doc As Object
    Dim elm As Object
#Else
    Dim doc As MSXML2.DOMDocument
    Dim elm As MSXML2.IXMLDOMElement
#End If
    Dim byteArray() As Byte
    Dim retStr As String
    
    byteArray = StrConv(str, vbFromUnicode)
    
    Set doc = CreateObject("MSXML2.DOMDocument")
    Set elm = doc.createElement("b64")

    elm.DataType = "bin.base64"
    elm.nodeTypedValue = byteArray
    retStr = elm.text

    Set elm = Nothing
    Set doc = Nothing
    
    '�u���s�v�폜
    EncodeBase64 = Replace$(retStr, vbLf, "")
End Function


'
' �� URL �G���R�[�h
'    Excel �̕W���֐������b�v���Ă��邾��
'
Public Function Encode_url(str As String) As String

    Encode_url = Application.WorksheetFunction.encodeUrl(str)
    
End Function


'�� �L�[���`�F�b�N�����v�f�̒l�擾

' ����������Dictionary�Ȃ̂Ŗ{���]�|�Ȃ̂����A�L�[�������ꍇ�̃G���[�������̂�
' �����ăL�[�̑��݂��`�F�b�N���Ă���l���擾������@��񋟂���B
' ���x��v���Ȃ��ꍇ�Ɏg���B
Public Function GetDictionaryValue(dic As Dictionary, key As Variant) As Variant
 
    Dim item As Variant
    For i = 0 To dic.Count - 1
        If key = dic.Keys(i) Then
            GetDictionaryValue = dic(key)
            Exit Function
        End If
    Next
    MsgBox "Dictionary �ɃL�[�F" & key & "������܂���I"
    getDictionarvValu = Nothing
End Function

