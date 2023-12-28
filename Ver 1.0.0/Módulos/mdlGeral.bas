Attribute VB_Name = "mdlGeral"
Option Explicit
Public usuario As String
Public senha As String
Public cn As ADODB.Connection
Public csql As String
Public Auxiliar As String

Public Sub Main()
    FormLogin.Show 1
End Sub

Public Function RemoveAcento(ByVal sString As String, _
                             Optional bPulaPipe As Boolean, _
                             Optional bRemoveVirgula As Boolean = False) As String
100     If Trim(sString) = "" Then Exit Function
101     RemoveAcento = Replace$(Trim(sString), "á", "a")
102     RemoveAcento = Replace(RemoveAcento, "Ã‡Ãƒ", "CA")
103     If bRemoveVirgula = True Then
104         RemoveAcento = Replace$(RemoveAcento, ",", "")
105     End If
106     RemoveAcento = Replace(RemoveAcento, "Ã­", "i")
107     RemoveAcento = Replace(RemoveAcento, "A‰", "e")
108     RemoveAcento = Replace$(RemoveAcento, "ñ", "n")
109     RemoveAcento = Replace$(RemoveAcento, "à", "a")
110     RemoveAcento = Replace$(RemoveAcento, "â", "a")
111     RemoveAcento = Replace$(RemoveAcento, "ã", "a")
112     RemoveAcento = Replace$(RemoveAcento, "ä", "a")
113     RemoveAcento = Replace$(RemoveAcento, "é", "e")
114     RemoveAcento = Replace$(RemoveAcento, "è", "e")
115     RemoveAcento = Replace$(RemoveAcento, "ê", "e")
116     RemoveAcento = Replace$(RemoveAcento, "ë", "e")
117     RemoveAcento = Replace$(RemoveAcento, "í", "i")
118     RemoveAcento = Replace$(RemoveAcento, "ì", "i")
119     RemoveAcento = Replace$(RemoveAcento, "î", "i")
120     RemoveAcento = Replace$(RemoveAcento, "ï", "i")
121     RemoveAcento = Replace$(RemoveAcento, "ó", "o")
122     RemoveAcento = Replace$(RemoveAcento, "ò", "o")
123     RemoveAcento = Replace$(RemoveAcento, "ô", "o")
124     RemoveAcento = Replace$(RemoveAcento, "õ", "o")
125     RemoveAcento = Replace$(RemoveAcento, "ö", "o")
126     RemoveAcento = Replace$(RemoveAcento, "ú", "u")
127     RemoveAcento = Replace$(RemoveAcento, "ù", "u")
128     RemoveAcento = Replace$(RemoveAcento, "û", "u")
129     RemoveAcento = Replace$(RemoveAcento, "ü", "u")
130     RemoveAcento = Replace$(RemoveAcento, "ç", "c")
131     RemoveAcento = Replace$(RemoveAcento, "Á", "A")
132     RemoveAcento = Replace$(RemoveAcento, "À", "A")
133     RemoveAcento = Replace$(RemoveAcento, "Â", "A")
134     RemoveAcento = Replace$(RemoveAcento, "Ã", "A")
135     RemoveAcento = Replace$(RemoveAcento, "Ä", "A")
136     RemoveAcento = Replace$(RemoveAcento, "É", "E")
137     RemoveAcento = Replace$(RemoveAcento, "È", "E")
138     RemoveAcento = Replace$(RemoveAcento, "Ê", "E")
139     RemoveAcento = Replace$(RemoveAcento, "Ë", "E")
140     RemoveAcento = Replace$(RemoveAcento, "Í", "I")
141     RemoveAcento = Replace$(RemoveAcento, "Ì", "I")
142     RemoveAcento = Replace$(RemoveAcento, "Î", "I")
143     RemoveAcento = Replace$(RemoveAcento, "Ï", "I")
144     RemoveAcento = Replace$(RemoveAcento, "Ñ", "N")
145     RemoveAcento = Replace$(RemoveAcento, "Ó", "O")
146     RemoveAcento = Replace$(RemoveAcento, "Ò", "O")
147     RemoveAcento = Replace$(RemoveAcento, "Ô", "O")
148     RemoveAcento = Replace$(RemoveAcento, "Õ", "O")
149     RemoveAcento = Replace$(RemoveAcento, "Ö", "O")
150     RemoveAcento = Replace$(RemoveAcento, "Ú", "U")
151     RemoveAcento = Replace$(RemoveAcento, "Ù", "U")
152     RemoveAcento = Replace$(RemoveAcento, "Û", "U")
153     RemoveAcento = Replace$(RemoveAcento, "Ü", "U")
154     RemoveAcento = Replace$(RemoveAcento, "Ç", "C")
155     RemoveAcento = Replace$(RemoveAcento, Chr(34), "") '"
156     RemoveAcento = Replace$(RemoveAcento, Chr(39), "") ''
157     RemoveAcento = Replace$(RemoveAcento, "&", "E")
158     RemoveAcento = Replace$(RemoveAcento, vbTab, " ")
159     RemoveAcento = Replace$(RemoveAcento, vbCrLf, " ")
160     RemoveAcento = Replace$(RemoveAcento, "<", "")
161     RemoveAcento = Replace$(RemoveAcento, ">", "")
162     If bPulaPipe = False Then
163         RemoveAcento = Replace$(RemoveAcento, "|", "")
164     End If
165     RemoveAcento = Replace$(RemoveAcento, "§", "")
166     RemoveAcento = Replace$(RemoveAcento, "£", "")
167     RemoveAcento = Replace$(RemoveAcento, "¹", "")
168     RemoveAcento = Replace$(RemoveAcento, "²", "")
169     RemoveAcento = Replace$(RemoveAcento, "³", "")
170     RemoveAcento = Replace$(RemoveAcento, "¢", "")
171     RemoveAcento = Replace$(RemoveAcento, "¬", "")
172     RemoveAcento = Replace$(RemoveAcento, Chr(186), "") 'º
173     RemoveAcento = Replace$(RemoveAcento, Chr(170), "") 'ª
174     RemoveAcento = Replace$(RemoveAcento, "€", "")
175     RemoveAcento = Replace$(RemoveAcento, "¦", "")
176     RemoveAcento = Replace$(RemoveAcento, "µ", "")
177     RemoveAcento = Replace$(RemoveAcento, "´", "")
178     RemoveAcento = Replace$(RemoveAcento, "‡", "A")
179     RemoveAcento = Replace$(RemoveAcento, "?", "")
180     RemoveAcento = Replace$(RemoveAcento, "'", "")
181     RemoveAcento = Replace$(RemoveAcento, "ƒ", "")
182     RemoveAcento = Replace$(RemoveAcento, "'", "")
183     RemoveAcento = Replace$(RemoveAcento, "¼", "")
184     RemoveAcento = Replace$(RemoveAcento, "\", "")
185     RemoveAcento = Replace$(RemoveAcento, "*", "")
186     RemoveAcento = Replace$(RemoveAcento, "°", "")
End Function

Public Function GetString(dbField As ADODB.Field, _
                          Optional tamanhoString As Integer = 0, _
                          Optional bRemoveVirgula As Boolean = False) As String
100     Dim returnString As String
101     returnString = ""
102     If Not dbField Is Nothing Then
103         If Not IsNull(dbField.Value) Then
104             If tamanhoString = 0 Then
105                 returnString = RemoveAcento(CStr(dbField.Value), , bRemoveVirgula)
106             Else
107                 returnString = RemoveAcento(CStr(Left(dbField.Value, tamanhoString)), bRemoveVirgula)
108             End If
109         End If
110     End If
111     If UCase(returnString) = "NULL" Then returnString = ""
112     GetString = UCase(returnString)
End Function
Public Function GetInteger(dbField As ADODB.Field) As Integer
100     Dim returnValue As Integer
101     returnValue = 0
102     If Not dbField Is Nothing Then
103         If Not IsNull(dbField.Value) Then
104             returnValue = Val(dbField.Value)
107         End If
108     End If
109     GetInteger = returnValue
End Function

Public Function GetDate(dbField As ADODB.Field) As Date
100     Dim returnValue As Date
101     returnValue = 0
102     If Not dbField Is Nothing Then
103         If Not IsNull(dbField.Value) Then
104             If IsDate(dbField.Value) Then
105                 returnValue = CDate(dbField.Value)
106             End If
107         End If
108     End If
109     GetDate = returnValue
End Function
Public Function GetBoolean(dbField As ADODB.Field) As Boolean
100     Dim returnValue As Boolean
101     returnValue = False
102     If Not dbField Is Nothing Then
103         If Not IsNull(dbField.Value) Then
104             returnValue = CBool(Val(dbField.Value))
105         End If
106     End If
107     GetBoolean = returnValue
End Function
Public Function GetCurrency(dbField As ADODB.Field) As Currency
100     Dim returnValue As Currency
101     returnValue = 0
102     If Not dbField Is Nothing Then
103         If Not IsNull(dbField.Value) Then
104             returnValue = DblVal(dbField.Value)
105         End If
106     End If
107     GetCurrency = returnValue
End Function

Public Function GetDouble(dbField As ADODB.Field) As Double
100     Dim returnValue As Double
101     returnValue = 0
102     If Not dbField Is Nothing Then
103         If Not IsNull(dbField.Value) Then
104             returnValue = DblVal(dbField.Value)
105         End If
106     End If
107     GetDouble = returnValue
End Function

Public Function GetLong(dbField As ADODB.Field) As Long
100     Dim returnValue As Long
101     returnValue = 0
102     If Not dbField Is Nothing Then
103         If Not IsNull(dbField.Value) Then
104             returnValue = Val(dbField.Value)
105         End If
106     End If
107     GetLong = returnValue
End Function

Public Function DblVal(entrada As Variant) As Double
100     If IsNull(entrada) Then
101         DblVal = 0#
102     Else
103         If Trim(entrada) = "" Then
104             DblVal = 0#
105         Else
106             If IsNumeric(entrada) Then
107                 If InStr(1, CStr(entrada), ".") > 0 And InStr(1, CStr(entrada), ",") > 0 Then
108                     entrada = Replace(entrada, ".", "")
109                 End If
110                 DblVal = CDbl(Replace(entrada, ".", ","))
111             Else
112                 DblVal = 0#
113             End If
114         End If
115     End If
End Function

Public Function ObrigaNumerosInteiros(texto As String) As String
100     If IsNumeric(texto) = False Then
101         texto = ""
102     Else
103         If CBool(InStr(texto, ",")) = True Then
104             texto = ""
105         End If
106         If CBool(InStr(texto, ".")) = True Then
107             texto = ""
108         End If
109         If CBool(InStr(texto, "+")) = True Then
110             texto = ""
111         End If
112         If CBool(InStr(texto, "-")) = True Then
113             texto = ""
114         End If
115     End If

116     If texto = "" Then
117         ObrigaNumerosInteiros = ""
118     Else
119         ObrigaNumerosInteiros = texto
120     End If

End Function
