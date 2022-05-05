Imports System.Math
Imports Microsoft.VisualBasic

Public Class cFormulas
    '--------------------------------------------------------------------------
    ' cFormulas                                                         (06/Feb/99)
    '   Clase para evaluar expresiones

    '--------------------------------------------------------------------------
    ' Revisión de  7/Feb/99 Ahora se evaluan correctamente los operadores con más
    '                       precedencia, en este orden: % ^ * / \ + -
    '                       También se permiten paréntesis no balanceados,
    '                       al menos los de apertura.
    '                       Si faltan los de cierre se añaden al final.
    '                       Se permiten definir Funciones y usar algunas internas
    '                       siempre que no necesiten parámetros (por ahora Rnd)
    ' Revisión de  9/Feb/99 Se pueden usar funciones internas con parámetros:
    '                       Por ahora:
    '                       Int, Fix, Abs, Sgn, Sqr, Cos, Sin, Tan, Atn, Exp, Log
    ' Revisión de  9/Feb/99 Se permiten funciones con un número variable de parámetros
    '                       aunque las operaciones a realizar en los parámetros
    '                       opcionales siempre sea la misma... algo es algo
    '                       Se evalúa la expresión por si hay asignaciones a variables
    '                       y de ser así, se crean esas variables y después se evalúa
    '                       el resto de la expresión.
    '                       Esto permite crear variables en la misma cadena a evaluar,
    '                       sin tener que asignarlas antes.
    '                       Este ejemplo devolverá el valor mayor de X o Y
    '                       X = A+B : Y = C-X : Max(x,y)
    '                       Las asignaciones NO se pueden hacer así:
    '                       x:=10:y:=20:x+y (devolvería 0)
    '                       También permite cadenas de caracteres, aunque lo que haya
    '                       entre comillas simplemente se repetirá:
    '                       x=10:y=20:"El mayor de x e y =";Max(x,y)
    '                       Si se hace esto otro:
    '                       x=10:y=20:s="El mayor de x e y =";Max(x,y):s
    '                       Devolverá:"El mayor de x e y ="20
    '                       Ya que se evalúa como s=Max(x,y)
    '                       Esto otro:
    '                       x=10:y=20:s="El mayor de x e y =";z:z=Max(x,y)
    '                       Mostrará:
    '                       "El mayor de x e y ="
    '                       Ya que se asigna s=z, pero no se muestra el valor de z
    '                       RESUMIENDO:
    '                       Se pueden usar cadenas entre comillas pero no se pueden
    '                       asignar a variables... (al menos por ahora)
    '                       Se puede incrementar una variable en una asignación
    '                       x=10:x=x+5+x sería igual a 10+5+10 = 25
    '                       x=10:x=x+1 sería igual a 10+1 = 11
    ' Revisión de 10/Feb/99 Algunas correcciones de las cadenas y otras cosillas
    '                       Cuando se asigna un valor a una variable existente
    '                       Método para recuperar la fórmula de una función
    ' Revisión de 11/Feb/99 Función de Redondeo
    ' Revisión de 12/Feb/99 Nuevas funciones de la clase y definidas y otras mejoras
    ' Revisión de 13/Feb/99 Comprobación de números con notación científica
    ' Revisión de 14/Feb/99 Acepta hasta 100 parámetros
    ' Revisión de 11/Ene/01 Evaluar correctamente la precedencia en los cálculos
    ' Revisión de 22/Ene/01 Arreglado nuevo bug en Calcular
    ' Revisión de 28/Ene/01 Cambio en la forma de calcular los números,
    '                       los almaceno en Variant para hacer los cálculos con Cdec(
    '                       ya que fallaba con números de notación científica
    ' Revisión de 29/Ene/01 Propiedad para devolver un valor con notación científica
    '                       o decimal, para el caso de valores muy grandes o pequeños
    ' Revisión de 22/Feb/01 Fallaba en cálculos simples como: 3*2+5
    ' Revisión de 29/Oct/02 Arreglo al realizar Instr con una cadena vacía
    ' Revisión de 02/Nov/02 Nuevas funciones Hex, Oct, Round (usando la función de VB)
    '                       Bin, Bin2Dec, Dec2Bin
    '                       No usar notación científica con las funciones Bin...
    '             03/Nov/02 No recalcular las funciones internas, (ver esFunVB)
    '                       Nuevas funciones: Ln, (es igual que Log), Log10, LogX,
    '                       Hex2Dec, Oct2Dec, Dec2Hex (=Hex), Dec2Oct (=Oct)
    '                       Las fórmulas de las funciones internas prevalecen
    '                       sobre los cambios hechos externamente.
    '                       Declaro PI para efectuar la conversión de grados
    '                       a radianes y viceversa (Grados2Radianes, Radianes2Grados)
    '                       Arreglado bug cuando la expresión está entre paréntesis
    '
    '--------------------------------------------------------------------------
    ' Esta es una nueva implementación del módulo Formula.bas y la clase cEvalOp
    ' Aunque los métodos usados son totalmente diferentes y realmente no es una
    ' mejora, están basados en dichos módulos... o casi...
    '

    '--------------------------------------------------------------------------
    'Option Explicit
    'Option Compare Text

    Private mNumFunctions As Long       ' El número de funciones "propias"
    Private esFunVB As Boolean
    Private m_NotacionCientifica As Boolean
    '
    ' Funciones Internas soportadas en el programa,
    ' debe indicarse el paréntesis y un espacio de separación
    'Const FunVBNum As String = "Int( Fix( Abs( Sgn( Sqr( Cos( Sin( Tan( Atn( Exp( Log( Iif( "
    ' Bin( no es una función de VB, pero se usará como si fuera...      (02/Nov/02)
    Const FunVBNum As String = "Int( Fix( Abs( Sgn( Sqr( Cos( Sin( Tan( Atn( Exp( Log( Ln( Log10( Round( Hex( Dec2Hex( Oct( Dec2Oct( Bin2Dec( Hex2Dec( Oct2Dec( "
    ' Símbolos a usar para separar los Tokens
    Private Simbols As String
    ' Signos a usar para comentarios
    Private RemSimbs As String


    Public Class tVariable
        Public Name As String
        Public Value As String
    End Class
    ' Array de variables
    Private aVariables(-1) As tVariable


    Public Class tFunctions
        Public Name As String
        Public Params As String
        Public Formula As String
        'Descripcion As String

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
    ' Array de funciones
    Private aFunctions(-1) As tFunctions

    Public Function MTrim(ByVal sVar As String, Optional ByVal NoEval As String = "") As String
        '--------------------------------------------------------------------------
        ' Quita todos los espacios y blancos del parámetro pasado           (09/Feb/99)
        ' Los parámetros:
        '   sVar    Cadena a la que se quitarán los blancos
        '   NoEval  Si se especifica, pareja de caracteres que encerrarán una cadena
        '           a la que no habrá que quitar los espacios y blancos
        '--------------------------------------------------------------------------
        Dim i As Long
        Dim j As Long
        Dim sTmp As String
        Dim sBlancos As String

        ' Se entienden como blancos: espacios, Tabs y Chr$(0)
        sBlancos = " " & vbTab & Chr(0)
        ' NoEval tendrá el caracter que no se evaluará para quitar espacios
        ' por ejemplo si no queremos quitar los caracteres entre comillas
        ' NoEval será chr$(34), e irá por pares o hasta el final de la cadena

        sTmp = ""
        For i = 1 To Len(sVar)
            ' Si es el caracter a no evaluar
            If Mid$(sVar, i, 1) = NoEval Then
                ' Buscar el siguiente caracter
                j = InStr(i + 1, sVar, NoEval, CompareMethod.Text)
                If j = 0 Then
                    sVar = sVar & NoEval
                    j = Len(sVar)
                End If
                sTmp = sTmp & Mid$(sVar, i, j - i + 1)
                i = j '+ 1
                ' Si no es uno de los caracteres "blancos"
            ElseIf InStr(sBlancos, Mid$(sVar, i, 1)) = 0 Then
                ' Asignarlo a la variable final
                sTmp = sTmp & Mid$(sVar, i, 1)
            End If
        Next
        MTrim = sTmp
    End Function

    Public Function AsignarVariables(ByVal v As String, _
                                     Optional ByVal NoEval As String = "") As String
        '--------------------------------------------------------------------------
        ' Asignar las variables, si las hay                                 (09/Feb/99)
        ' Los parámetros de entrada:
        '   v       Expresión con posibles asignaciones
        '   NoEval  Si se especifica, pareja de caracteres que encerrarán una cadena
        '           en la que no se buscarán variables
        '
        ' Devolverá el resto de la cadena que será la expresión a evaluar
        '--------------------------------------------------------------------------
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim sNom As String
        Dim sVal As String
        Dim sExpr As String
        Dim sValAnt As String
        Dim sOp As String

        ' Quitar todos los espacios, excepto los que estén entre comillas
        v = MTrim(v, NoEval)

        sExpr = ""
        ' Buscar los caracteres a no evaluar y ponerlos después de las asignaciones
        If Len(NoEval) Then
            Do
                j = InStr(v, NoEval)
                ' Si hay caracteres a no evaluar
                If j Then
                    ' Buscar el siguiente caracter NoEval y comprobar si después
                    ' hay asignaciones
                    k = InStr(j + 1, v, NoEval, CompareMethod.Text)
                    If k = 0 Then k = Len(v)
                    sExpr = sExpr & Mid$(v, j, k - j + 1)
                    v = Left$(v, j - 1) & Mid$(v, k + 1)
                End If
            Loop While j
        End If

        ' Buscar el signo de dos puntos que será el separador,
        ' pero hay que tener en cuenta de que las asignaciones pueden ser con :=
        Do
            ' Buscar el siguiente signo igual
            i = InStr(v, "=")
            If i Then
                ' Lo que haya delante debe ser el nombre de la variable
                sNom = Left$(v, i - 1)
                ' Si sNom contiene : lo que haya antes de los dos puntos será
                ' parte de la expresión y el resto será el nombre.
                '*********************************************
                '***  Esto NO permite asignaciones con :=  ***
                '*********************************************
                j = InStr(sNom, ":")
                If j Then
                    sExpr = sExpr & Left$(sNom, j - 1)
                    sNom = Mid$(sNom, j + 1)
                End If
                ' Comprobar si a continuación hay dos puntos,
                ' (será el separador de varias asignaciones)
                j = InStr(i + 1, v, ":", CompareMethod.Text)
                ' Si no hay, tomar la longitud completa restante
                If j = 0 Then j = Len(v) + 1
                ' Asignar el valor desde el signo igual hasta los dos puntos
                ' (o el fin de la cadena, valor de j)
                sVal = Mid$(v, i + 1, j - (i + 1))
                ' Dejar en v el resto de la cadena
                v = Mid$(v, j + 1)
                ' Si ya no hay nada más en la cadena, preparar para salir del Do
                If Len(v) = 0 Then i = 0
                ' Comprobar si está en la lista de variables, si no está, añadirla
                j = IsVariable(sNom)
                If j Then
                    ' Esta variable ya existe, sustituir la expresión
                    '//////////////////////////////////////////////////////////////////
                    ' Si en la expresión asignada está la misma variable
                    ' sustituirla por el valor que tuviera
                    sValAnt = aToken(sVal, sOp)
                    If sValAnt = sNom Then
                        ' Sustituir la variable por el valor
                        sVal = aVariables(j).Value & sOp & sVal
                        'aVariables(j).Value = sVal
                        ' y calcularlo
                        sVal = ParseFormula(sVal)
                        sVal = Calcular(sVal)
                    Else
                        ' En el caso que se reasigne el valor               (10/Feb/99)
                        sVal = sValAnt
                    End If
                    '//////////////////////////////////////////////////////////////////
                    ' Asignar el valor asignado
                    aVariables(j).Value = sVal
                Else
                    ' No existe la variable, añadirla
                    NewVariable(sNom, sVal)
                End If
            End If
        Loop While i
        ' Devolver el resto de la cadena, si queda algo...
        AsignarVariables = sExpr & v
    End Function

    Private Function aToken(ByRef sF As String, ByRef sSimbol As String) As String
        '--------------------------------------------------------------------------
        ' Devuelve el siguiente TOKEN y el Símbolo siguiente
        ' un Token es una variable, instrucción, función o número
        '
        ' Los parámetros se deben especificar por referencia ya que se modifican:
        '   sF          Cadena con la fórmula o expresión a Tokenizar
        '   sSimbol     El símbolo u operador a usar
        ' Se devolverá la cadena con lo hallado o una cadena vacía si no hay nada
        '--------------------------------------------------------------------------
        Dim i As Long
        Dim j As Long
        Dim sSimbs As String

        ' Usar los símbolos normales y los usados para los comentarios
        ' Pero no usar el de comillas dobles para que se puedan usar cadenas...
        'sSimbs = Chr$(34) & " " & Simbols & RemSimbs & " "
        sSimbs = Simbols & RemSimbs & " " & Chr(34) & " "
        'sSimbs = Simbols & RemSimbs & " "

        ' Si la cadena de entrada está vacía o sólo tiene blancos
        If Len(Trim$(sF)) = 0 Then
            aToken = ""
            sF = ""
        Else
            j = MultInStr(sF, sSimbs, sSimbol)
            ' El valor devuelto será el símbolo que esté más a la izquierda
            If j = 0 Then
                ' Devolver la cadena completa
                aToken = sF
                sF = ""
            Else
                ' Si hay algo entre dos comillas dobles devolverlo          (10/Feb/99)
                i = InStr(sF, Chr(34))
                If i Then
                    ' Buscar la siguiente
                    j = InStr(i + 1, sF, Chr(34))
                    If j = 0 Then
                        sF = sF & Chr(34)
                        j = Len(sF)
                    End If
                    aToken = Mid$(sF, i, j - i + 1)
                    sF = Left$(sF, i - 1) & Mid$(sF, j + 1)
                    sSimbol = ""
                    Exit Function
                End If
                ' Devolver lo hallado hasta el token
                aToken = Left$(sF, j - 1)
                sF = Mid$(sF, j + Len(sSimbol))
                ' Si el número está en notación científica xxxEyyy
                If Right$(aToken, 1) = "E" Then
                    ' Comprobar si TODOS los caracteres anteriores a la E   (13/Feb/99)
                    ' son números, ya que una variable o función puede acabar con E
                    sSimbs = Left$(aToken, Len(aToken) - 1)
                    j = 0
                    For i = 1 To Len(sSimbs)
                        If InStr("0123456789", Mid$(sSimbs, i, 1)) Then
                            j = j + 1
                        End If
                    Next
                    ' Si el número de cifras es igual a la longitud de la cadena
                    If j = Len(sSimbs) Then
                        ' Es que DEBERÍA ser un número con notación científica
                        '//////////////////////////////////////////////////////////////
                        ' IMPORTANTE:
                        '   No se procesarán correctamente variables o funciones
                        '   que empiecen por números y acaben con la letra E
                        '//////////////////////////////////////////////////////////////
                        aToken = aToken & sSimbol & Left$(sF, 2)
                        sF = Mid$(sF, 3)
                        sSimbol = ""
                    End If
                End If
            End If
        End If
    End Function

    Public Function MultInStr(ByVal String1 As String, ByVal String2 As String, _
                              Optional ByRef sSimb As String = "", Optional ByVal Start As Long = 1) As Long
        '--------------------------------------------------------------------------
        ' Siempre se especificarán los tres parámetros,
        ' opcionalmente, el último será la posición de inicio o 1 si no se indica
        '
        ' Busca en la String1 cualquiera de los caracteres de la String2,
        ' devolviendo la posición del que esté más a la izquierda.
        ' El parámetro sSep se pasará por referencia y en él se devolverá
        ' el separador hallado.
        '
        ' En String2 se deberán separar con espacios los caracteres a buscar
        '--------------------------------------------------------------------------
        Dim j As Long
        Dim sTmp As String
        ' La posición con un valor menor
        Dim elMenor As Long
        ' Caracter de separación
        Const sSep As String = " "

        ' Hacer un bucle entre cada uno de los valores indicados en String2
        elMenor = 0
        If Start <= Len(String1) Then
            String2 = Trim$(String2) & sSep
            ' Se buscarán todas las subcadenas de String2
            Do
                j = InStr(String2, sSep)
                If j Then
                    sTmp = Left$(String2, j - 1)
                    String2 = Mid$(String2, j + Len(sSep))
                    If Len(sTmp) Then
                        j = InStr(Start, String1, sTmp, CompareMethod.Text)
                    Else
                        j = 0
                    End If
                    If j Then
                        If elMenor = 0 Or elMenor > j Then
                            elMenor = j
                            sSimb = sTmp
                        End If
                        ' Si es la posición de inicio, no habrá ninguno menor
                        ' así que salimos del bucle
                        If elMenor = Start Then
                            String2 = ""
                        End If
                    End If
                Else
                    String2 = ""
                End If
            Loop While Len(String2)
        End If
        MultInStr = elMenor
    End Function

    Public Sub New()
        '--------------------------------------------------------------------------
        ' Iniciar algunos valores y algunas de las funciones internas del VB
        ' soportadas por esta clase, (ver la constante FunVBNum)
        '--------------------------------------------------------------------------
        Dim i As Integer
        Dim sName As String
        Dim sFunVB As String
        '
        ' Por defecto, se devuelven los valores con notación científica (29/Ene/01)
        m_NotacionCientifica = True
        '
        ' Símbolos
        Simbols = ":= < > = >= <= ( ) ^ * / \ - + $ ! # @ { } [ ] "
        ' Comentarios
        RemSimbs = "; ' // "
        ' Inicializar el array con el elemento cero, que no se usará
        'ReDim aVariables(0)

        '--------------------------------------------------------------------------
        ' Quitada esta función por defecto, ya que Sum() hace lo mismo  (14/Mar/99)
        '
        'ReDim aFunctions(0)
        '    ReDim aFunctions(1)
        '    ' Esta debe ser la primera función para uso generérico,         (12/Feb/99)
        '    ' hasta 10 parámetros
        '    With aFunctions(1)
        '        ' Si hay más de 10 parámetros se sumarán a lo que se haya puesto
        '        .Formula = "N1+N2+N3+N4+N5+N6+N7+N8+N9+N0"
        '        .Name = "<Genérica>"
        '        .Params = "N1,N2,N3,N4,N5,N6,N7,N8,N9,N0"
        '    End With
        '--------------------------------------------------------------------------
        ' Funciones con más de un parámetro,                            (09/Feb/99)
        ' incluso un número indefinido.
        ' Los parámetros pueden ser cualquier tipo de expresión que esta clase evalue
        ' Cuando se usa más de un parámetro, asegurarse de que son nombres distintos
        ' por ejemplo "Num, Num" no funcionaría, usar "Num, Num1"
        '--------------------------------------------------------------------------
        ' Suma varios números, admite uno o más parámetros
        ' Suma dos números, también admite sólo un parámetro,
        ' Si se especifica uno, devuelve ese valor... luego no suma
        NewFunction("Sum", "Num1,Num2,...", "Num1+Num2+...")
        ' Resta números
        NewFunction("Subs", "Num1,Num2,...", "Num1-Num2-...")
        ' Multiplica números
        NewFunction("Mult", "Num1,Num2,...", "Num1*Num2*...")
        ' Las funciones que se van a evaluar de forma especial, se deben indicar con @
        ' aunque estas funciones deben estar previamente contempladas y sólo se
        ' evaluan si está el código dentro de la clase...
        ' Max, devuelve el valor mayor de los dos indicados
        NewFunction("Max", "Num1,Num2", "@Max(Num1,Num2)")
        ' Min, devuelve el valor menor de los dos indicados
        NewFunction("Min", "Num1,Num2", "@Min(Num1,Num2)")
        '**************************************************************************
        '*** ATENCION ***
        '****************
        ' Si se usa Max o Min dentro de Max o Min hay que usar el símbolo @
        ' Por ejemplo: (devolvería 20)
        ' Max(@Max(10,20),@Min(5,4))
        '**************************************************************************
        ' Nueva función de redondeo                                     (11/Feb/99)
        'NewFunction "Round", "Num", "Int(Num+0.5)"
        '
        ' Función Rnd del VB
        NewFunction("Rnd", "", "@Rnd")
        '
        ' Esta función TIENE que tener el @                             (03/Nov/02)
        NewFunction("Bin", "Num,Precision", "@Bin(Num,Precision)")
        NewFunction("Dec2Bin", "Num", "Bin(Num)")
        NewFunction("Dec2Bin2", "Num,Precision", "Bin(Num,Precision)")
        '
        ' Logaritmo en base X
        NewFunction("LogX", "Num,Base", "Log(Num)/Log(Base)")
        '
        ' Para efectuar los cáculos en radianes y grados                (03/Nov/02)
        NewVariable("Pi", "3.1415926535897932384626433832795")
        NewFunction("Grados2Radianes", "Num", "Num*(Pi/180)")
        NewFunction("Radianes2Grados", "Num", "Num*(180/Pi)")
        'NewFunction "CosGrados", "Num", "(Cos(Num))*(180/Pi)"
        'NewFunction "CosGrados", "Num", "Cos(Num)*(180/Pi)"
        ' para probar, esto vuelve a ponerlo en radianes
        'NewFunction "Cos2", "Num", "CosGrados(Num)*Pi/180"
        'NewFunction "Cos2", "Num", "Grados2Radianes(CosGrados(Num))"
        '
        ' Añadir las declaradas en la constante FunVar
        ' Las funciones deben estar separadas por espacios y acabar con el (
        ' por ejemplo: "Int( Fix( "
        sFunVB = FunVBNum
        Do
            i = InStr(sFunVB, "( ")
            If i Then
                sName = Left$(sFunVB, i - 1)
                sFunVB = Mid$(sFunVB, i + 2)
                NewFunction(sName, "Num", "@" & sName & "(Num)")
            End If
        Loop While Len(sFunVB)
        ' Añadir otras para que sirvan de ejemplo
        ' No usar el signo @ si hacen uso de algunas de las definidas
        ' sino el resultado sería el de esa función... esto habrá que arreglarlo...
        NewFunction("Sec", "Num", "1/Cos(Num)")
        NewFunction("CoSec", "Num", "1/Sin(Num)")
        NewFunction("CoTan", "Num", "1/Tan(Num)")
        '
        ' El número de funciones propias de la clase,
        ' para evitar que se modifiquen
        mNumFunctions = UBound(aFunctions)
    End Sub

    Public Function IsVariable(ByVal sName As String) As Long
        '--------------------------------------------------------------------------
        ' Comprueba si es una variable,
        ' de ser así, devolverá el índice en el array de variables
        ' o cero si no se ha hallado.
        ' En caso de no hallar la variable, la añade con el valor cero
        '--------------------------------------------------------------------------
        Dim i As Long

        sName = Trim$(sName)
        IsVariable = 0
        If Len(sName) Then
            For i = 1 To UBound(aVariables)
                If aVariables(i).Name = sName Then
                    IsVariable = i
                    Exit For
                End If
            Next
        End If
        ' Si no existe la variable
        ' No se hace nada, que es lo mismo que si no existiera...       (09/Feb/99)
        '    If IsVariable = 0 Then
        '        ' y no es un número
        '        If Len(sName) Then
        '            If Not IsNumeric(sName) Then
        ''                ' Este caso ya no se dará, pero por si las moscas       (08/Feb/99)
        ''                sVar = sName
        ''                If Right$(sName, 1) = "E" Then
        ''                    sName = sName & "-02"
        ''                End If
        ''                ' Por si es un número con E
        ''                BuscarCifra sName, sVal
        ''                NewVariable sVar, sVal
        '                NewVariable sName, "0"
        '                IsVariable = UBound(aVariables)
        '            End If
        '        End If
        '    End If
    End Function

    Public Function IsFunction(ByVal sName As String) As Long
        '--------------------------------------------------------------------------
        ' Comprueba si es una función,
        ' de ser así, devolverá el índice en el array de fórmulas
        ' o cero si no se ha hallado
        '--------------------------------------------------------------------------
        Dim i As Long
        '
        sName = Trim$(sName)
        IsFunction = 0
        If Len(sName) Then
            For i = 1 To UBound(aFunctions)
                If aFunctions(i).Name = sName Then
                    IsFunction = i
                    Exit For
                End If
            Next
        End If
    End Function

    Public Function FunctionVal(ByVal sName As String) As String
        '--------------------------------------------------------------------------
        ' Comprueba si sName contiene una fórmula interna
        ' Estará en el formato: @FormulaInterna
        ' Sólo será válido para funciones que no necesiten parámetros
        ' Ahora se permite un parámetro en las funciones soportadas     (08/Feb/99)
        '--------------------------------------------------------------------------
        Dim i As Long, j As Long
        Dim sValue As String, sParams As String, sNameFun As String
        '
        sName = Trim$(sName)
        i = InStr(sName, "@")
        esFunVB = False
        If i Then
            sValue = Left$(sName, i - 1) & Mid$(sName, i + 1)
            ' Si es Rnd
            i = InStr(sValue, "Rnd")
            ' El formato será Rnd [* valor]
            If i Then
                sName = Left$(sValue, i - 1) & "(" & Rnd() & ")" & Mid$(sValue, i + 3)
            End If
            '//////////////////////////////////////////////////////////////////////////
            ' Si es alguna de las definidas en la constante FunVBNum
            ' sNameFun devolverá el nombre de la función hallada
            i = MultInStr(sValue, FunVBNum, sNameFun)
            ' El formato será NombreFunción(expresión)
            If i Then
                j = InStr(i, sValue, "(", CompareMethod.Text)
                ' esto sólo permite funciones de tres letras
                'sName = Mid$(sValue, i + 3)
                If j = 0 Then j = 3
                sName = Mid$(sValue, j)
                sParams = parametros(sName)
                ' Calcular los parámetros
                sParams = Calcular(sParams)
                '
                esFunVB = True
                '
                ' Convertir el parámetro para usar con estas funciones numéricas
                Select Case sNameFun
                    Case "Int("
                        sParams = Int(sParams)
                    Case "Fix("
                        sParams = Fix(sParams)
                    Case "Abs("
                        sParams = Abs(CDbl(sParams))
                    Case "Sgn("
                        sParams = Sign(CDbl(sParams))
                    Case "Sqr("
                        sParams = Sqrt(CDbl(sParams))
                    Case "Cos("
                        sParams = Cos(sParams)
                    Case "Sin("
                        sParams = Sin(sParams)
                    Case "Tan("
                        sParams = Tan(sParams)
                    Case "Atn("
                        sParams = Atan(sParams)
                    Case "Exp("
                        sParams = Exp(sParams)
                    Case "Log(", "Ln("  ' logaritmo natural (en base e)
                        sParams = Log(sParams)
                    Case "Log10("       ' logaritmo en base 10              (03/Nov/02)
                        sParams = Log(sParams) / Log(10)
                        ' Nuevas funciones                                      (02/Nov/02)
                    Case "Hex(", "Dec2Hex("
                        sParams = Hex(sParams)
                    Case "Oct(", "Dec2Oct("
                        sParams = Oct(sParams)
                    Case "Round("
                        sParams = Round(CDbl(sParams))
                    Case "Bin2Dec("
                        sParams = Bin2Dec(sParams)
                    Case "Hex2Dec("
                        sParams = Val("&H" & sParams)
                    Case "Oct2Dec("
                        sParams = Val("&O" & sParams)
                End Select
                sName = "(" & sParams & ")" & sName
            Else
                ' Algunas otras que se evaluarán aquí
                ' deben estar declaradas con @
                i = MultInStr(sValue, "Max( Min( Bin( ", sNameFun)
                ' El formato será Max(Num1, Num2)
                If i Then
                    On Error Resume Next
                    ' con esto sólo se pueden tener funciones de 3 caracteres
                    'sName = Mid$(sValue, i + 3)
                    j = InStr(i, sValue, "(", CompareMethod.Text)
                    If j = 0 Then j = 3
                    sName = Mid$(sValue, j)
                    '
                    sParams = parametros(sName)
                    i = InStr(sParams, ",")
                    If i Then
                        sValue = Left$(sParams, i - 1)
                        sParams = Mid$(sParams, i + 1)
                        sValue = ParseFormula(sValue)
                        sParams = ParseFormula(sParams)
                        sValue = Calcular(sValue)
                        sParams = Calcular(sParams)
                        If sNameFun = "Max(" Then
                            sName = IIf(CDbl(sValue) > CDbl(sParams), sValue, sParams)
                        ElseIf sNameFun = "Min(" Then
                            sName = IIf(CDbl(sValue) < CDbl(sParams), sValue, sParams)
                        ElseIf sNameFun = "Bin(" Then
                            '
                            esFunVB = True
                            '
                            If Len(sParams) Then
                                sName = Dec2Bin(sValue, sParams)
                            Else
                                sName = Dec2Bin(sValue)
                            End If
                        End If
                        'Pendiente
                        'If Err() Then
                        '    sName = "Error: hay que usarla con @ o alguna variable no está definida"
                        'End If
                        'On Error GoTo 0
                        'Err = 0
                    Else
                        sName = sParams
                    End If
                End If
            End If
            '//////////////////////////////////////////////////////////////////////////
        End If
        FunctionVal = sName
    End Function

    Public Function VariableVal(ByVal sName As String) As String
        '--------------------------------------------------------------------------
        ' Comprueba si es una variable,
        ' de ser así, devolverá el contenido o valor de esa variable
        '
        ' Las variables estarán en un array
        '--------------------------------------------------------------------------
        Dim i As Long
        Dim sValue As String
        '
        sName = Trim$(sName)
        sValue = ""
        If Len(sName) Then
            For i = 1 To UBound(aVariables)
                If aVariables(i).Name = sName Then
                    sValue = aVariables(i).Value
                    Exit For
                End If
            Next
        End If
        VariableVal = sValue
    End Function

    Public Sub NewFunction(ByVal sName As String, ByVal sParams As String, ByVal sFormula As String)
        '--------------------------------------------------------------------------
        ' Asigna una nueva función al array de funciones
        ' Los parámetros serán el nombre, los parámetros y la fórmula a usar
        '
        ' Si la función indicada ya existe, se sustituirán los valores especificados
        '--------------------------------------------------------------------------
        Dim i As Long
        Dim Hallado As Boolean
        Dim NumF As Long
        '
        sName = Trim$(sName)
        sParams = Trim$(sParams)
        sFormula = Trim$(sFormula)
        '
        NumF = UBound(aFunctions)
        Hallado = False
        ' Comprobar si la función ya existe
        For i = LBound(aFunctions) To UBound(aFunctions)
            ' Si es así, asignar el nuevo valor
            If aFunctions(i).Name = sName Then
                ' sólo añadirla si no es de las predefinidas            (03/Nov/02)
                If i > mNumFunctions Then
                    aFunctions(i).Params = sParams
                    aFunctions(i).Formula = sFormula
                End If
                Hallado = True
                Exit For
            End If
        Next
        ' Si no se ha hallado la función, añadirla
        If Not Hallado Then
            Dim a As New tFunctions
            With a
                .Name = sName
                .Params = sParams
                .Formula = sFormula
            End With
            NumF = NumF + 1
            ReDim Preserve aFunctions(NumF)
            aFunctions(NumF) = a
        End If
    End Sub

    Public Sub NewVariable(ByVal sName As String, ByVal sValue As String)
        '--------------------------------------------------------------------------
        ' Asigna una nueva variable al array de variables
        ' Los parámetros serán el nombre y el valor
        '
        ' Si la variable indicada ya existe, se sustituirá el valor por el indicado
        '--------------------------------------------------------------------------
        Dim i As Long
        Dim Hallado As Boolean
        Dim NumVars As Long
        '
        sName = Trim$(sName)
        sValue = Trim$(sValue)
        '
        NumVars = UBound(aVariables)
        Hallado = False
        ' Comprobar si la variable ya existe
        For i = LBound(aVariables) To UBound(aVariables)

            ' Si es así, asignar el nuevo valor
            If aVariables(i).Name = sName Then
                aVariables(i).Value = sValue
                Hallado = True
                Exit For
            End If
        Next
        ' Si no se ha hallado la variable, añadir una nueva
        If Not Hallado Then
            Dim a As New tVariable
            With a
                .Name = sName
                .Value = sValue
            End With
            NumVars = NumVars + 1
            ReDim Preserve aVariables(NumVars)
            aVariables(NumVars) = a
        End If
    End Sub

    Public Function ParseFormula(ByVal sF As String) As String
        '--------------------------------------------------------------------------
        ' Analiza la fórmula indicada, sustituyendo las variables y funciones
        ' por sus valores, después habrá que calcular el resultado devuelto.
        ' En esta función se analizan las variables y funciones, dejando el valor
        ' que devolverían.
        ' Si la variable o función tiene otras variables o funciones se analizan
        ' y se ponen los valores devueltos.
        '--------------------------------------------------------------------------
        Dim qFuncion As Long
        Dim sFormula As String
        Dim sToken As String
        Dim sOp As String
        Dim sVar As String
        Dim sFunFormula As String
        '
        Do
            ' Asignar a sToken el siguiente elemento a procesar
            sOp = ""
            sToken = aToken(sF, sOp)
            ' Si no es una función ni una variable, usar el valor indicado
            If Len(sToken) Then
                ' Si no es una variable o función
                If Not IsFuncOrVar(sToken) Then
                    sFormula = sFormula & sToken & sOp
                Else
                    ' Comprobar si el Token es una variable,
                    ' si es así sustituirla por el valor
                    sVar = VariableVal(sToken)
                    If Len(sVar) Then
                        ' Comprobar si la variable contiene otras variables
                        ' o funciones
                        sVar = ParseFormula(sVar)
                        ' Asigna a sToken el valor obtenido
                        sToken = Calcular(sVar)
                    End If
                    ' Comprobar si el Token es una función
                    qFuncion = IsFunction(sToken)
                    ' Si es una función, qFuncion tiene el índice de la función
                    If qFuncion Then
                        ' Asignar los parámetros que usa la función
                        sVar = aFunctions(qFuncion).Params
                        ' La fórmula a usar para esta función
                        sFunFormula = aFunctions(qFuncion).Formula
                        ' Si admite parámetros
                        If Len(sVar) Then
                            '//////////////////////////////////////////////////////////
                            ' Usar la funcion Parametros para analizar los prámetros
                            '//////////////////////////////////////////////////////////
                            If sOp = "(" Then
                                sF = sOp & sF
                                'sParams = Parametros(sF)
                                'If Len(sParams) Then
                                sOp = ""
                                'End If
                            End If
                            sFunFormula = ConvertirParametros(sFunFormula, sVar, sF)
                        End If
                        ' Si tiene @FuncionInterna
                        ' usar esa función
                        sVar = ""
                        sVar = FunctionVal(sFunFormula)
                        If Len(sVar) Then
                            sFunFormula = ParseFormula(sVar)
                        End If
                        'sFormula = sFormula & Calcular(sFunFormula & sOp & sF)
                        If sOp <> Chr(34) Then
                            sFunFormula = sFunFormula & sOp & sF
                            ' Esto daba problemillas                    (03/Nov/02)
                            ' a ver si así se soluciona...
                            If sFormula = "(" Then
                                ' poner la del final
                                sFunFormula = sFormula & sFunFormula
                                sFormula = ""
                            End If
                            sFunFormula = ParseFormula(sFunFormula)
                            sFormula = sFormula & Calcular(sFunFormula)
                            sF = ""
                        Else
                            sFormula = sFormula & Calcular(sFunFormula)
                            sFormula = sFormula & sOp & sF
                            sF = ""
                        End If
                    Else
                        sFormula = sFormula & sToken & sOp
                    End If
                End If
            Else
                sFormula = sFormula & sToken & sOp
            End If
            '
        Loop While Len(sF)
        ' Devolver la expresión lista para calcular el valor
        ParseFormula = sFormula
    End Function

    Public Function Calcular(ByVal sFormula As String) As String
        '--------------------------------------------------------------------------
        ' Calcula el resultado de la expresión que entra en sFormula    (22/Oct/91)
        ' Modificado por la cuenta de la vieja...                 (01.12  7/May/93)
        ' Revisado para usar con cFormulas                              (06/Feb/99)
        '--------------------------------------------------------------------------
        Dim i As Long, j As Long, k As Long
        Dim j1 As Long, k1 As Long, n As Long
        Dim pn As Long
        Dim n1 As Object  'Double
        Dim n2 As Object  'Double
        Dim n3 As Object  'Double
        Dim Operador As String
        Dim Cifra1 As String
        Dim Cifra2 As String
        Dim strP As String
        Dim sOperadores As String
        ' Estos son los símbolos a buscar para el operador anterior
        ' se deben incluir los paréntesis ya que estos separan precedencias
        Const cOperadores As String = "%^*/\+-()"
        '
        Static sFormulaAnt As String
        '
        ' Quitarle los espacios extras
        sFormula = Trim$(sFormula)
        '
        sOperadores = "% ^ * / \ "
        ' Si la fórmula tiene una operación, calcularla
        If MultInStr(sFormula, sOperadores) Then
            esFunVB = False
        End If
        '
        If esFunVB Then
            ' Si es una función interna                                 (03/Nov/02)
            ' Devolver lo que ya se ha calculado, quitar los paréntesis, etc.
            If Left$(sFormula, 1) = "@" Then sFormula = Trim$(Mid$(sFormula, 2))
            Do
                If Left$(sFormula, 1) = "(" Then
                    sFormula = Mid$(sFormula, 2)
                    If Right$(sFormula, 1) = ")" Then
                        sFormula = Left$(sFormula, Len(sFormula) - 1)
                    End If
                    sFormula = Trim$(sFormula)
                Else
                    Exit Do
                End If
            Loop
            Calcular = sFormula
            Exit Function
        End If
        '
        '//////////////////////////////////////////////////////////////////////////
        ' Para analizar siguiendo las operaciones de más "peso",        (07/Feb/99)
        ' se buscarán operaciones en este orden % ^ * / \ + -
        ' y si se encuentran, se incluirán entre paréntesis para que se procesen
        ' antes que el resto:
        ' 25 + 100 * 3 se convertiría en: 25 + (100 * 3)
        '
        ' Buscar cada uno de los operadores y añadir los paréntesis necesarios
        ' No se incluyen la suma y resta ya que son las que menos peso tienen
        sOperadores = "% ^ * / \ "
        ' Sólo procesar si tiene uno de los operadores
        If MultInStr(sFormula, sOperadores, Operador) Then
            Cifra1 = sFormula
            n = Len(Cifra1)
            For i = 1 To Len(sOperadores) Step 2
                Operador = Mid$(sOperadores, i, 1)
                ' Se debería buscar de atrás para delante
                ' (ya se busca)
                pn = RInStr(n, Cifra1, Operador)
                If pn Then
                    ' Tenemos ese operador
                    ' buscar el signo anterior
                    k = 0
                    For j = pn - 1 To 1 Step -1
                        k = InStr(cOperadores, Mid$(Cifra1, j, 1))
                        If k Then
                            ' Sólo procesar si el signo anterior es diferente de )
                            If Mid$(cOperadores, k, 1) <> ")" Then
                                ' Buscar el signo siguiente
                                k1 = 0
                                For j1 = pn + 1 To Len(Cifra1)
                                    k1 = InStr(cOperadores, Mid$(Cifra1, j1, 1))
                                    If k1 Then
                                        ' Añadirle los paréntesis
                                        ' Si se multiplica por un número negativo
                                        k = MultInStr(Cifra1, "*- /- \- ")
                                        If k Then
                                            Cifra1 = Left$(Cifra1, j) & "(" & Mid$(Cifra1, j + 1, j1 - j - 2) & ")" & Mid$(Cifra1, k)
                                        Else
                                            If Right$(Mid$(Cifra1, j + 1, j1 - j - 1) & ")" & Mid$(Cifra1, j1, 1), 3) = "*)(" Then
                                                Cifra1 = Left$(Cifra1, j) & "(" & Mid$(Cifra1, j + 1, j1 - j - 2) & Mid$(Cifra1, j1 - 1) & ")"
                                            Else
                                                Cifra1 = Left$(Cifra1, j) & "(" & Mid$(Cifra1, j + 1, j1 - j - 1) & ")" & Mid$(Cifra1, j1)
                                            End If
                                        End If
                                        Exit For
                                    End If
                                Next
                                ' Si no hay ningún signo siguiente
                                If k1 = 0 Then
                                    Cifra1 = Left$(Cifra1, j) & "(" & Mid$(Cifra1, j + 1) & ")"
                                End If
                            End If
                            Exit For
                        End If
                    Next
                    pn = RInStr(n, Cifra1, Operador)
                    n = pn - 1
                    i = i - 2
                End If
            Next
            sFormula = Cifra1
        End If
        '
        '//////////////////////////////////////////////////////////////////////////
        '
        ' Buscar paréntesis e ir procesando las expresiones.
        Do While InStr(sFormula, "(")
            pn = InStr(sFormula, ")")
            ' Si hay paréntesis de cierre
            If pn Then
                For i = pn To 1 Step -1
                    If Mid$(sFormula, i, 1) = "(" Then
                        ' Calcular lo que está entre paréntesis
                        strP = Mid$(sFormula, i + 1, pn - i - 1)
                        strP = Calcular(strP)
                        sFormula = Left$(sFormula, i - 1) & strP & Mid$(sFormula, pn + 1)
                        Exit For
                    End If
                Next
            Else
                sFormula = sFormula & ")"
            End If
        Loop

        ' Si la fórmula a procesar tiene algún operador
        sOperadores = "% ^ * / \ + - "
        If MultInStr(sFormula, sOperadores, Operador) Then
            '//////////////////////////////////////////////////////////////////////
            ' Si hay más de un operador,                                (11/Ene/01)
            ' ponerlos dentro de paréntesis según el nivel de precedencia
            ' He añadido el + y - ya que no hacía los cálculos bien     (22/Ene/01)
            '//////////////////////////////////////////////////////////////////////
            If MultipleStr2InStr1(sFormula, "%^*/\+-") Then
                '
                ' A ver si esto arregla los cálculos "normales"         (22/Feb/01)
                ' ya que daba error al calcular: 3*2+5
                ' Gracias a Luis Americo Popiti
                '
                If Len(sFormulaAnt) = 0 Then
                    sFormulaAnt = sFormula
                End If
                If sFormulaAnt <> sFormula Then
                    sFormula = Calcular(sFormula)
                End If
                sFormulaAnt = ""
            End If
            Operador = ""
            Cifra1 = ""
            Cifra2 = ""
            Do
                ' Buscar la primera cifra
                If Len(sFormula) Then
                    If Cifra1 = "" Then
                        buscarCifra(sFormula, Cifra1)
                    End If
                    Operador = Left$(sFormula, 1)
                    sFormula = Mid$(sFormula, 2)
                    ' Buscar la segunda cifra
                    buscarCifra(sFormula, Cifra2)
                    '
                    n1 = 0
                    If Len(Cifra1) Then
                        n1 = CDec(Cifra1)
                    End If
                    ' Esto es necesario por si no se ponen los paréntesis de apertura
                    n2 = 0
                    If Len(Cifra2) Then
                        n2 = CDec(Cifra2)
                    End If
                    ' Efectuar el cálculo
                    Select Case Operador
                        Case "+"
                            n3 = n1 + n2
                        Case "-"
                            n3 = n1 - n2
                        Case "*"
                            n3 = n1 * n2
                            ' Si se divide por cero, se devuelve cero en lugar de dar error
                        Case "/"
                            If n2 <> 0.0# Then
                                n3 = n1 / n2
                            Else
                                n3 = 0.0#
                            End If
                        Case "\"
                            If n2 <> 0.0# Then
                                n3 = n1 \ n2
                            Else
                                n3 = 0.0#
                            End If
                        Case "^"
                            n3 = n1 ^ n2
                            ' Cálculo de porcentajes:
                            ' 100 % 25 = 25 (100 * (25 / 100))
                        Case "%"
                            n3 = n1 * CDec(n2 / CDec(100))
                            ' Si es comillas dobles, no evaluar
                        Case Chr(34)
                            ' Calcular el resto después de las comillas
                            i = InStr(sFormula, Chr(34))
                            If i Then
                                Cifra1 = Mid$(sFormula, i + 1)
                                sFormula = Operador & Left$(sFormula, i)
                                Operador = ""
                                sFormula = sFormula & Calcular(Cifra1)
                                Calcular = sFormula
                                Exit Function
                            Else
                                sFormula = Operador & sFormula
                                Operador = ""
                                Calcular = sFormula
                                Exit Function
                            End If
                            ' Si no es una operación reconocida, devolver la suma,
                            ' ya que esto puede ocurrir con los valores asignados a variables
                        Case Else
                            ' Por si se incluye una palabra que no está declarada
                            ' (variable o función)
                            If Len(Cifra1 & Cifra2) Then
                                If Len(Cifra1) = 0 Then
                                    Cifra1 = "0"
                                End If
                                If Len(Cifra2) = 0 Then
                                    Cifra2 = "0"
                                End If
                                n3 = CDec(Cifra1) + CDec(Cifra2)
                            Else
                                n3 = 0
                            End If
                    End Select
                    Cifra1 = CStr(n3)
                Else
                    Exit Do
                End If
            Loop While Operador <> ""
            Calcular = CStr(n3)
        Else
            ' Si no tiene ningún operador, devolver la fórmula
            ' Habría que quitarle los caracteres extraños               (10/Feb/99)
            If Left$(sFormula, 1) <> Chr(34) Then
                ' tener en cuenta los números hexadecimales             (03/Nov/02)
                sOperadores = "0123456789,.ABCDEF"
                Cifra1 = ""
                For i = 1 To Len(sFormula)
                    If InStr(sOperadores, Mid$(sFormula, i, 1)) Then
                        Cifra1 = Cifra1 & Mid$(sFormula, i, 1)
                    End If
                Next
                sFormula = Cifra1
            End If
            Calcular = sFormula
        End If
    End Function

    Public Function RInStr(ByVal v1 As Object, ByVal v2 As Object, _
                           Optional ByVal v3 As Object = Nothing) As Long
        '--------------------------------------------------------------------------
        ' Devuelve la posición de v2 en v1, empezando por atrás
        '--------------------------------------------------------------------------
        Dim i As Long
        Dim posIni As Long
        Dim sTmp As String
        Dim s1 As String
        Dim s2 As String
        '
        If Length(v3) Then
            ' Si no se especifican los tres parámetros
            s1 = CStr(v1)       ' La primera cadena
            s2 = CStr(v2)       ' la segunda cadena
            posIni = Len(s1)    ' el último caracter de la cadena
        Else
            posIni = CLng(v1)   ' la posición por la que empezar
            s1 = CStr(v2)       ' la primera cadena (segundo parámetro)
            s2 = CStr(v3)       ' la segunda cadena (tercer parámetro)
        End If
        ' Valor inicial de la búsqueda, si no se encuentra, es cero
        RInStr = 0
        ' Siempre se empieza a buscar por el final
        For i = posIni - Len(s2) + 1 To 1 Step -1
            ' Tomar el número de caracteres que tenga la segunda cadena
            sTmp = Mid$(s1, i, Len(s2))     ' Si son iguales...
            If sTmp = s2 Then               ' esa es la posición
                RInStr = i
                Exit For
            End If
        Next
    End Function

    Private Sub buscarCifra(ByRef Expresion As String, ByRef Cifra As String)
        '--------------------------------------------------------------------------
        ' Buscar en Expresion una cifra                             ( 5 / 10/May/93)
        ' Devuelve la cifra y el resto de la expresión
        '--------------------------------------------------------------------------
        Const OPERADORES As String = "+-*/\^%"
        Const CIFRAS As String = "0123456789., "
        Const POSITIVO As Long = 1&
        Const NEGATIVO As Long = -1&
        '
        Dim Signo As Long
        Dim ultima As Long
        Dim i As Long
        Dim s As String
        Dim sCifras As String
        Dim sSigno As String
        '
        ' Quitar los espacios del principio
        Expresion = LTrim$(Expresion)
        '
        ' Capturar errores por si se usan varios parámetros
        On Error Resume Next
        '
        ' Evaluar sólo si no está entre comillas
        If Left$(Expresion, 1) <> Chr(34) Then
            Signo = POSITIVO                    'Comprobar si es un número negativo
            If Left$(Expresion, 1) = "-" Then
                Signo = NEGATIVO
                Expresion = Mid$(Expresion, 2)
            End If
            '
            ultima = 0
            s = ""
            For i = 1 To Len(Expresion)
                If InStr(CIFRAS, Mid$(Expresion, i, 1)) Then
                    s = s & Mid$(Expresion, i, 1)
                    ultima = i
                Else
                    Exit For
                End If
            Next i
            ' El val funciona sólo si el decimal es el punto,
            ' cuando es una coma toma sólo la parte entera
            If Len(s) Then
                ' Convertir adecuadamente los decimales
                s = ConvDecimal(s)
                Cifra = CStr((s) * Signo)
            Else
                Cifra = ""
            End If
            Expresion = LTrim$(Mid$(Expresion, ultima + 1))
            If Left$(Expresion, 1) = "E" Then
                ultima = Val(Mid$(Expresion, 3))
                sSigno = Mid$(Expresion, 2, 1)
                s = ""
                For i = 1 To ultima
                    s = s & "0"
                Next
                s = "1" & s
                If sSigno = "-" Then
                    'Cifra = CCur((Cifra) / (s))
                    Cifra = (Cifra) / (s)
                Else
                    'Cifra = CCur((Cifra) * (s))
                    Cifra = (Cifra) * (s)
                End If
                Expresion = Mid$(Expresion, 5)
            End If
        End If
        On Error GoTo 0
        'Err = 0
    End Sub

    Public Function ConvDecimal(ByVal strNum As String, _
                                Optional ByRef sDecimal As String = ",", _
                                Optional ByRef sDecimalNo As String = ".") As String
        '--------------------------------------------------------------------------
        ' Asigna el signo decimal adecuado (o lo intenta)                   (10/Ene/99)
        ' Devuelve una cadena con el signo decimal del sistema
        '--------------------------------------------------------------------------
        Dim i As Long, j As Long
        Dim sNumero As String
        '
        ' Averiguar el signo decimal
        sNumero = Format$(25.5, "#.#")
        If InStr(sNumero, ".") Then
            sDecimal = "."
            sDecimalNo = ","
        Else
            sDecimal = ","
            sDecimalNo = "."
        End If
        '
        strNum = Trim$(strNum)
        If Left$(strNum, 1) = sDecimalNo Then
            Mid$(strNum, 1, 1) = sDecimal
        End If
        '
        ' Si el número introducido contiene signos no decimales
        j = 0
        i = 1
        Do
            i = InStr(i, strNum, sDecimalNo, CompareMethod.Text)
            If i Then
                j = j + 1
                i = i + 1
            End If
        Loop While i
        '
        If j = 1 Then
            ' Cambiar ese símbolo por un espacio, si sólo hay uno de esos signos
            i = InStr(strNum, sDecimalNo)
            If i Then
                If InStr(strNum, sDecimal) Then
                    Mid$(strNum, i, 1) = " "
                Else
                    Mid$(strNum, i, 1) = sDecimal
                End If
            End If
        Else
            ' En caso de que tenga más de uno de estos símbolos
            ' convertirlos de manera adecuada.
            ' Por ejemplo:
            ' si el signo decimal es la coma:
            '   1,250.45 sería 1.250,45 y quedaría en 1250,45
            ' si el signo decimal es el punto:
            '   1.250,45 sería 1,250.45 y quedaría en 1250.45
            '
            ' Aunque no se arreglará un número erróneo:
            ' si el signo decimal es la coma:
            '   1,250,45 será lo mismo que 1,25
            '   12,500.25 será lo mismo que 12,50
            ' si el signo decimal es el punto:
            '   1.250.45 será lo mismo que 1.25
            '   12.500,25 será lo mismo que 12.50
            '
            i = 1
            Do
                i = InStr(i, strNum, sDecimalNo, CompareMethod.Text)
                If i Then
                    j = j - 1
                    If j = 0 Then
                        Mid$(strNum, i, 1) = sDecimal
                    Else
                        Mid$(strNum, i, 1) = " "
                    End If
                    i = i + 1
                End If
            Loop While i
        End If
        '
        j = 0
        ' Quitar los espacios que haya por medio
        Do
            i = InStr(strNum, " ")
            If i = 0 Then Exit Do
            strNum = Left$(strNum, i - 1) & Mid$(strNum, i + 1)
        Loop
        '
        ConvDecimal = strNum
    End Function



    Public Sub ShowFunctions(ByRef aList As Object)
        '--------------------------------------------------------------------------
        ' Devuelve las funciones y las fórmulas usadas en el formato:
        '   Nombre = Función | Parámentros
        ' El parámetro indicará una colección o un ListBox/ComboBox
        '--------------------------------------------------------------------------
        Dim i As Long
        '
        For i = 1 To UBound(aFunctions)
            With aFunctions(i)
                If TypeOf aList Is Collection Then
                    aList.Add(.Name & " = " & .Formula & " | " & .Params)
                Else
                    aList.AddItem(.Name & " = " & .Formula & " | " & .Params)
                End If
            End With
        Next
    End Sub

    Public Sub ShowVariables(ByVal aList As Object)
        '--------------------------------------------------------------------------
        ' Devuelve las variables y los valores en el formato:
        '   Nombre = Valor
        ' El parámetro indicará una colección o un ListBox/ComboBox
        '--------------------------------------------------------------------------
        Dim i As Long
        '
        For i = 1 To UBound(aVariables)
            If TypeOf aList Is Collection Then
                aList.Add(aVariables(i).Name & " = " & aVariables(i).Value)
            Else
                aList.AddItem(aVariables(i).Name & " = " & aVariables(i).Value)
            End If
        Next
    End Sub

    Public Function Formula(ByVal sExpresion As String) As String
        '--------------------------------------------------------------------------
        ' Esta función calcula directamente la expresión
        '--------------------------------------------------------------------------
        Dim tmpCientifica As Boolean
        Dim s As String
        '
        tmpCientifica = m_NotacionCientifica
        '
        ' Comprobar si hay asignaciones en la expresión
        sExpresion = AsignarVariables(sExpresion, Chr(34))
        ' Si se usa Bin, Bin2Dec o Dec2Bin no usar notación cientifica  (02/Nov/02)
        If MultInStr(sExpresion, "Bin( Bin2Dec Dec2Bin Hex( Oct( ") Then
            m_NotacionCientifica = False
        End If
        ' Interpretar la expresión
        sExpresion = ParseFormula(sExpresion)
        '
        '
        ' Calcular la expresión
        s = Calcular(sExpresion)
        '
        ' Convertir el resultado en Double                              (29/Ene/01)
        ' Si así se ha especificado en la propiedad NotacionCientifica,
        ' que por defecto es True
        '
        ' Si da error, usar el valor devuelto por Calcular
        On Error Resume Next
        '
        If m_NotacionCientifica Then
            Formula = CDbl(s)
            'Pendiente
            'If Err() Then
            '    Formula = s
            'End If
        Else
            Formula = s
        End If
        '
        m_NotacionCientifica = tmpCientifica
        'Pendiente
        '  Err = 0
    End Function

    Public Function IsFuncOrVar(ByVal sName As String) As Boolean
        '--------------------------------------------------------------------------
        ' Comprobar si es una función o variable

        ' Es importante comprobar primero las funciones
        ' para que no se añada una función como si fuese una variable no declarada
        '--------------------------------------------------------------------------
        ' Si no es un número
        If Not IsNumeric(sName) Then
            If IsFunction(sName) Then
                IsFuncOrVar = True
            ElseIf IsVariable(sName) Then
                IsFuncOrVar = True
            End If
        End If
    End Function

    Private Function parametros(ByRef sExp As String) As String
        '--------------------------------------------------------------------------
        ' Devolverá los parámetros de la expresión pasada por referencia(08/Feb/99)
        ' Los parámetros deben estar encerrados entre paréntesis
        ' En sExp, se devolverá el resto de la cadena.
        '--------------------------------------------------------------------------
        Dim i As Long, j As Long, k As Long
        Dim sParams As String
        Dim sExpAnt As String
        '
        sExp = Trim$(sExp)
        sExpAnt = sExp
        '
        '
        ' Buscarlos, estarán entre paréntesis
        '
        If Left$(sExp, 1) = "(" Then
            sExp = Mid$(sExp, 2)
            ' Buscar el siguiente )
            k = 0
            j = 0
            For i = 1 To Len(sExp)
                If Mid$(sExp, i, 1) = "(" Then
                    j = j + 1
                End If
                If Mid$(sExp, i, 1) = ")" Then
                    j = j - 1
                    If j = -1 Then
                        k = i
                        Exit For
                    End If
                End If
            Next
            If k Then
                sParams = Left$(sExp, k - 1)
                sExp = Mid$(sExp, k + 1)
            End If
        Else
            sParams = ""
            sExp = sExpAnt
        End If
        '
        parametros = sParams
    End Function

    Public Function FunctionParams(ByVal sName As String) As String
        ' Devuelve los parámetros de la función indicada                (12/Feb/99)
        Dim i As Long
        '
        sName = Trim$(sName)
        FunctionParams = ""
        If Len(sName) Then
            For i = 1 To UBound(aFunctions)
                If aFunctions(i).Name = sName Then
                    FunctionParams = aFunctions(i).Params
                    Exit For
                End If
            Next
        End If
    End Function

    Public Function FunctionFormula(ByVal sName As String) As String
        ' Devuelve la fórmula de la función indicada                    (10/Feb/99)
        Dim i As Long
        '
        sName = Trim$(sName)
        FunctionFormula = ""
        If Len(sName) Then
            For i = 1 To UBound(aFunctions)
                If aFunctions(i).Name = sName Then
                    FunctionFormula = aFunctions(i).Formula
                    Exit For
                End If
            Next
        End If
    End Function

    Public Function ConvertirParametros(ByVal sFunFormula As String, _
                                        ByVal sVar As String, _
                                        ByRef sF As String) As String
        '--------------------------------------------------------------------------
        ' Sustituir parámetros                                          (12/Feb/99)
        ' Sustituye en sFunFormula los parámetros indicados en sVar que están
        ' en la expresión sF.
        ' Devuelve el valor procesado.
        ' sF debe pasarse por referencia, ya que se devovlerá lo que quede después
        ' de procesarse los parámetros
        '--------------------------------------------------------------------------
        Dim i As Long, j As Long, k As Long, n As Long
        Dim sParams As String
        Dim sParamF As String       ' Parámetro en la fórmula
        Dim sParamE As String       ' Parámetro en la expresión
        Dim sParamX As String       ' Para añadir parámetros a la fórmula
        '
        ConvertirParametros = ""
        '//////////////////////////////////////////////////////////
        ' Usar la funcion Parametros para analizar los prámetros
        '//////////////////////////////////////////////////////////
        sParams = parametros(sF)
        '//////////////////////////////////////////////////////////
        If Len(sParams) Then
            ' Comprobar si los parámetros contienen alguna variable
            ' u otra función
            sParams = ParseFormula(sParams)
            '
            ' Sustituir los parámetros por los indicados en la fórmula
            ' (en principio sólo se admite uno)
            ' Sustituir en la fórmula el nombre de la variable
            ' por el parámetro
            '
            ' Si sólo tiene un parámetro
            If InStr(sVar, ",") = 0 Then
                ' comprobar si sParams tiene más de uno
                i = InStr(sParams, ",")
                If i Then
                    ' De ser así, quedarse sólo con el primero
                    sParams = Trim$(Left$(sParams, i - 1))
                    ' Puede que los parámetros estuviesen ente paréntesis
                    If Left$(sParams, 1) = "(" Then
                        ' Si le falta el del final, añadirselo
                        If Right$(sParams, 1) <> ")" Then
                            sParams = sParams & ")"
                        End If
                    End If
                End If
                Do
                    ' Si sVar es una cadena vacía,                      (29/Oct/02)
                    ' esta comprobación dará un nuevo positivo
                    i = InStr(sFunFormula, sVar)
                    If Len(sVar) = 0 Then i = 0
                    If i Then
                        ' Poner los parámetros dentro de paréntesis
                        sFunFormula = Left$(sFunFormula, i - 1) & "(" & sParams & ")" & Mid$(sFunFormula, i + Len(sVar))
                    End If
                    ' Por si se queda colgado convirtiendo parámetros...
                Loop While i > 0 And Len(sFunFormula) < 3072&
            Else
                ' Resolver los parámetros
                ' sParams tiene los parámetros a evaluar
                ' sVar tiene los nombres de los parámetros
                sVar = sVar & ","
                sParams = sParams & ","
                If InStr(sFunFormula, "...") Then
                    ' Contar el número de parámetros que se han pasado
                    ' para el caso de parámetros opcionales (se usan ...)
                    i = 0
                    For j = 1 To Len(sParams)
                        If Mid$(sParams, j, 1) = "," Then i = i + 1
                    Next
                    ' Para convertir los parámetros opcionales
                    ' en variables que después se puedan sustituir.
                    ' Las variables deben ser diferentes.
                    sParamX = "NumX"
                    n = 0
                    ' Obtener el último parámetro de la fórmula
                    sParamF = Right$(sFunFormula, 4)
                    sFunFormula = Left$(sFunFormula, Len(sFunFormula) - 4)
                    sVar = Left$(sVar, Len(sVar) - 5)
                    Do
                        k = 0
                        For j = 1 To Len(sVar)
                            If Mid$(sVar, j, 1) = "," Then k = k + 1
                        Next
                        If i > k + 1 Then
                            ' Buscar el último de sVar
                            For j = Len(sVar) - 1 To 1 Step -1
                                If Mid$(sVar, j, 1) = "," Then
                                    ' De esta forma aceptará hasta 100 parámetros   (14/Feb/99)
                                    sVar = sVar & "," & sParamX & Format$(n, "00") ' CStr(n)
                                    sFunFormula = sFunFormula & Left$(sParamF, 1) & sParamX & Format$(n, "00") 'CStr(n)
                                    n = n + 1
                                    Exit For
                                End If
                            Next
                        End If
                    Loop While i > k + 1
                    If Right$(sVar, 1) <> "," Then sVar = sVar & ","
                End If
                '
                Do
                    j = InStr(sVar, ",")
                    If j Then
                        sParamF = Trim$(Left$(sVar, j - 1))
                        sVar = Trim$(Mid$(sVar, j + 1))
                        i = InStr(sParams, ",")
                        If i Then
                            sParamE = Trim$(Left$(sParams, i - 1))
                            sParams = Trim$(Mid$(sParams, i + 1))
                        Else
                            sParamE = sParams
                            sParams = ""
                        End If
                        ' Reemplazar sParamF por el parámetro
                        Do
                            ' Si sParamF es una cadena vacía,           (29/Oct/02)
                            ' esta comprobación dará un nuevo positivo
                            i = InStr(sFunFormula, sParamF)
                            If Len(sParamF) = 0 Then i = 0
                            If i Then
                                ' Poner los parámetros dentro de paréntesis
                                sFunFormula = Left$(sFunFormula, i - 1) & "(" & sParamE & ")" & Mid$(sFunFormula, i + Len(sParamF))
                            End If
                        Loop While i
                    End If
                Loop While j
            End If
            ConvertirParametros = sFunFormula
        End If
    End Function

    Public Function MultipleStr2InStr1(ByVal Str1 As String, _
                                       ByVal Str2 As String) As Boolean
        '--------------------------------------------------------------------------
        ' Devuelve True si:                                             (11/Ene/01)
        '   Str1 tiene más de un caracter de los indicados en Str2
        '--------------------------------------------------------------------------
        Dim i As Long
        Dim n As Long
        '
        ' Buscar cada uno de los caracteres de Str2 en Str1
        n = 0
        For i = 1 To Len(Str2)
            ' Comprobar si está
            If InStr(Str1, Mid$(Str2, i, 1)) Then
                ' si es así, incrementar el contador
                n = n + 1
                ' si ya se han encontrado más de uno, no seguir buscando
                If n > 1 Then Exit For
            End If
        Next
        MultipleStr2InStr1 = (n > 1)
    End Function
    Property NotacionCientifica() As Boolean
        Get
            Return m_NotacionCientifica
        End Get
        Set(ByVal Value As Boolean)
            m_NotacionCientifica = Value
        End Set
    End Property


    Public Function IsFunVB(ByVal sName As String) As Boolean
        ' Comprueba si la función indicada es una función de VB         (02/Nov/02)
        ' (las usadas en FunVBNum)
        Dim i As Long
        Dim sValue As String
        '
        i = InStr(sName, "@")
        If i Then
            sValue = Left$(sName, i - 1) & Mid$(sName, i + 1)
        Else
            sValue = sName
        End If
        '//////////////////////////////////////////////////////////////////////////
        ' Si es alguna de las definidas en la constante FunVBNum
        i = MultInStr(sValue, FunVBNum)
        '
        IsFunVB = CBool(i)
    End Function

    Public Function Dec2Bin(ByVal n As Long, _
                             Optional ByVal nCifras As Long = 16) As String
        ' Convertir el número indicado a binario
        Dim i As Long
        Dim s As String
        '
        On Error GoTo Err2Bin
        s = ""
        For i = nCifras - 1 To 0 Step -1
            If n And (2 ^ i) Then
                s = s & "1"
            Else
                s = s & "0"
            End If
        Next
        Dec2Bin = s
        Exit Function
Err2Bin:
        Dec2Bin = "Error: " & Err.Description & ", al convertir 2^" & CStr(i)
    End Function

    Public Function Bin2Dec(ByVal sDec As String) As Long
        ' Convierte un número binario en decimal
        ' El parámetro debería ser un número con sólo 1 y ceros,
        ' pero se considerará como cero, cualquier carácter que no sea un uno,
        ' excepto los espacios que no se tendrán en cuenta.
        Dim n As Long
        Dim i As Long, j As Long
        Dim C As String
        '
        i = 0
        For j = Len(sDec) To 1 Step -1
            C = Mid$(sDec, j, 1)
            If C = "1" Then
                n = n + 2 ^ i
                i = i + 1
            ElseIf C <> " " Then
                i = i + 1
            End If
        Next
        Bin2Dec = n
    End Function


End Class
