
'Copyright 2018 Kelvys B. Pantaleão

'This file is part of LeaCrypto

'LeaCrypto Is free software: you can redistribute it And/Or modify
'it under the terms Of the GNU General Public License As published by
'the Free Software Foundation, either version 3 Of the License, Or
'(at your option) any later version.

'This program Is distributed In the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty Of
'MERCHANTABILITY Or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License For more details.

'You should have received a copy Of the GNU General Public License
'along with this program.  If Not, see <http://www.gnu.org/licenses/>.


'Este arquivo é parte Do programa LeaCrypto

'LeaCrypto é um software livre; você pode redistribuí-lo e/ou 
'modificá-lo dentro dos termos da Licença Pública Geral GNU como 
'publicada pela Fundação Do Software Livre (FSF); na versão 3 da 
'Licença, ou(a seu critério) qualquer versão posterior.

'Este programa é distribuído na esperança de que possa ser  útil, 
'mas SEM NENHUMA GARANTIA; sem uma garantia implícita de ADEQUAÇÃO
'a qualquer MERCADO ou APLICAÇÃO EM PARTICULAR. Veja a
'Licença Pública Geral GNU para maiores detalhes.

'Você deve ter recebido uma cópia da Licença Pública Geral GNU junto
'com este programa, Se não, veja <http://www.gnu.org/licenses/>.

'GitHub: https://github.com/Kelvysb/LeaCrypto

Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Public Class LeaCryptoEngine

#Region "Constructors"
    Private Sub New()
    End Sub
#End Region

#Region "Functions and SubRoutines"

#Region "Public functions"

    'Encrypt
    '-----------------------------------------------------------------------------------------------
    'From Boolean
    Public Shared Function fnEncryptToBoolean(ByVal p_strKey As String, ByVal p_blnData As List(Of Boolean)) As List(Of Boolean)
        Dim blnReturn As List(Of Boolean)
        Dim objKeys As List(Of clsPairKey)

        Try

            blnReturn = New List(Of Boolean)

            objKeys = fnGenerateKeys(p_strKey, p_blnData.Count)

            blnReturn = fnEncrypt(objKeys, p_blnData)

            Return blnReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnEncryptToHexString(ByVal p_strKey As String, ByVal p_blnData As List(Of Boolean), ByVal p_objEncoding As Encoding) As String

        Dim strReturn As String
        Dim bytReturn As Byte()

        Try

            strReturn = ""

            bytReturn = fnEncryptToByte(p_strKey, p_blnData)

            strReturn = BitConverter.ToString(bytReturn).Replace("-", "")

            Return strReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnEncryptToByte(ByVal p_strKey As String, ByVal p_blnData As List(Of Boolean)) As Byte()
        Try
            Return ToByteArray(fnEncryptToBoolean(p_strKey, p_blnData).ToArray)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Sub sbEncryptToFile(ByVal p_strKey As String, ByVal p_blnData As List(Of Boolean), ByVal p_strOutputFile As String, ByVal p_objEncoding As Encoding)

        Dim bytReturn As Byte()
        Dim objFileOutput As StreamWriter

        Try

            objFileOutput = New StreamWriter(p_strOutputFile, False, p_objEncoding)

            bytReturn = fnEncryptToByte(p_strKey, p_blnData)

            objFileOutput.Write(bytReturn)

            objFileOutput.Close()
            objFileOutput.Dispose()
            objFileOutput = Nothing

        Catch ex As Exception
            Throw
        End Try

    End Sub

    Public Shared Sub sbEncryptToFile(ByVal p_strKey As String, ByVal p_blnData As List(Of Boolean), ByVal p_strOutputFile As String)
        Try
            Call sbEncryptToFile(p_strKey, p_blnData, p_strOutputFile, Encoding.ASCII)
        Catch ex As Exception
            Throw
        End Try
    End Sub


    'From Byte
    Public Shared Function fnEncryptToBoolean(ByVal p_strKey As String, ByVal p_blnData As Byte()) As List(Of Boolean)
        Try
            Return fnEncryptToBoolean(p_strKey, FromByteArray(p_blnData).ToList)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnEncryptToHexString(ByVal p_strKey As String, ByVal p_blnData As Byte(), ByVal p_objEncoding As Encoding) As String
        Try
            Return fnEncryptToHexString(p_strKey, FromByteArray(p_blnData).ToList, p_objEncoding)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnEncryptToByte(ByVal p_strKey As String, ByVal p_blnData As Byte()) As Byte()
        Try
            Return ToByteArray(fnEncryptToBoolean(p_strKey, p_blnData).ToArray)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Sub sbEncryptToFile(ByVal p_strKey As String, ByVal p_blnData As Byte(), ByVal p_strOutputFile As String, ByVal p_objEncoding As Encoding)
        Try
            Call sbEncryptToFile(p_strKey, FromByteArray(p_blnData).ToList, p_strOutputFile, p_objEncoding)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Shared Sub sbEncryptToFile(ByVal p_strKey As String, ByVal p_blnData As Byte(), ByVal p_strOutputFile As String)
        Try
            Call sbEncryptToFile(p_strKey, p_blnData, p_strOutputFile, Encoding.ASCII)
        Catch ex As Exception
            Throw
        End Try
    End Sub


    'from String
    Public Shared Function fnEncryptToBoolean(ByVal p_strKey As String, ByVal p_strData As String, ByVal p_objEncoding As Encoding) As List(Of Boolean)
        Try
            Return fnEncryptToBoolean(p_strKey, FromByteArray(p_objEncoding.GetBytes(p_strData)).ToList)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnEncryptToHexString(ByVal p_strKey As String, ByVal p_strData As String, ByVal p_objEncoding As Encoding) As String
        Try
            Return fnEncryptToHexString(p_strKey, FromByteArray(p_objEncoding.GetBytes(p_strData)).ToList, p_objEncoding)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnEncryptToHexString(ByVal p_strKey As String, ByVal p_strData As String) As String
        Try
            Return fnEncryptToHexString(p_strKey, p_strData, Encoding.ASCII)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnEncryptToByte(ByVal p_strKey As String, ByVal p_strData As String, ByVal p_objEncoding As Encoding) As Byte()
        Try
            Return ToByteArray(fnEncryptToBoolean(p_strKey, p_strData, p_objEncoding).ToArray)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnEncryptToByte(ByVal p_strKey As String, ByVal p_strData As String) As Byte()
        Try
            Return fnEncryptToByte(p_strKey, p_strData, Encoding.ASCII)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Sub sbEncryptToFile(ByVal p_strKey As String, ByVal p_strData As String, ByVal p_strOutputFile As String, ByVal p_objEncoding As Encoding)
        Try
            Call sbEncryptToFile(p_strKey, FromByteArray(p_objEncoding.GetBytes(p_strData)).ToList, p_strOutputFile, p_objEncoding)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Shared Sub sbEncryptToFile(ByVal p_strKey As String, ByVal p_strData As String, ByVal p_strOutputFile As String)
        Try
            Call sbEncryptToFile(p_strKey, p_strData, p_strOutputFile, Encoding.ASCII)
        Catch ex As Exception
            Throw
        End Try
    End Sub


    'From File
    Public Shared Function fnEncryptFileToBoolean(ByVal p_strKey As String, ByVal p_strInputFile As String) As List(Of Boolean)

        Dim blnReturn As List(Of Boolean)
        Dim objFile As StreamReader
        Dim strFile As String

        Try

            blnReturn = New List(Of Boolean)

            objFile = New StreamReader(p_strInputFile)

            strFile = objFile.ReadToEnd()

            blnReturn = fnEncryptToBoolean(p_strKey, strFile, objFile.CurrentEncoding)

            objFile.Close()
            objFile.Dispose()
            objFile = Nothing

            Return blnReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnEncryptFileToHexString(ByVal p_strKey As String, ByVal p_strInputFile As String) As String

        Dim strReturn As String
        Dim bytReturn As Byte()

        Try

            strReturn = ""

            bytReturn = fnEncryptFileToByte(p_strKey, p_strInputFile)

            strReturn = BitConverter.ToString(bytReturn).Replace("-", "")

            Return strReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnEncryptFileToByte(ByVal p_strKey As String, ByVal p_strInputFile As String) As Byte()
        Try
            Return ToByteArray(fnEncryptFileToBoolean(p_strKey, p_strInputFile).ToArray)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Sub sbEncryptFileToFile(ByVal p_strKey As String, ByVal p_strInputFile As String, ByVal p_strOutputFile As String)

        Dim bytReturn As Byte()
        Dim objFileInput As StreamReader
        Dim objFileOutput As StreamWriter
        Dim strFile As String

        Try

            objFileInput = New StreamReader(p_strInputFile)
            objFileOutput = New StreamWriter(p_strOutputFile, False, objFileInput.CurrentEncoding)

            strFile = objFileInput.ReadToEnd()

            bytReturn = fnEncryptToByte(p_strKey, strFile, objFileInput.CurrentEncoding)

            objFileOutput.Write(bytReturn)

            objFileInput.Close()
            objFileInput.Dispose()
            objFileInput = Nothing

            objFileOutput.Close()
            objFileOutput.Dispose()
            objFileOutput = Nothing

        Catch ex As Exception
            Throw
        End Try
    End Sub
    '-----------------------------------------------------------------------------------------------


    'Decrypt
    '-----------------------------------------------------------------------------------------------
    'From Boolean
    Public Shared Function fnDecryptToBoolean(ByVal p_strKey As String, ByVal p_blnData As List(Of Boolean)) As List(Of Boolean)
        Dim blnReturn As List(Of Boolean)
        Dim objKeys As List(Of clsPairKey)

        Try

            blnReturn = New List(Of Boolean)

            objKeys = fnGenerateKeys(p_strKey, p_blnData.Count)

            blnReturn = fnDecrypt(objKeys, p_blnData)

            Return blnReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnDecryptToString(ByVal p_strKey As String, ByVal p_blnData As List(Of Boolean), p_objEncoding As Encoding) As String

        Dim strReturn As String
        Dim bytReturn As Byte()

        Try

            strReturn = ""

            bytReturn = fnDecryptToByte(p_strKey, p_blnData)

            strReturn = p_objEncoding.GetChars(bytReturn)

            Return strReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnDecryptToByte(ByVal p_strKey As String, ByVal p_blnData As List(Of Boolean)) As Byte()
        Try
            Return ToByteArray(fnDecryptToBoolean(p_strKey, p_blnData).ToArray)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Sub sbDecryptToFile(ByVal p_strKey As String, ByVal p_blnData As List(Of Boolean), ByVal p_strOutputFile As String, ByVal p_objEncoding As Encoding)

        Dim bytReturn As Byte()
        Dim objFileOutput As StreamWriter

        Try

            objFileOutput = New StreamWriter(p_strOutputFile, False, p_objEncoding)

            bytReturn = fnDecryptToByte(p_strKey, p_blnData)

            objFileOutput.Write(bytReturn)

            objFileOutput.Close()
            objFileOutput.Dispose()
            objFileOutput = Nothing

        Catch ex As Exception
            Throw
        End Try

    End Sub

    Public Shared Sub sbDecryptToFile(ByVal p_strKey As String, ByVal p_blnData As List(Of Boolean), ByVal p_strOutputFile As String)
        Try
            Call sbDecryptToFile(p_strKey, p_blnData, p_strOutputFile, Encoding.ASCII)
        Catch ex As Exception
            Throw
        End Try
    End Sub


    'From Byte
    Public Shared Function fnDecryptToBoolean(ByVal p_strKey As String, ByVal p_blnData As Byte()) As List(Of Boolean)
        Try
            Return fnDecryptToBoolean(p_strKey, FromByteArray(p_blnData).ToList)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnDecryptToString(ByVal p_strKey As String, ByVal p_blnData As Byte(), p_objEncoding As Encoding) As String
        Try
            Return fnDecryptToString(p_strKey, FromByteArray(p_blnData).ToList, p_objEncoding)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnDecryptToString(ByVal p_strKey As String, ByVal p_blnData As Byte()) As String
        Try
            Return fnDecryptToString(p_strKey, FromByteArray(p_blnData).ToList, Encoding.ASCII)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnDecryptToByte(ByVal p_strKey As String, ByVal p_blnData As Byte()) As Byte()
        Try
            Return ToByteArray(fnDecryptToBoolean(p_strKey, p_blnData).ToArray)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Sub sbDecryptToFile(ByVal p_strKey As String, ByVal p_blnData As Byte(), ByVal p_strOutputFile As String, ByVal p_objEncoding As Encoding)
        Try
            Call sbDecryptToFile(p_strKey, FromByteArray(p_blnData).ToList, p_strOutputFile, p_objEncoding)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Shared Sub sbDecryptToFile(ByVal p_strKey As String, ByVal p_blnData As Byte(), ByVal p_strOutputFile As String)
        Try
            Call sbDecryptToFile(p_strKey, p_blnData, p_strOutputFile, Encoding.ASCII)
        Catch ex As Exception
            Throw
        End Try
    End Sub


    'from Hex String
    Public Shared Function fnDecryptToBooleanFromHex(ByVal p_strKey As String, ByVal p_strData As String) As List(Of Boolean)
        Try
            Return fnDecryptToBoolean(p_strKey, HexToByteArray(p_strData))
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnDecryptToStringFromHex(ByVal p_strKey As String, ByVal p_strData As String, p_objEncoding As Encoding) As String
        Try
            Return fnDecryptToString(p_strKey, HexToByteArray(p_strData), p_objEncoding)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnDecryptToStringFromHex(ByVal p_strKey As String, ByVal p_strData As String) As String
        Try
            Return fnDecryptToString(p_strKey, HexToByteArray(p_strData), Encoding.ASCII)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnDecryptToByteFromHex(ByVal p_strKey As String, ByVal p_strData As String) As Byte()
        Try
            Return ToByteArray(fnDecryptToBoolean(p_strKey, HexToByteArray(p_strData)).ToArray)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Sub sbDecryptToFileFromHex(ByVal p_strKey As String, ByVal p_strData As String, ByVal p_strOutputFile As String)
        Try
            Call sbDecryptToFile(p_strKey, HexToByteArray(p_strData), p_strOutputFile)
        Catch ex As Exception
            Throw
        End Try
    End Sub


    'From File
    Public Shared Function fnDecryptFileToBoolean(ByVal p_strKey As String, ByVal p_strInputFile As String) As List(Of Boolean)

        Dim blnReturn As List(Of Boolean)
        Dim objFile As StreamReader
        Dim strFile As String

        Try

            blnReturn = New List(Of Boolean)

            objFile = New StreamReader(p_strInputFile)

            strFile = objFile.ReadToEnd()

            blnReturn = fnDecryptToBoolean(p_strKey, objFile.CurrentEncoding.GetBytes(strFile))

            objFile.Close()
            objFile.Dispose()
            objFile = Nothing

            Return blnReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnDecryptFileToHexString(ByVal p_strKey As String, ByVal p_strInputFile As String) As String

        Dim strReturn As String
        Dim bytReturn As Byte()

        Try

            strReturn = ""

            bytReturn = fnDecryptFileToByte(p_strKey, p_strInputFile)

            strReturn = BitConverter.ToString(bytReturn).Replace("-", "")

            Return strReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function fnDecryptFileToByte(ByVal p_strKey As String, ByVal p_strInputFile As String) As Byte()
        Try
            Return ToByteArray(fnDecryptFileToBoolean(p_strKey, p_strInputFile).ToArray)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Sub sbDecryptFileToFile(ByVal p_strKey As String, ByVal p_strInputFile As String, ByVal p_strOutputFile As String)

        Dim bytReturn As Byte()
        Dim objFileInput As StreamReader
        Dim objFileOutput As StreamWriter
        Dim strFile As String

        Try

            objFileInput = New StreamReader(p_strInputFile)
            objFileOutput = New StreamWriter(p_strOutputFile, False, objFileInput.CurrentEncoding)

            strFile = objFileInput.ReadToEnd()

            bytReturn = fnDecryptToByte(p_strKey, objFileInput.CurrentEncoding.GetBytes(strFile))

            objFileOutput.Write(bytReturn)

            objFileInput.Close()
            objFileInput.Dispose()
            objFileInput = Nothing

            objFileOutput.Close()
            objFileOutput.Dispose()
            objFileOutput = Nothing

        Catch ex As Exception
            Throw
        End Try
    End Sub
    '-----------------------------------------------------------------------------------------------

#End Region

#Region "Private Functions"
    'principal functions
    Private Shared Function fnEncrypt(ByVal p_objKeys As List(Of clsPairKey), ByVal p_strData As List(Of Boolean)) As List(Of Boolean)

        Dim objReturn As List(Of Boolean)
        Dim objAuxReturn As List(Of Boolean)
        Dim objInput As List(Of Boolean)
        Dim objAuxBlock As List(Of Boolean)
        Dim intIndex As Integer

        Try


            objReturn = New List(Of Boolean)
            objAuxReturn = New List(Of Boolean)
            objInput = p_strData

            intIndex = 0

            For Each Key As clsPairKey In p_objKeys

                objAuxBlock = New List(Of Boolean)

                objAuxBlock = objInput.GetRange(intIndex, Key.blockSize)
                If Key.direction Then
                    objAuxBlock = fnRotR(objAuxBlock, Key.rotations)
                Else
                    objAuxBlock = fnRotL(objAuxBlock, Key.rotations)
                End If

                objAuxReturn.AddRange(objAuxBlock)

                If (intIndex + Key.blockSize) > (objInput.Count - 1) Then
                    intIndex = 0
                    objInput = New List(Of Boolean)
                    objInput.AddRange(objAuxReturn)
                    objReturn = New List(Of Boolean)
                    objReturn.AddRange(objAuxReturn)
                    objAuxReturn = New List(Of Boolean)
                Else
                    intIndex = intIndex + Key.blockSize
                End If

            Next

            Return objReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Shared Function fnDecrypt(ByVal p_objKeys As List(Of clsPairKey), ByVal p_strData As List(Of Boolean)) As List(Of Boolean)


        Dim objReturn As List(Of Boolean)
        Dim objAuxReturn As List(Of Boolean)
        Dim objInput As List(Of Boolean)
        Dim objAuxBlock As List(Of Boolean)
        Dim intIndex As Integer

        Try
            objReturn = New List(Of Boolean)
            objAuxReturn = New List(Of Boolean)
            objInput = p_strData

            intIndex = 0

            objInput.Reverse()
            p_objKeys.Reverse()

            For Each Key As clsPairKey In p_objKeys

                objAuxBlock = New List(Of Boolean)

                objAuxBlock = objInput.GetRange(intIndex, Key.blockSize)
                If Key.direction Then
                    objAuxBlock = fnRotR(objAuxBlock, Key.rotations)
                Else
                    objAuxBlock = fnRotL(objAuxBlock, Key.rotations)
                End If

                objAuxReturn.AddRange(objAuxBlock)

                If (intIndex + Key.blockSize) > (objInput.Count - 1) Then
                    intIndex = 0
                    objInput = New List(Of Boolean)
                    objInput.AddRange(objAuxReturn)
                    objReturn = New List(Of Boolean)
                    objReturn.AddRange(objAuxReturn)
                    objAuxReturn = New List(Of Boolean)
                Else
                    intIndex = intIndex + Key.blockSize
                End If

            Next

            objReturn.Reverse()

            Return objReturn

        Catch ex As Exception
            Throw
        End Try

    End Function

    'Suport functions
    Private Shared Function fnRotL(p_objInput As List(Of Boolean), p_intRotations As Integer) As List(Of Boolean)

        Dim objReturn As List(Of Boolean)
        Dim blnAuxItem As Boolean

        Try

            objReturn = p_objInput

            For i = 0 To p_intRotations - 1

                blnAuxItem = objReturn.First
                objReturn.RemoveAt(0)
                objReturn.Add(blnAuxItem)

            Next

            Return objReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Shared Function fnRotR(p_objInput As List(Of Boolean), p_intRotations As Integer) As List(Of Boolean)

        Dim objReturn As List(Of Boolean)
        Dim blnAuxItem As Boolean

        Try

            objReturn = p_objInput

            For i = 0 To p_intRotations - 1

                blnAuxItem = objReturn.Last
                objReturn.RemoveAt(objReturn.Count - 1)
                objReturn.Insert(0, blnAuxItem)

            Next

            Return objReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Shared Function fnGenerateKeys(ByVal p_strKey As String, ByVal p_intSize As Integer) As List(Of clsPairKey)

        Dim objReturn As List(Of clsPairKey)
        Dim strBaseKey As String
        Dim strPrimaryKey As String
        Dim strSecondaryKey As String
        Dim intTotalSize As Integer
        Dim blnOk As Boolean
        Dim intInterator As Integer

        Try

            objReturn = New List(Of clsPairKey)

            intTotalSize = p_intSize

            strBaseKey = fnHashMD5String(p_strKey)
            strPrimaryKey = fnHashMD5String(fnGetSemiKeys(strBaseKey, False))
            strSecondaryKey = fnHashMD5String(fnGetSemiKeys(strBaseKey, True))

            intInterator = 0

            For i = 0 To 2

                blnOk = False
                Do While blnOk = False

                    objReturn.Add(New clsPairKey(fnGetNumberByKey(strBaseKey, intInterator), fnGetNumberByKey(strPrimaryKey, intInterator), If(fnGetNumberByKey(strSecondaryKey, intInterator) Mod 2 = 0, True, False)))

                    If objReturn.Last.blockSize >= intTotalSize Then
                        objReturn.Last.blockSize = intTotalSize
                        intTotalSize = p_intSize
                        blnOk = True
                    Else
                        intTotalSize = intTotalSize - objReturn.Last.blockSize
                    End If

                    intInterator = intInterator + 1
                Loop

            Next

            Return objReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Shared Function fnGetNumberByKey(ByVal p_strKey As String, ByVal p_intNumber As String) As Integer

        Dim intReturn As Integer
        Dim intNumber As Integer
        Dim objAuxBytes As Byte()
        Dim blnOk As Boolean

        Try

            intReturn = 0
            intNumber = p_intNumber
            objAuxBytes = Encoding.ASCII.GetBytes(p_strKey)
            blnOk = False

            Do While blnOk = False
                If intNumber > objAuxBytes.Length - 1 Then
                    intNumber = intNumber - (objAuxBytes.Length - 1)
                Else
                    blnOk = True
                End If
            Loop

            intReturn = objAuxBytes(intNumber)

            'limit to 1 ~ 32
            If intReturn = 0 Then
                intReturn = 1
            End If
            Do While intReturn > 32
                If intReturn > 32 Then
                    intReturn = intReturn - 32
                End If
            Loop

            Return intReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Shared Function fnGetSemiKeys(ByVal p_strKey As String, p_blnEven As Boolean) As String

        Dim strReturn As String

        Try

            strReturn = ""

            For i = 0 To p_strKey.Length - 1
                If (i Mod 2 = 0 And p_blnEven) Or (i Mod 2 <> 0 And Not p_blnEven) Then
                    strReturn = strReturn & p_strKey(i)
                End If
            Next

            Return strReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Shared Function fnHashMD5(ByVal p_strEntrada As String) As Byte()

        Dim Retorno As String
        Dim objMD5Provider As MD5CryptoServiceProvider

        Try

            Retorno = String.Empty

            objMD5Provider = New MD5CryptoServiceProvider
            Return objMD5Provider.ComputeHash(System.Text.Encoding.ASCII.GetBytes(p_strEntrada))

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Shared Function fnHashMD5String(ByVal p_strEntrada As String) As String
        Try
            Return BitConverter.ToString(fnHashMD5(p_strEntrada)).Replace("-", String.Empty)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Shared Function ToByteArray(p_objInput As Boolean()) As Byte()

        Dim bytReturn As Byte() = {}
        Dim intIndex As Integer = 0

        Try

            Array.Resize(bytReturn, p_objInput.Length / 8)

            For i = 0 To p_objInput.Length - 1 Step 8

                bytReturn(intIndex) = 0

                For i2 = 0 To 7

                    If p_objInput(i + i2) Then
                        bytReturn(intIndex) = bytReturn(intIndex) Or (1 << i2)
                    Else
                        bytReturn(intIndex) = bytReturn(intIndex) Or (0 << i2)
                    End If

                Next

                intIndex = intIndex + 1

            Next


            Return bytReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Shared Function FromByteArray(p_objInput As Byte()) As Boolean()

        Dim blnAuxReturn As List(Of Boolean) = New List(Of Boolean)
        Dim intIndex As Integer = 0

        Try

            For i = 0 To p_objInput.Length - 1

                For i2 = 0 To 7
                    blnAuxReturn.Add((p_objInput(i) And (1 << i2)) <> 0)
                Next

            Next

            Return blnAuxReturn.ToArray

        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Shared Function HexToByteArray(ByVal p_strHex As String) As Byte()

        Dim objAuxReturn As List(Of Byte)

        Try

            objAuxReturn = New List(Of Byte)
            For i = 0 To p_strHex.Length - 1 Step 2
                objAuxReturn.Add(Convert.ToByte(p_strHex.Substring(i, 2), 16))
            Next

            Return objAuxReturn.ToArray

        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#End Region

#Region "Properties"

#End Region

End Class

Friend Class clsPairKey
#Region "Declarations"
    Private intBlockSize As Integer
    Private intRotations As Integer
    Private blnDirection As Boolean
#End Region

#Region "Contructors"
    Public Sub New(ByVal p_intBlockSize As Integer, ByVal p_intRotations As Integer, ByVal p_blnDirection As Boolean)
        Try
            intBlockSize = p_intBlockSize
            intRotations = p_intRotations
            blnDirection = p_blnDirection
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Properties"
    ''' <summary>
    ''' Size of the bit block
    ''' </summary>
    ''' <returns></returns>
    Public Property blockSize() As Integer
        Get
            Return intBlockSize
        End Get
        Set(ByVal value As Integer)
            intBlockSize = value
        End Set
    End Property

    ''' <summary>
    ''' number of rotations
    ''' </summary>
    ''' <returns></returns>
    Public Property rotations() As Integer
        Get
            Return intRotations
        End Get
        Set(ByVal value As Integer)
            intRotations = value
        End Set
    End Property

    ''' <summary>
    ''' Direction of the rotation
    ''' true = right
    ''' false = left
    ''' </summary>
    ''' <returns></returns>
    Public Property direction() As Boolean
        Get
            Return blnDirection
        End Get
        Set(ByVal value As Boolean)
            blnDirection = value
        End Set
    End Property
#End Region
End Class
