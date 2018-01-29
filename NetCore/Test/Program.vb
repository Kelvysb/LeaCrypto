Imports System
Imports LeaCrypto

Module Program
    Sub Main(args As String())

        Dim strKey As String
        Dim strData As String
        Dim objauxData As String

        Do
            Try
                Console.WriteLine("Insert key...")
                strKey = Console.ReadLine()

                Console.WriteLine("Insert Data...")
                strData = Console.ReadLine()

                Console.WriteLine("Encrypt Result:")
                objauxData = LeaCryptoEngine.fnEncryptToHexString(strKey, strData)
                Console.WriteLine(objauxData)

                Console.WriteLine("Decrypt Result:")
                Console.WriteLine(LeaCryptoEngine.fnDecryptToStringFromHex(strKey, objauxData))

            Catch ex As Exception
                Console.WriteLine("Error " & ex.Message)
            End Try

        Loop

    End Sub
End Module
