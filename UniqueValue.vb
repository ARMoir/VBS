Module Module1

    Sub Main()

        Dim Info As Array

        Dim Array As String = ("Test,Test,Test,Test,New,New")

        Dim X

        Dim Z

        Dim Unique = ""

        Info = Split(Array, ",")

        X = 0

        Z = UBound(Info)

        Do

            X = X + 1

            Do

                Z = Z - 1

                If X = Z Then

                    Info(X) = Info(Z)

                ElseIf Info(X) = Info(Z) Then

                    Info(X) = ""

                End If

            Loop Until Z = 0

            Z = UBound(Info)

        Loop Until X = UBound(Info)

        For Each X In Info

            If X <> "" Then

                Unique = Unique & " " & X

            End If

        Next

        Console.WriteLine(Unique)

        Console.ReadLine()

    End Sub

End Module
