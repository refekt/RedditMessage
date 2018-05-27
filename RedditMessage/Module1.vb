Imports System
Imports System.IO
Imports RedditSharp
Imports RedditSharp.Things
Imports System.Diagnostics

Module Module1

    Sub Main()

        Dim startRM As String ' Used to read input of whether or not you want to continue or not

        Console.WriteLine(" Do you want to start messaging? Y/N")
        startRM = Console.ReadLine()

        ' Check if user wants to start the messages.
        If startRM.ToUpper = "N" Then
            Console.WriteLine(" Exiting application. Press any key to continue...")
            Console.ReadKey()
            Exit Sub
        End If

        If startRM.ToUpper <> "Y" Then

            Console.WriteLine(" Unexpected input. Application exiting. Press any key to continue...")
            Console.ReadKey()
            Exit Sub

        End If

        ' Get the path of the executable 
        Dim exePath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) ' Get the current location of the .exe
        exePath = Right(exePath, Len(exePath) - 6) ' Remove "file:/" from the string

        Console.WriteLine(" Which set of orders do you want to issue?" & vbNewLine & "  1. General Orders" & vbNewLine & "  2. Advanced Orders")

        Dim whichOrders As String
        Dim loadOrders() As String
        Dim messageText As String
        Dim usersTo() As String ' Arrays to hold users and opted-out users
        Dim optedOutUsers() As String
        Dim messageSubject As String

        whichOrders = Console.ReadLine()

        If whichOrders = 1 Then

            usersTo = File.ReadAllLines(exePath & "\GenUsers.txt") ' Read text files into array; super simple btw

            loadOrders = File.ReadAllLines(exePath & "\GenOrders.txt") ' The text file with the actual message in it
            messageText = String.Join(vbNewLine, loadOrders) ' Converts each index (paragraphs) into a single string with new lines between each one.

            messageSubject = "/r/HuskersRisk GENERAL Daily Orders" ' Message subject

        ElseIf whichOrders = 2 Then

            usersTo = File.ReadAllLines(exePath & "\AdvUsers.txt") ' Read text files into array; super simple btw

            loadOrders = File.ReadAllLines(exePath & "\AdvOrders.txt") ' The text file with the actual message in it
            messageText = String.Join(vbNewLine, loadOrders) ' Converts each index (paragraphs) into a single string with new lines between each one.

            messageSubject = "/r/HuskersRisk ADVANCED Daily Orders" ' Message subject

        Else

            Console.WriteLine(" Unexpected input. Exiting application. Press any key to continue...")
            Console.ReadKey()
            Exit Sub

        End If

        optedOutUsers = File.ReadAllLines(exePath & "\OptedOut.txt")

        ' Bot username and password
        Dim username As String = "*"
        Dim password As String = "*"

        Dim redObj As Reddit = New Reddit(username, password, True) ' Opens the bot account ... will also crash the app if the username/pw is wrong

        Dim i As Integer, j As Integer
        Dim skipMe As Boolean

        Dim intLowRnd As Integer = 7000 ' Minimum seconds to wait
        Dim intHighRnd As Integer = 14000 ' Maximum seconds to wait

        Console.WriteLine(" Starting messaging sequence.")

        ' Loop through every username loaded into usersTo
        For i = LBound(usersTo) To UBound(usersTo)

            ' This process can slow down the program if there becomes a lot of users
            For j = LBound(optedOutUsers) To UBound(optedOutUsers) ' Highly inefficient!
                Debug.WriteLine(usersTo(i) & vbTab & optedOutUsers(j))
                If usersTo(i) = optedOutUsers(j) Then
                    skipMe = True
                    Console.WriteLine(" Skipping due to opt-out : " & optedOutUsers(j))
                End If
            Next j ' Highly inefficient!

            If skipMe = False Then ' Only run if not opted-out 

                Try
                    Console.WriteLine(" [" & i & "] - Sending mesasge to: ** " & usersTo(i) & " **")

                    redObj.ComposePrivateMessage(messageSubject, messageText, usersTo(i))

                Catch ex As Exception

                    ' If an error is caught it'll skip the user it was caught on. Example error is a 500 Internal Server Error
                    Console.WriteLine(" Exception: " & ex.Message)
                    Console.WriteLine(" Skipping")

                End Try

                If i <> UBound(usersTo) Then

                    Console.WriteLine(" [" & i & "] - [" & DateTime.Now & "]Waiting random duration...")

                    Threading.Thread.Sleep(CInt(Math.Floor((intHighRnd - intLowRnd + 1) * Rnd()))) ' Pause processing 

                    Console.WriteLine(" [" & i & "] - [" & DateTime.Now & "]Waiting complete.")

                End If

                Console.WriteLine("")

            End If

            skipMe = False

        Next i

        Console.WriteLine(" All users complete. Press any key to close.")
        Console.ReadKey()

    End Sub

End Module
