Imports InventorRuleWrapper

Module Module1
    Sub Main()
        Try
            RuleRunner.RunMainFormRule()
            Console.WriteLine("Success.")
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub
End Module
