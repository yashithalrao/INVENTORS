Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.IO

Public Class RuleRunner
    Public Shared Sub RunMainFormRule()
        Dim inventorApp As Inventor.Application = Nothing

        Dim names = Assembly.GetExecutingAssembly().GetManifestResourceNames()
        For Each n In names
            Console.WriteLine("📦 Found: " & n)
        Next

        Try
            inventorApp = Marshal.GetActiveObject("Inventor.Application")
        Catch
            inventorApp = Activator.CreateInstance(Type.GetTypeFromProgID("Inventor.Application"))
            inventorApp.Visible = True
        End Try

        If inventorApp.Documents.Count = 0 Then
            Console.WriteLine("❌ No document open in Inventor.")
            Return
        End If

        Dim oDoc As Document = inventorApp.ActiveDocument
        Dim iLogicAddIn As ApplicationAddIn = inventorApp.ApplicationAddIns.ItemById("{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}")

        If Not iLogicAddIn.Activated Then iLogicAddIn.Activate()

        Dim iLogic As Object = iLogicAddIn.Automation
        If iLogic Is Nothing Then
            Console.WriteLine("❌ iLogic Automation not available.")
            Return
        End If

        ' 🔒 Step 1: Load embedded MAIN_FORM.vb from resources
        Dim ruleText As String = GetEmbeddedRuleText()
        Dim tempRuleName As String = "hidden"

        '' 💉 Step 2: Inject rule
        'iLogic.AddRule(oDoc, tempRuleName, ruleText)

        '' 🚀 Step 3: Run it
        'iLogic.RunRule(oDoc, tempRuleName)

        '' 🧹 Step 4: Delete it
        'iLogic.DeleteRule(oDoc, tempRuleName)

        'Console.WriteLine("✅ Hidden rule injected, executed, and deleted.")
        Try
            iLogic.AddRule(oDoc, tempRuleName, ruleText)
            iLogic.RunRule(oDoc, tempRuleName)
        Catch ex As Exception
            Console.WriteLine("❌ Error running rule: " & ex.Message)
        Finally
            Try
                iLogic.DeleteRule(oDoc, tempRuleName)
                Console.WriteLine("🧹 Temp rule deleted.")
            Catch exDel As Exception
                Console.WriteLine("⚠️ Could not delete temp rule: " & exDel.Message)
            End Try
        End Try
    End Sub

    Private Shared Function GetEmbeddedRuleText() As String
        ' This assumes you added MAIN_FORM.vb to the project and marked as "Embedded Resource"
        Using stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("InventorRuleWrapper.MAIN_FORM.vb")
            Using reader As New StreamReader(stream)
                Return reader.ReadToEnd()
            End Using
        End Using
    End Function
End Class
