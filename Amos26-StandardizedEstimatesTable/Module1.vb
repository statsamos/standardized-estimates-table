Imports System
Imports Microsoft.VisualBasic
Imports Amos
Imports AmosEngineLib
Imports AmosEngineLib.AmosEngine.TMatrixID
Imports MiscAmosTypes
Imports MiscAmosTypes.cDatabaseFormat
Imports System.Xml

<System.ComponentModel.Composition.Export(GetType(Amos.IPlugin))>
Public Class CustomCode
    Implements IPlugin

    Public Function Name() As String Implements IPlugin.Name
        Return "Standardized Estimates Table"
    End Function

    Public Function Description() As String Implements IPlugin.Description
        Return "A plugin to combine the SRW and RW tables into a single, easy to use format."
    End Function

    Public Function Mainsub() As Integer Implements IPlugin.MainSub

        'Ensure standardized estimates are checked.
        pd.GetCheckBox("AnalysisPropertiesForm", "StandardizedCheck").Checked = True

        'Fits the specified model.
        pd.AnalyzeCalculateEstimates()

        'Produce the output
        CreateOutput()

    End Function

    Sub CreateOutput()

        'Get the regression weights and standardized regression weights tables.
        Dim tableStdRegression As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody")
        Dim tableRegression As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Regression Weights:']/table/tbody")
        Dim numRegression As Integer = GetNodeCount(tableStdRegression) 'Number of rows in the regression weights table.

        If (System.IO.File.Exists("StandardizedEstimatesTable.html")) Then 'Check if the output file already exists.
            System.IO.File.Delete("StandardizedEstimatesTable.html")
        End If

        'Set up the listener to output the debugger.
        Dim debug As New AmosDebug.AmosDebug
        Dim resultWriter As New TextWriterTraceListener("StandardizedEstimatesTable.html")
        Trace.Listeners.Add(resultWriter)

        'HTML DOC
        debug.PrintX("<html><body><h1>Standardized Regression Weights</h1><hr/>")

        'Populate model fit measures in data table
        debug.PrintX("<table><tr><th>Predictor</th><th>Outcome</th><th>Std Beta</th></tr><tr>")

        For i = 1 To numRegression
            debug.PrintX("<td>" + MatrixName(tableStdRegression, i, 2) + "</td>") 'Name of variable 1
            debug.PrintX("<td>" + MatrixName(tableStdRegression, i, 0) + "</td>") 'Name of variable 2
            debug.PrintX("<td>" + MatrixName(tableStdRegression, i, 3)) 'Standardized regression weight

            'Output the significance of the estimate
            If MatrixName(tableRegression, i, 6) = "***" Then
                debug.PrintX("***</td>")
            ElseIf MatrixName(tableRegression, i, 6) = "" Then
                debug.PrintX("</td>")
            ElseIf MatrixElement(tableRegression, i, 6) = 0 Then
                debug.PrintX("</td>")
            ElseIf MatrixElement(tableRegression, i, 6) < 0.01 Then
                debug.PrintX("**</td>")
            ElseIf MatrixElement(tableRegression, i, 6) < 0.05 Then
                debug.PrintX("*</td>")
            ElseIf MatrixElement(tableRegression, i, 6) < 0.1 Then
                debug.PrintX("&#x271D;</td>")
            Else
                debug.PrintX("</td>")
            End If
            debug.PrintX("</td></tr>")
        Next

        'References
        debug.PrintX("</table><h3>References</h3>Significance of Estimates:<br>*** p < 0.001<br>** p < 0.010<br>* p < 0.050<br>&#x271D; p < 0.100<br>")
        debug.PrintX("<p>If you would like to cite this tool directly, please use the following:")
        debug.PrintX("Gaskin, J. & Lim, J. (2018), ""Merge SRW Tables"", AMOS Plugin. <a href=""http://statwiki.kolobkreations.com"">Gaskination's StatWiki</a>.</p>")

        'Write style And close
        debug.PrintX("<style>table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}</style>")
        debug.PrintX("</body></html>")

        'Take down our debugging, release file, open html
        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Process.Start("StandardizedEstimatesTable.html")
    End Sub

    'Get a string from an xml table.
    Function MatrixName(eTableBody As XmlElement, row As Long, column As Long) As String
        Dim e As XmlElement
        'The row is offset one.

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means the rows are not 0 based.
            MatrixName = e.InnerText
        Catch ex As NullReferenceException
            MatrixName = ""
        End Try

    End Function

    'Get a number from an xml table.
    Function MatrixElement(eTableBody As XmlElement, row As Long, column As Long) As Double

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means the rows are not 0 based.
            MatrixElement = CDbl(e.GetAttribute("x"))
        Catch ex As NullReferenceException
            MatrixElement = 0
        End Try

    End Function

    'Use an output table path to get the xml version of the table.
    Function GetXML(path As String) As XmlElement

        'Gets the xpath expression for an output table.
        Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
        doc.Load(Amos.pd.ProjectName & ".AmosOutput")
        Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
        Dim eRoot As Xml.XmlElement = doc.DocumentElement

        Return eRoot.SelectSingleNode(path, nsmgr)

    End Function

    'Get the number of rows in an xml table.
    Function GetNodeCount(table As XmlElement) As Integer

        Dim nodeCount As Integer = 0

        'Handles a model with zero correlations
        Try
            nodeCount = table.ChildNodes.Count
        Catch ex As NullReferenceException
            nodeCount = 0
        End Try

        Return nodeCount

    End Function

End Class
