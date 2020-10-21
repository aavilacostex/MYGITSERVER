Public Class LikeSession
    Public Shared dsData As DataTable
    Public Shared dsDatagridview1 As DataSet
    Public Shared dsDgvProjectDetails As DataSet
    Public Shared dsErrorSession As DataSet
    Public Shared dsResultsSession As DataSet
    Public Shared flyingValue As String
    Public Shared acceptChanges As Boolean
    Public Shared gridEnable As Boolean = False
    Public Shared isPageLoad As Boolean = True
    Public Shared referencedExistence As String
    Public Shared currentAction As String
    Public Shared panelCollapseProp As Boolean
    Public Shared flagAccessAllow As Integer
    Public Shared searchControls As List(Of Object)
    Public Shared retrieveUser As String
    Public Shared focussedControl As Control
    Public Shared fullFilePath As String
    Public Shared objToFill As String
End Class
