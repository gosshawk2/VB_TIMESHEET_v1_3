Public Class clsTotals
    Private _TotalOperatives As Long
    Private _TotalShortParts As Long
    Private _TotalExtraParts As Long
    Private _HighestOpTAGID As Long
    Private _HighestShortTAGID As Long
    Private _HighestExtraTAGID As Long
    Private _HighestOpBtnTAGID As Long
    Private _HighestOpTABIndex As Long
    Private _HighestShortTABIndex As Long
    Private _HighestExtraTABIndex As Long
    Private _TotalOpHours As Double
    Private _TotalFLMHours As Double
    Private _OpStartTAG As Long
    Private _OpFinishTAG As Long


    Public Property Total_Operatives() As Long
        Get
            Return _TotalOperatives
        End Get
        Set(value As Long)
            _TotalOperatives = value
        End Set
    End Property

    Public Property Total_ShortParts() As Long
        Get
            Return _TotalShortParts
        End Get
        Set(value As Long)
            _TotalShortParts = value
        End Set
    End Property

    Public Property Total_ExtraParts() As Long
        Get
            Return _TotalExtraParts
        End Get
        Set(value As Long)
            _TotalExtraParts = value
        End Set
    End Property

    '_HighestOpTAGID

    Public Property HighestOpTAGID() As Long
        Get
            Return _HighestOpTAGID
        End Get
        Set(value As Long)
            _HighestOpTAGID = value
        End Set
    End Property

    Public Property HighestShortTAGID() As Long
        Get
            Return _HighestShortTAGID
        End Get
        Set(value As Long)
            _HighestShortTAGID = value
        End Set
    End Property

    Public Property HighestExtraTAGID() As Long
        Get
            Return _HighestExtraTAGID
        End Get
        Set(value As Long)
            _HighestExtraTAGID = value
        End Set
    End Property

    Public Property HighestOpBtnTAGID() As Long
        Get
            Return _HighestOpBtnTAGID
        End Get
        Set(value As Long)
            _HighestOpBtnTAGID = value
        End Set
    End Property

    Public Property TotalOpHours() As Double
        Get
            Return _TotalOpHours
        End Get
        Set(value As Double)
            _TotalOpHours = value
        End Set
    End Property

    Public Property TotalFLMHours() As Double
        Get
            Return _TotalFLMHours
        End Get
        Set(value As Double)
            _TotalFLMHours = value
        End Set
    End Property

    'Private _OpStartTAG As Long
    'Private _OpFinishTAG As Long

    Public Property OpStartTAG() As Long
        Get
            Return _OpStartTAG
        End Get
        Set(value As Long)
            _OpStartTAG = value
        End Set
    End Property

    Public Property OpFinishTAG() As Long
        Get
            Return _OpFinishTAG
        End Get
        Set(value As Long)
            _OpFinishTAG = value
        End Set
    End Property

    'Private _HighestOpTABIndex As Long
    'Private _HighestShortTABIndex As Long
    'Private _HighestExtraTABIndex As Long
    'Inserted 10-SEP-2018

    Public Property HighestOpTabIndex() As Long
        Get
            Return _HighestOpTABIndex
        End Get
        Set(value As Long)
            _HighestOpTABIndex = value
        End Set
    End Property

    Public Property HighestShortTABIndex() As Long
        Get
            Return _HighestShortTABIndex
        End Get
        Set(value As Long)
            _HighestShortTABIndex = value
        End Set
    End Property

    Public Property HighestExtraTABIndex() As Long
        Get
            Return _HighestExtraTABIndex
        End Get
        Set(value As Long)
            _HighestExtraTABIndex = value
        End Set
    End Property

End Class
