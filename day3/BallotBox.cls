Option Explicit

Private votesA_ As Long
Private votesB_ As Long
Private electionName_ As String

Public Sub VoteForA()
    votesA_ = votesA_ + 1
End Sub

Public Sub VoteForB()
    votesB_ = votesB_ + 1
End Sub

Public Property Get Winner() As String
    If votesA_ > votesB_ Then
        Winner = "Candidate A"
    ElseIf votesB_ > votesA_ Then
        Winner = "Candidate B"
    Else
        Winner = "Tie"
    End If
End Property

Public Property Get ElectionName() As String
    ElectionName = electionName_
End Property

Public Property Let ElectionName(ByVal newName As String)
    If electionName_ = "" Then
        electionName_ = newName
    Else
        Err.Raise 61616, "BallotBox", "Tampering detected!"
    End If
End Property

Private Sub Class_Initialize()
    votesA_ = 0
    votesB_ = 0
    electionName_ = ""
End Sub
