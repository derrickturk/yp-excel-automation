Option Explicit

Public Sub TestBallotBox()
    Dim box As BallotBox
    Set box = New BallotBox
    box.ElectionName = "City Council"

    box.VoteForA
    box.VoteForB
    box.VoteForA
    box.VoteForB
    box.VoteForA

    Debug.Print box.Winner

    ' box.ElectionName = "Mayor"
End Sub
