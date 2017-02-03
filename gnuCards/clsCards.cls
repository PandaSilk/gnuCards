VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private intCardStyle As Integer
Private intShuffled As Boolean

'Public Function DeckShuffle()
'0 to 51
'End Function

'Public Function DeckGetCard()
'
'End Function

Public Enum CardStyle
 TechnoGray = 0
 TechnoRed = 1
End Enum



Public Sub SetCardStyle(CardStyle As CardStyle)
Attribute SetCardStyle.VB_Description = "Sets the style for the back of the card."
 intCardStyle = CardStyle
End Sub

Public Function SetCardBack() As StdPicture
Attribute SetCardBack.VB_Description = "Retrieves the image of the back of the card."
 Set SetCardBack = frmCards.imgBack(intCardStyle)
End Function





Public Function SetCardImage(Index As Integer) As StdPicture
Attribute SetCardImage.VB_Description = "Retrieves the image of the card."
 Dim Tmp
 Tmp = Index - 1
 If Index < 0 Or Index > 51 Then
 Err.Raise Number:=vbObjectError + 32144, Description:="Only 52 cards in a deck"
 Exit Function
 End If
 Set SetCardImage = frmCards.imgCard(Tmp)
End Function

Public Function GetCardSuit(x As Integer) As String
Attribute GetCardSuit.VB_Description = "Retrieves the suit in the form of an UPPERCASE string."
 Dim n As Integer
 Dim Suit As String
 n = x \ 13
 Select Case n
 Case 0
 Suit = "SPADES"
 Case 1
 Suit = "CLUBS"
 Case 2
 Suit = "DIAMONDS"
 Case 3
 Suit = "HEARTS"
 End Select
 GetCardSuit = Suit
End Function

Public Function GetCardValue(x As Integer) As String
Attribute GetCardValue.VB_Description = "Retrieves the card value in the form of an UPPERCASE string."
 Dim n As Integer
 Dim value As String
 n = x Mod 13
 Select Case n
 Case 0
 value = "ACE"
 Case 1 To 9
 value = Str$(n + 1)
 Case 10
 value = "JACK"
 Case 11
 value = "QUEEN"
 Case 12
 value = "KING"
 End Select
 GetCardValue = value
End Function

