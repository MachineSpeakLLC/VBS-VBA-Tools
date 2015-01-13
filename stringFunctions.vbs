' --- substringBefore looks for an occurrence of the string value appearing in the "key"
'     argument within the string value held by the "text" argument.  If the key value is
'     found, then substringBefore will return the portion of "text" that precedes the "key"
Function substringBefore(text, key)
    keyPosn = InStr(text, key)
    retval = ""
    If keyPosn Then retval = Left(text, keyPosn - 1)
    substringBefore = retval
End Function

' --- substringBefore_ looks for an occurrence of the string value appearing in the "key"
'     argument within the string value held by the "text" argument.  If the key value is
'     found, then substringBefore will return the portion of "text" that precedes the "key"
'     If the key fragment is not found, then substringBefore_ will return the "text" argument in full
Function substringBefore_(text, key)
    keyPosn = InStr(text, key)
    retval = text
    If keyPosn Then retval = Left(text, keyAt - 1)
    substringBefore_ = retval
End Function

' --- substringAfter looks for an occurrence of the string value appearing in the "key"
'     argument within the string value held by the "text" argument.  If the key value is
'     found, then substringAfter will return the portion of "text" that precedes the "key"
Function substringAfter(text, key)
    keyPosn = InStr(text, key)
    retval = ""
    If keyPosn Then retval = Right(text, Len(text) - (keyPosn + Len(key) - 1))
    substringAfter = retval
End Function

' --- substringAfter_ looks for an occurrence of the string value appearing in the "key"
'     argument within the string value held by the "text" argument.  If the key value is
'     found, then substringAfter will return the portion of "text" that precedes the "key"
'     If the key fragment is not found, then substringAfter_ will return the "text" argument in full
Function substringAfter_(text, key)
    keyPosn = InStr(text, key)
    retval = text
    If keyPosn Then retval = Right(text, Len(text) - (keyPosn + Len(key)))
    substringAfter_ = retval
End Function

' --- substringBeforeLast looks for the last occurrence of the string value appearing in the "key"
'     argument within the string value held by the "text" argument. and returns the text that precedes that last occurrence
Function substringBeforeLast(text, key)
    revtext = StrReverse(text)
    revkey = StrReverse(key)
    revretval = substringAfter(revtext, revkey)
    substringBeforeLast = StrReverse(revretval)
End Function

' --- substringAfterLast looks for the last occurrence of the string value appearing in the "key"
'     argument within the string value held by the "text" argument. and returns the text that follows that last occurrence
Function substringAfterLast(text, key)
    revtext = StrReverse(text)
    revkey = StrReverse(key)
    revretval = substringBefore(revtext, revkey)
    substringAfterLast = StrReverse(revretval)
End Function

' --- contains returns trus if the value in "text" contains the value in "key" as s substring
Function contains(text, key)
    contains = InStr(text, key) > 0
End Function
