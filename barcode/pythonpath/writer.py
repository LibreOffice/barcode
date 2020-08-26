
def getWord( view ):
    try:
        cursorpos = view.getViewCursor()
        if cursorpos.Cell is not None:
            text = cursorpos.Cell
        elif cursorpos.Text is not None:
            text = cursorpos.Text
        cursor = text.createTextCursorByRange( cursorpos )
        cursor.collapseToEnd()
        cursor.gotoStartOfWord( False )
        cursor.gotoEndOfWord( True )
        return cursor
    except:
        # there is a large number of cases (such as when an inserted image is selected) when we can not get the word
        return None
