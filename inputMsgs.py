menuCmd = """Select command:
        (F)ind rows that contain exact phrase
        (W)rite cost value for rows that contains all exact phrase
        (H)elp
        (E)xit program
"""

menuHelp = """
'(F)ind' and '(W)rite' both work by matching rows that contain ALL phrases that are within title column.
Ex) ['Apple Cider', '$4.00]
Searching with these two phrase will return rows that contain both 'Apple Cider' and '$4.00' within title column.
You can also do '! Apple cider', which will mean that title must NOT contain 'Apple cider' in title.
Note: Capilizations are ignore. 'Apple' will match with 'aPPle'.

Examples with phrases of ['Apple Cider', '$4.00]:
'Apple cider $4.00'             - Match. 'Apple cider' and '$4.00' both found.
'Apple cid $4.00 Cider'         - No match because 'Apple cider' are not together
'Apple   cider $4.00'           - No match because 'Apple' and 'Cider' contain too much extra spaces, so it does not match with 'apple cider'
'Apple cider Apple cider $4.00' - Match. Phrase being found twice is the same as phrase being found once. Both phrases were found.

Examples with phrases of ['Apple Cider', '! $4.00']:
'Apple cider $4.00'             - No Match. 'The '!' means it must not contain '$4.00'.
'Apple cider $3.00'             - Match. Contains 'Apple Cider' and does not contain '$4.00'
"""
