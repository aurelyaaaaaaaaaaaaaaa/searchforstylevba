# Search for Style VBA

## How this script works
1. The script asks the user for a style to search for
2. Script searches the document for paragraphs using that style
3. The script will try to put it into a table that lists all the paragraphs with that style (This is denoted by an alt text description of Style Name + " Table")
4. If the script doesn't find a table for that specific style it will create one
5. The script adds all the paragraphs of that style to the table in the format below
6. If ran again the script will exclude the already included paragraphs

## For example for a table of reccomendataions
| No. | Content | Status |
|-----|---------|--------|
| 1 | Rec 1 | Open |
| 2 | Rec 2 | Open |
| 3 | Rec 3 | Open |