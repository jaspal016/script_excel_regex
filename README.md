# Script for using regex in Excel

VBS Script to use regex in MS Excel. Please refer to https://stackoverflow.com/a/43128681
for detailed explaination on using this code.

For convenience, here are the instructions as copied from Source 1 given below.

1. Open Excel workbook.
2. Alt+F11 to open VBA/Macros window.
3. Add reference to regex under **Tools** then **References**
4. Select **Microsoft VBScript Regular Expression 5.5**
5. Insert a new module (code needs to reside in the module otherwise it doesn't work)
6. Save the module and return to excel spreadsheet

Find ex. Type "apple", "ball" and "cat" in A1, A2 and A3. Then fill corresponding cells in
   B column with the formula **=Regex(A1, "..*a.*$")** to find words with the character
   'a' in the middle of the word.

Replace ex. Type "hello world" in cell A1, then go to B1 and simply type =Regex(A1, "world", "user")
   to get the result "hello user" using regex.

Source: https://stackoverflow.com/a/43128681

Source: https://www.oreilly.com/library/view/vbscript-in-a/1565927206/re155.html
