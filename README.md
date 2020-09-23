<div align="center">

## Automation of the Word Spell Checker


</div>

### Description

To help out my Dutch friend Frederik who needed to spell check in various languages, I designed this simple example.

To adapt it to use a particular dictionary, supply the dictionary path as an option like so:

w1.CheckSpelling(Text1.Text,"c:\path\MyDic.DIC)
 
### More Info
 
Works better if you type one word at a time.

Of course, you'll write you code to automate this.

You'll need to own a fairly recent copy of Word. I used Word 97.

Tells you if the spelling is good or a list of suggestions if not.

Lot's of votes I hope!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Edward Colman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-edward-colman.md)
**Level**          |Intermediate
**User Rating**    |3.6 (18 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-edward-colman-automation-of-the-word-spell-checker__1-12126/archive/master.zip)

### API Declarations

```
No Windows API.
Create a form with a TextBox CommandButton and ListBox.
Also include the Word library by selecting these menu options:
 Project > References:
   and checking "Microsoft Word 8.0 Object Library".
```


### Source Code

```
Option Explicit
'Create a reference to the Word Automation Object
Dim w1 As Word.Application
Private Sub Command1_Click()
  Dim I As Variant
  'Empty the list box
  List1.Clear
  'Check the spelling of the word...
  'If not in dictionary, fill a list box with suggestions
  If w1.CheckSpelling(Text1.Text) = False Then
    Beep
    For Each I In w1.GetSpellingSuggestions(Text1.Text)
      List1.AddItem I
    Next
    If List1.ListCount = 0 Then
      List1.AddItem "No suggestions"
    End If
  Else
    List1.AddItem "Spelling Correct"
  End If
End Sub
Private Sub Form_Load()
  'Open a new instance of Word
  Set w1 = New Word.Application
  'Create a new document (necessary)
  w1.Application.Documents.Add
  'Disable the following line if you don't want to see Word
  w1.Visible = True
End Sub
Private Sub Form_Terminate()
  'Quit, ignoring changes
  w1.Quit False
  Set w1 = Nothing
End Sub
```

