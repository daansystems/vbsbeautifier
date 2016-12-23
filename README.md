# vbsbeatifier

VBScript beautifier beautifies VBScript
files.

Features:

- Works on serverside and clientside VBScript.
- Skips HTML.
- Properizes keywords.
- Splits Dim statements.
- Places spaces around operators.
- Indents blocks.
- Lines out assignment statements.
- Removes redundant endlines.
- Makes backups.


Instructions:
-------------
This is the Perl source code of the beautifier.
Run it from the commandline with parameters:
```
Usage: vbsbeaut [options] [files]

options:
 -i         Use standard input (as text filter).
 -s <val>   Uses spaces instead of tabs.
 -u         Make keywords uppercase.
 -l         Make keywords lowercase.
 -n         Don\'t change keywords.
 -d         Don\'t split Dim statements.
```

IMPORTANT: Make sure the VBScript code works before you
try to beautify it.

You can also use the commandline utility vbsbeaut.exe.
Enter vbsbeaut.exe without any arguments to see the commandline
options.

You can add vbscript keywords to the keywords.txt file.
