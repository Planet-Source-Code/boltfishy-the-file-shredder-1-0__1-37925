Attribute VB_Name = "HELP"
' App Name                : The File Shredder
' App Version             : 1.0
' Purpose                 : Securely delete files
' Written by              : Mischa Balen
' Contact E-mail          : boltfish@eml.cc
' Website                 : http://boltfish.b0x.com
' Date Created            : July / August 2002
' Last Updated            : 13th August 2002
'----------------------------------------------------------

'CONTENTS:

'   1. Foreword
'   2. Global Variables
'   3. Application Model and Functions
'   4. Delete Clicked
'   5. Shred File Function
'   6. Hex Corrupt Function



'           Brief Foreword and Introduction
'      ----------------------------------------

'Hello all! Thanks for downloading my code. If you have
'downloaded it from Planet Source Code, then I'm not
'going to beg for 5 globes (it doesn't deserve them) but
'I would appreciate it greatly if you would offer some
'feedback, either on PSC or by email. Thank you. That way,
'I can make this a better programme, which would be great.

'This code is copyrighted and you may not resuse it in any
'way without my permission. Please respect that; I have
'been kind enough to share it with you in the first place.

'I recommend that you read this article through before you
'do anything else; just to get a feel for how the app
'works, if anything at all.

'Keep on coding!

'Mischa Balen aka ~boltfish~



'              GLOBAL VARIABLES - Useage
'      ----------------------------------------

'OK, so firstly lets deal with the global variables,
'which are:

' ---------------
'| NumberOfTimes |
'| HexCorrupt    |
'| FileTemp      |
' ---------------

'1. NumberOfTimes =
'This is a global variable which the user can set from
'frmOptions.frm. It stores the data telling us how many
'times we should overwrite the file. Upon startup, its
'value is automatically 1000. If the user changes it,
'then it is updated and called by the ShredFile function.
'The maximum value is 999,999,999.

'2. HexCorrupt =

'The user can also specify whether or not to corrupt the
'file(s). This works by writing completely random hex
'values to them. It is accessed via frmOptions2.frm
'It is therefore a Boolean and is called by the ShredFile
'Routine.

'3. FileTemp =
'When the user clicks 'delete file' it goes into a loop
'until all the files have been removed by the ShredFile
'routine. Before the file is deleted, its path is written
'to FileTemp. Then, we use the GetFileName sub to return
'the file name. This is so we can add the file name to
'the status bar panel even if it has just been deleted.

'          APPLICATION MODEL AND FUNCTIONS:
'      ----------------------------------------

'Now that we have got those sorted out, let's take a look
'at what happens when the use clicks the delete button.
'Here is the basic model process for the whole app:

'Delete Clicked -> File Name Read -> File Encrypted ->
'File OverWritten -> File Replaced with "" -> File Deleted

'    DELETE CLICKED - We declare the following:
'    ------------------------------------------

    ' ------------------------
    '| Dim i As Integer       |
    '| Dim b As Integer       |
    '| Dim File2Del As String |
    '| Dim msg As String      |
    ' ------------------------

'1. I / B = Counter. B is labelled as the number of files
'in the listbox - i.e. the number of files to be deleted.
'I can be thought of as the current file (which is being
'deleted).

'Using I, we progress in steps of 1.

'In every stage, we:
'
'   1)Set the display panel to "Deleting" i "of" b
'   2)Set the other panel to the file name of the current
'     file which is being deleted
'   3)Use ShredFile Function

'This loop runs until I = B, i.e. when the current file
'being deleted = the total number of files. Therefore we
'must have finished, so we take appropriate action.


'       SHRED FILE FUNCTION - Explained below:
'      ----------------------------------------

'This is the primary and most important feature in the
'programme. It makes the file safe before finally deleting
'it. Again, look at the model below:

'   1. Generates random characters
'   2. Overwrites data with random characters
'   3. Does this until NumberOfTimes is satisfied
'------------END OF MAIN OVERWRITING LOOP------------
'   4. Checks to see if we should HexCorrupt the file (= True)
'   5. Corrupts the file IF HeckCorrupt is True
'   6. Overwrites all the characters with ""
'   7. Deletes the file
'   8. Repeats until every file has been deleted

'In order to generate the random data, we generate a random
'number - Rnd(*255).
'Then we take as a character code so we can convert it to
'a character data.

'Once this has been done, the file is opened for binary
'and we replace every character in it with the random
'character data we just generated.
'This process goes on until NumberOfTimes has been
'satisfied.
'We must flush the file buffers. If windoze sees that
'we are going to delete the file anyway it won't
'bother to overwrite it etc, so we use this API call
'in order to clear its "memory".

'If the user wants to hex corrupt the file (we look at
'the value of HexCorrupt), then we do so at this point.

'Finally, we remove all the data in the file and replace it
'with "".

'The file is now deleted.

'      HEX CORRUPT FUNCTION - Explained below:
'     -----------------------------------------

'How this works:

'Basically, this writes hex values straight to the file(s).
'It generates 4 totally random numbers:

'A
'B
'C
'D

'These are then used to replace the data in the file in the
'following weird equation:

'Str$(((((mTemp * a) / c) * d * b) * i / 3.141592654 * Sqr(a)))

'Because so many random entries are used, I don't think that
'any decryption could ever be written for it. If you can
'crack it, please e-mail me via boltfish@eml.cc. Thanks!

'So it essentially corrupts the file before deleting it.









