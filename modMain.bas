Attribute VB_Name = "modMain"
Option Explicit

'-------------------------------------------------------------
'This program is simply a map creator and editor that works
'with binary files. I wrote this program not for my own use,
'but as a reference to anyone that needs to learn how to
'access binary files and retreive information from them
'accruately and effecinently. My spelling isn't the best
'in the world today, so please try and bear with me.
'This program was written in Visual Basic 6.0 Enterprise
'Edition. Enough technical stuff...here is a little info about me:
'My name is Joe Estock and I (obviously) am a computer programmer.
'I write many commercial applications, as well as code fragments and
'small business applications. I started programming about six or seven
'years ago and have stuck with it ever since. The first programming
'environment I started in was Borland C++; I do not recall the version
'but it was a rather old version (a new one at the time). I then moved
'on to Visual C++ 6.0 and became rather bored as well as upset
'with the way the code is organized so I gave Visual Basic a try
'since I used to write QBASIC applications. Visual Basic quickly
'became an indispensable tool for me that has paid for itself
'many times. And here I am now. I release the source code to most
'of my products (I am not greedy or stingy like Microsoft, even though
'I do like them, but ony because of Visual Studio and Windows 2000
'Advanced Server), however I do not release the source code to many
'of my commercial products because I usually sell the source code to
'anotehr company and they throw their copyrights on it and do not
'allow me to share. This is it, I have taken up a ratehr large chunk
'of your screen, and you probabally haven't even read anything I have
'written beyond "This program". How do I know? Because I am the same way.
'I like to dive right in and learn what I need to know, dry myself off and
'start producing whatever it is I need to produce. Why am I still talking?
'I have no clue ;)
'-------------------------------------------------------------
'I got tired of typin all the colors, so
'I assigned them to a variable. Pretty intellegent aren't I? ;)
Global Const black = &H0&
Global Const grey = &H808080
Global Const green = &H8000&
Global Const red = &HC0&
Global Const fDebug = True

'This is the heart of our code
'Slot0 is, you guessed it: lblMap index number 0
'If you don't know what Slot1 through Slot164 is
'then you better not become a programmer (Just Kidding)
'We also have the variable for the author, title, difficulty,
'and a password
Public Type MapDefinition
    Slot0 As Long
    Slot1 As Long
    Slot2 As Long
    Slot3 As Long
    Slot4 As Long
    Slot5 As Long
    Slot6 As Long
    Slot7 As Long
    Slot8 As Long
    Slot9 As Long
    Slot10 As Long
    Slot11 As Long
    Slot12 As Long
    Slot13 As Long
    Slot14 As Long
    Slot15 As Long
    Slot16 As Long
    Slot17 As Long
    Slot18 As Long
    Slot19 As Long
    Slot20 As Long
    Slot21 As Long
    Slot22 As Long
    Slot23 As Long
    Slot24 As Long
    Slot25 As Long
    Slot26 As Long
    Slot27 As Long
    Slot28 As Long
    Slot29 As Long
    Slot30 As Long
    Slot31 As Long
    Slot32 As Long
    Slot33 As Long
    Slot34 As Long
    Slot35 As Long
    Slot36 As Long
    Slot37 As Long
    Slot38 As Long
    Slot39 As Long
    Slot40 As Long
    Slot41 As Long
    Slot42 As Long
    Slot43 As Long
    Slot44 As Long
    Slot45 As Long
    Slot46 As Long
    Slot47 As Long
    Slot48 As Long
    Slot49 As Long
    Slot50 As Long
    Slot51 As Long
    Slot52 As Long
    Slot53 As Long
    Slot54 As Long
    Slot55 As Long
    Slot56 As Long
    Slot57 As Long
    Slot58 As Long
    Slot59 As Long
    Slot60 As Long
    Slot61 As Long
    Slot62 As Long
    Slot63 As Long
    Slot64 As Long
    Slot65 As Long
    Slot66 As Long
    Slot67 As Long
    Slot68 As Long
    Slot69 As Long
    Slot70 As Long
    Slot71 As Long
    Slot72 As Long
    Slot73 As Long
    Slot74 As Long
    Slot75 As Long
    Slot76 As Long
    Slot77 As Long
    Slot78 As Long
    Slot79 As Long
    Slot80 As Long
    Slot81 As Long
    Slot82 As Long
    Slot83 As Long
    Slot84 As Long
    Slot85 As Long
    Slot86 As Long
    Slot87 As Long
    Slot88 As Long
    Slot89 As Long
    Slot90 As Long
    Slot91 As Long
    Slot92 As Long
    Slot93 As Long
    Slot94 As Long
    Slot95 As Long
    Slot96 As Long
    Slot97 As Long
    Slot98 As Long
    Slot99 As Long
    Slot100 As Long
    Slot101 As Long
    Slot102 As Long
    Slot103 As Long
    Slot104 As Long
    Slot105 As Long
    Slot106 As Long
    Slot107 As Long
    Slot108 As Long
    Slot109 As Long
    Slot110 As Long
    Slot111 As Long
    Slot112 As Long
    Slot113 As Long
    Slot114 As Long
    Slot115 As Long
    Slot116 As Long
    Slot117 As Long
    Slot118 As Long
    Slot119 As Long
    Slot120 As Long
    Slot121 As Long
    Slot122 As Long
    Slot123 As Long
    Slot124 As Long
    Slot125 As Long
    Slot126 As Long
    Slot127 As Long
    Slot128 As Long
    Slot129 As Long
    Slot130 As Long
    Slot131 As Long
    Slot132 As Long
    Slot133 As Long
    Slot134 As Long
    Slot135 As Long
    Slot136 As Long
    Slot137 As Long
    Slot138 As Long
    Slot139 As Long
    Slot140 As Long
    Slot141 As Long
    Slot142 As Long
    Slot143 As Long
    Slot144 As Long
    Slot145 As Long
    Slot146 As Long
    Slot147 As Long
    Slot148 As Long
    Slot149 As Long
    Slot150 As Long
    Slot151 As Long
    Slot152 As Long
    Slot153 As Long
    Slot154 As Long
    Slot155 As Long
    Slot156 As Long
    Slot157 As Long
    Slot158 As Long
    Slot159 As Long
    Slot160 As Long
    Slot161 As Long
    Slot162 As Long
    Slot163 As Long
    Slot164 As Long
    Title As String
    Author As String
    Password As String
    Difficulty As Integer
End Type

'A reference to our custom Type
Global Map As MapDefinition

'Hmm...I forgot what this code does...oh yes, it saves the file ;)
Public Function SaveMap(sFileName As String)
On Error GoTo SaveMapError
    'To save our file, we simply open a binary file
    'for output and save everything contained in our
    'custom type that we made at the beginning of this
    'module. No mess, no hassle, and completely free.
    
    'First let's get our own unused file handle
    'so that we don't cause any errors with other
    '(possible) open files.
    Dim FileNum As Long
    FileNum = FreeFile()
    Open sFileName For Binary As #FileNum Len = Len(sFileName)
    'Write our data in the file...
    Put #FileNum, , Map
    '...close the file...
    Close #FileNum
    '...Tell grandpa that everything went as planned...
    SaveMap = "Success"
    '...and exit before we hit our error handler
    Exit Function
SaveMapError:
    'fDebug is a variable I used to turn on and off debugging
    'lines without deleting sever lines of code. Look at the top of
    'this module if you want to turn it off. If you compile it this
    'way, it won't hurt anything and it will be completely invisible
    'to the end user, unless they are running it under Visual Basic
    #If fDebug Then
        'Houston, we have a problem! Ed done busted out the capsule
        'window trying to hit a satelite with a beer bottle!
        Debug.Print "Error " & Err.Number & " " & Err.Description
    #End If
    'Grandpa is gonna be mad...we got ourselves into some trouble
    'we might wanna go back and look at the code for an error in
    'either coding, or possibly we might be trying to write to
    'a write-protected disk. Whatever the case may be, we certainly
    'didn't save our file.
    SaveMap = "Error"
End Function
