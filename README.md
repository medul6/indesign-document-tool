InDesign Document Tool
=============

Description
-------


This is an AppleScript for modifying the pages in all open documents. Like:

* "Alle Kapitelanfänge löschen" -- remove all sections (except for the first of course ...)
* "Seiten löschen" -- delete pages
* "Seiten einfügen" -- insert pages
* "Seiten verschieben" -- move pages

After executing the script in InDesign you'll be presented with a "function" dialog.
This is where you can choose from one of the functions above.
After that it depends on the function you chose.
"remove all sections": doesn't have options, it does what it says.
"delete pages": You can enter page ranges like "2-3,8-19" or "1,3,5,7,9" but you have to do that in order(!): '2-3,8-19' NOT '8-19,2-3'"
"insert pages": ["How many pages should be inserted?"] ONLY integers, even or odd! e.g. "2" or "7"; ["After which page should the pages be inserted?"] ONLY integers, no ranges!
"move pages": ["Which pages should be moved?"] Only connected spreads, with hyphen or comma separated! e.g. "4-5" or "4,5"; ["After which page the pages to be moved?"] ONLY integers, no ranges!



Disclaimer
-----
	Use my scripts at your own risk! I am not responsible for any damages to your InDesign Documents!
	With my Repositories I just want to give the world back what I have got from others who share their code –
    usable, productive AppleScripts!

Contributing
------------
Want to contribute? Great! You sure know what to do, I am new to Github so I don't know if I am doing this right :)

File-Formats
-----------
    Normally, in InDesign, I use the *.scpt format because it it precompiled but at the same time contains the code and is viewable via QuickLook.
    I'm also commiting an *.applescript file, just because it it readable (If you just want to take a look) on github and the precompiled is not.
    Both types are usually commited together – always – if not, then it wasn't necessary.

Localization
-----------
    At the moment the script is mixed, english and german. with the documentation beeing mostly in english but most strings are in german.
    As far as I researched it is not trivial to localize AppleScripts. We (the studio I work in) use these scripts

Installation
-----------
	My scripts should work in a variety of InDesign Versions which is wy I use the Application ID instead of the name.
	But they should definitly work in the latest InDesign Version. We update our Adobe apps as soon as a new version comes out.
    Put the file in the application folder "Adobe InDesign CSx" > "Scripts" > "Scripts Panel"

Usage
-----
    open InDesign and some documents and start the script from the "Scripts Panel"
