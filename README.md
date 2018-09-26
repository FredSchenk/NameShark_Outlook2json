# NameShark_Outlook2json
Export a selection of your Outlook-contacts in json-format, ready for import in NameShark

1. Import the bas, frm and frx into your Outlook VBA-environment
2. Set the ROOTFOLDER-constant to your liking (easiest to put it straight into a Dropbox folder)
  example: Const ROOTFOLDER = "C:\Users\Fred.Schenk\Documents\Dropbox\NameShark-Outlook2json\"
3. Set the sGroup and sFilter to your liking (see the examples in the bas-file)
4. Run the NS_CreateNamesharkJSON-macro

NB: Make sure the json-file doesn't already exists

TODO:
- remove an already existing json-file
- set the creation of the group-folder (for the contact pictures) outside of the loop

PS: The progress-form is reused by me for this maxcro. I've created it long ago and it
    might seem a bit too much for this particular kind of macro. Feel free to adjust it
    and/or reuse it for your own code. As with all the code provided it's GPL'd.
