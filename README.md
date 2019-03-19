# HSC Generation Pipeline

The purpose of this program is to take a Excel output from JMP and convert it into a format readable by the Epmotion. The program generates a protocol instructing the user to take the relevant steps needed to start the machine. The program uses a GUI input created with tkinter and has a clickable file for Mac and Windows Users.


# Setup

### For Mac users:

Download the files to your computer as a zip file. Click to open the HSC_Generation_Pipeline folder. Open the terminal and type

```
chmod u+x /path/HSC_Generation_Pipeline/JMP\ to\ Epmotion.command
```


Where ``` /path ``` is the current directory of the HSC_Generation_Pipeline folder. Once this is done, the program can be run by clicking the JMP to Epmotion.command. This is only required the first time the file is run.

# Notes

There are a few assumptions that the program makes that might be an issue for future experimentation

* A 24 well rack will be used as a source
* Only 96 well plates as outputs are currently supported
* Assumes that each factor needs the same volume (i.e. Is not optimised for fractional factorials)
