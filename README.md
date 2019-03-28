# HSC Generation Pipeline

The purpose of this program is to take a Excel output from JMP and convert it into a format readable by the EpMotion. The program generates a protocol instructing the user to take the relevant steps needed to start the machine. The program uses a GUI input created with tkinter and has a clickable file for Mac and Windows Users.


# Setup

### For Mac and Linux users:

Download the files to your computer as a zip file. Click to open the HSC_Generation_Pipeline folder. Open the terminal and type

```
chmod u+x /path/HSC_Generation_Pipeline/JMP\ to\ Epmotion.command
```


Where ``` /path ``` is the current directory of the HSC_Generation_Pipeline folder. Once this is done, the program can be run by clicking the JMP to Epmotion.command. This is only required the first time the file is run.

## For all users:

Make sure that the following libraries are installed:

* xlrd
* numpy
* math
* tkinter

This can be easily achieved by installing Anaconda (Python 3) on to the computer.

# Usage

### For Mac and Linux Users:

Click the ```JMP to EpMotion.command``` file.

### For Windows Users:

Click the ```JMP to EpMotion.bat``` file.

## For all users:

0. Save an Excel output from JMP
1. Click the ```File Name``` button and choose the excel output file
2. Select the type of plate you will be using
3. Add the volume of each level you would need.
    i. A good rule of hand is to use ten percent extra (i.e. 110% of the required volume to be added)
4. Add the volume you would like to be added to each plate.
5. Follow instructions in the protocol file to setup experiment. 



# Notes

There are a few assumptions that the program makes that might be an issue for future experimentation

* A 24 well rack will be used as a source
* Only 96 and 384 well plates as outputs are currently supported
* Assumes that each factor needs the same volume (i.e. Is not optimised for fractional factorials)
