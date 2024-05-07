# VBScript Custom Shell

An alternative shell created in VBScript that allows you to perform basic file system operations and execute Windows operating commands if you are unable to run cmd or powershell due to blocking (:

**🌟 Highlights**

**Navigation in the file system:** Use **cd [path]** to change directory or **cd ..** to go one level up.

**File listing:** Use the **dir** command to list files and directories in the current directory.

**Directory management:** Create directories with **mkdir [directory_name].**

**File manipulation:** Copy files with **copy [source_file] [destination].**

**Operational commands:** Run common commands such as **whoami, ipconfig, curl** and many others.

**⚠️ Limitations**

**Although many operational commands can be executed, this shell is more limited than the standard Windows CMD.**

**Execution of VBS scripts may be disabled in environments with strict security policies.**

**The syntax for certain commands may differ or certain features may not be available.**

**🚀 Usage**

1. Save the script as **vbshell.vbs** or another name of your choice with the extension **.vbs.**

2. Run the .vbs file by double-clicking or via the command line.

3. An input window will be displayed showing the current directory. Enter the desired command and press OK.

4. Continue entering commands as necessary. Close the input window to exit.
