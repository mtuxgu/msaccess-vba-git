This simple module was created for my convenience because I occasionally have to develop 
some Code with VBA in Office-Apps and wanted to have that integrated in my GIT-Repos.

The module exports all of the code to a destination directory.
For every module (class, form, report, ...) a seperate file is created and also a
reasonable suffix is choosen.

Hosted on github to have it available when i need it...


Reference:
You must reference the "Microsoft Visual Basic for Applications Extensibility 5.3".

Usage:
Simply add the module to any of your VBA-Projects and put the desired path to
the const "c_ExportPath" and run the Main-Sub everytime you want to export your code.

ToDo:
- delete files if module is deleted (as an option)
- can the process be automated (execute whenever i safe a module...)
