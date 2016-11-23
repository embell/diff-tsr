diff-tsr
--------
diff-tsr script allows you to compare shared object repository files (.tsr) in WinMerge. 

Object Repository files (.tsr) are used by HP's Unified Functional Testing to save test objects. This script will use a UFT api to convert old and new ORs to XML, and then compare those files. The toolset I'm using this with is UFT, TortoiseSVN, and WinMerge. 

How to Use
----------
To install and use with TortoiseSVN: 

1. Copy diff-tsr.vbs into the directory "C:\Program Files\TortoiseSVN\Diff-Scripts"

2. Access the settings menu in TortoiseSVN (Right click -> TortoiseSVN -> Settings)

3. Choose "Diff Viewer" from the menu on the left

4. Click the "Advanced" button

5. Click the "Add" button

6. Set the file extension (first edit field) to 
        .tsr
        
7. Set the external program (second edit field) to 
        C:\Windows\SysWow64\wscript.exe "C:\Program Files\TortoiseSVN\Diff-Scripts\diff-tsr.vbs" %base %mine //E:vbscript

8. Navigate to an object repository in your working copy folder, and choose "Diff" or "Diff with previous version" to 
   verify that's it's working. You should see a semi-legible XML representation of the object repository.
