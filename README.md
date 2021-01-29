
# Paulo José dos Santos(pjdsant@gmail.com).

# Product Documentation

# Winhctl.dll - This library can control components from the windows desktop via handler.


First of all, you need making a dll registration in operation system.

1 - Choose your folder to install the winctl.dll -> ie "C:\dll\winhctl.dll".

2 - Run the command regasm.exe to register it. -> ie "c:\C:\Windows\Microsoft.NET\Framework\v4.0.30319 C:\dll\winhctl.dll /codebase /tlb" 

Now you can use the winctl.dll from your javascript code.

<script language="javascript" type="application/javascript">
  var activeX = new ActiveXObject("PJSIT.WinHandlerControl");
  var activeEx = new ActiveXObject("PJSIT.WinHandlerControlEx");
</script>

![](https://github.com/pjdsant/winhctl/blob/master/regasm.png)

> ReAsm.exe.

# Winhctl methods available:

# 1 - GetText(string windowTitle, int index)

This method returns a string from TextBox identified in the window form application through the index.

windowTitle = Is the Windows Application Form Caption.
index = Is the index of the component within the form to be controlled.

# Sample using javascript:
<script language="javascript" type="application/javascript">
  var activeX = new ActiveXObject("PJSIT.WinHandlerControl");
  var text =  activeEx.GetText("Form Caption", 1);
</script>

# 2 - SendText(string windowTitle, int index, string message)

This method is void and sends a param message to a textbox in the any Windows Application Desktop.

windowTitle = Is the Windows Application Form Caption.
index = Is the index of the component within the form to be controlled.
message = The text to be written in the TextBox to be controlled.

# 3 - SendClick(string windowTitle, int index)

This method is void and sends a click command to a Button in the any Windows Application Desktop. 

windowTitle = Is the Windows Application Form Caption.
index = Is the index of the component within the form to be controlled.

# 4 - GetComboItem(string windowTitle, int index, int item)

This method returns a string from TextBox identified in the window form application through the index.

windowTitle = Is the Windows Application Form Caption.
index = Is the index of the component within the form to be controlled.
item = is the item in the combo to be obtained.

# 5 - SetComboItem(string windowTitle, int index, int item)

This method select item in Combo to be controlled.

windowTitle = Is the Windows Application Form Caption.
index = Is the index of the component within the form to be controlled.
item = is the item in the combo to be selected.

# Extra Method

# 1 - GetPhones(string MainForm, int phoneOne, int phoneTwo, string ChildForm, int phoneThree)

This method return phones from GEO screen in the Main Form and the child Form.
ie.

Note.: This method uses internally regex function to filter phone without special characters like "-, (, ),"


<script language="javascript" type="application/javascript">
  var activeEx = new ActiveXObject("PJSIT.WinHandlerControlEx");
  var phones =  activeEx.GetPhones("Main Form Caption", 1, 2, "Child Form Caption", 1);
</script>


# Sample

Using GetPhones Method to obtain three phone number from textBox on the forms (Main/Child).

![](https://github.com/pjdsant/winhctl/blob/master/Sample.png)

> Sample use GetPhones Method.



# Discovery Session

# SpyClass - This library can discovery all ClassName, Handler and Index from Window Caption.
To use spyclass.exe, follow these steps:

1 - Download spyclass.exe and install in your OS. ie “C:\PJSIT\spyclass.exe”.

2 - Run the command “.\spyclass.exe “Main Form Caption”


![](https://github.com/pjdsant/winhctl/blob/master/spyclassRun.png)

> SpyClass Command.


3 - After Run spyclass the SpyLog will be generated with all information:


![](https://github.com/pjdsant/winhctl/blob/master/spyClass.png)

> SpyClass Sample Use GetPhones method.

[SpyClassRepository](https://github.com/pjdsant/spyclass)

