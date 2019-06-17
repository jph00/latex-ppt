## How to use LaTeX in PowerPoint

To use LaTeX in PowerPoint you have to complete a few setup steps first. (*I've only tested this on the latest Office 365 on Windows 10*.)

1. Download the latex PowerPoint addin from [here](https://github.com/jph00/latex-ppt/raw/master/latex.ppam)
1. Put the addin file somewhere convenient, and then add it to PowerPoint by clicking *File* then *Options*, clicking *Add-ins* in the options list on the left, then choose *PowerPoint Add-ins* from the *Manage* drop-down, and click *Go*. Choose *Add New* in the dialog box that pops up, and select the *latex.ppam* file you downloaded
1. Click *Enable Macros* in the security notice that pops up.

You'll now find that there's a new *LaTeX* tab in your ribbon. Each time you open a new PowerPoint session you'll need to switch it to "LaTeX mode". To do so, click inside a text box (so the cursor is flashing) and choose *Enable LaTeX* in the LaTeX tab. This file will now be in LaTeX mode until you close and reopen PowerPoint.

Now you are ready to insert your equation. Click inside a text box, and ensure the cursor is at the end of the text box (currently the macro only works if you're at the end of the selected text box). Now click *Paste LaTeX* in the LaTeX tab, and paste your equation into the input box that pops up (you can also type into it, of course, although I'd suggest you type your LaTeX into a regular text editor and paste it to PowerPoint from there, so you have a convenient source for all your equations' LaTeX source). That's it! The equation is now a regular PowerPoint equation, so when you click inside it, everything is editable, and you can also select the equation and change its font size, color, etc.

You can even select the equation and add Wordart effects to it, if you want to really ham things up!...

## Additional customization and tips

If you want to see the original LaTeX source again, click *Linear* on the *Equation* ribbon. However, *don't* edit this LaTeX directly in PowerPoint&mdash;it will mangle it as you type! Instead, copy it into an external editor and change it there, then create a new equation with the *Paste LaTeX* command as above. (This is why it's easier to simply keep all your original LaTeX source in a plain text file.)

Apparently Microsoft hates productivity, or at least that's the only reason I can think of that they decided to *remove* one of the most important features for productivity: the ability to customize and add keyboard shortcuts. So if you want to add a keyboard shortcut for *Paste LaTeX*, you instead have to right-click on the *Paste LaTeX* button in the ribbon, and choose *Add to Quick Access Toolbar*. You'll now see an extra button in the very top left of your window (that's the *Quick Access Toolbar*). Press and release *Alt*, and you'll be able to see what numeric shortcut has been assigned to that button. Press and release *Alt* again to remove the shortcut overlays. Now you're ready to use the keyboard shortcut. Click inside a textbox as before (at the end of it) and, while holding down *Alt*, press the number you noted down before. You should see the input box appear.
