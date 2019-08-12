**Tutorial**

Download via:

	git clone https://github.com/ConchoPecan/ReportMod.git <Local Path>

The easiest way to get started is to cd on over to the parent folder of ReportMod.py. This requirements opening up either command prompt or Terminal. We will assume for this tutorial that we're using Windows (I know, Linux users get no respect, but you're in my heart and I feel your pain). For example, let's assume we are a local user and we downloaded the this repo to a special git folder in Documents:

	cd C:\Users\ConchoPecan\Documents\git\ReportMod

If you are unsure you're in the right folder, use dir (or ls) and if you see **ReportMod.py** listed, you're in the right place. Now, ensure you have python 3 installed on this computer. If you unsure, check the version of both python and pip!

	pip --version
	python --version

You are looking for **at least Python 3.7.3** Older versions are not explicitly tested for, though I haven't personally ran into any problems with them. Please let me know! Once you are sure that you are indeed cooking with oil, type python OR python3 OR launch into Python 3.7+, however you have to do that. You now see python 3.7 OR ABOVE in the first line of the python startup script. If you don't find Python 3.7 or above on your system.

The first thing you do is to type:

	import ReportMod as RM

ReportMod is a data-oriented program. That means that you do not have to mess with classes and what not. All you need to do is use the functions and variables that ReportMod provides. Let's first look at Docx. Docx has two functions you need to be aware of:

	ReportMod.Docx.RedactRegex
	ReportMod.Docx.BlurImages

**ReportMod.Docx**

Find the path of where the document you want to change. We'll assume it's on your Documents under reports. Now, find the folder you want to put it in. In our case, it's under documents but in a folder called redacted. Here's the two filepaths for the places:

	c:\Users\ConchoPecan\Documents\reports\CasePaper.docx
	c:\Users\ConchoPecan\Documents\redacted

As you may notice, nothing is behind redacted. Our script will create the redacted information. Do so now:

	RM.Docx.RedactRegex(r"c:\Users\ConchoPecan\Documents\reports\CasePaper.docx", r"Sensitive Information", r"*****", r"c:\Users\ConchoPecan\Documents\redacted\CasePaper.docx")

This will look in CasePaper.docx and look for the words Sensitive Information. Right now, it can see into tables, captions and regular text. As this is still a work in progress, I hope to ensure that it will soon find every single instance.

We may find simple regex redaction insufficient. While not ready, you will be able to redact multiple regexes soon (there's a LOT on my wishlist!). But today, you can blur images as well on docx!

	RM.Docx.BlurImages(r"c:\Users\ConchoPecan\Documents\redacted\CasePaper.docx")

You should notice a new file created that says CasePaper.redacted.docx. This is the file with the blurred images. You now two redacted files, one with only the regex redacted, and the other with regex redacted and blurred. You can delete the one with just the redacted words. Command line solves this problem for you.

**ReportMod.HTML**

HTML is a lot easier to work with, so it has more features. Before you EVER do anything with RM.HTML, you must set the report up. We'll assume HTML reports are in the same area as the docx report was:

	RM.HTML.SetReports(r"c:\Users\ConchoPecan\Documents\reports\HTMLReport", r"c:\Users\ConchoPecan\Documents\redacted\HTMLReport")

We know have a HTML report that is EXACTLY indentical to original report. Any time we shrink, blur or use regex, it is affecting the copy. HTML is a little safer this way. Docx will one day be this safe.

RM.HTML.ShrinkImages:

	RM.HTML.ShrinkImages()

Simple, as it should be. Usually I shrink before blurring to save time. Doing so will typically mean that when you blur later, your images become just indistinct color stains though, so if you got the time, blurring then shrinking will be better.

RM.HTML.BlurImages:

	RM.HTML.BlurImages()

Sometimes, you need more than shrinkage to prevent damaged psyches. Blurring will do that for you. This is probably the most time intensive part of ReportMod though, so take that in consideration.

RM.HTML.RedactRegex:

	RM.HTML.RedactRegex(r"Sensitive Information", r"*****")

This does the same thing as it does in Docx. Takes the report, and ANYWHERE that exists in the HTML document... provided it's in text form and not in code pages.

The great thing about HTML is that once you set the report, everything you do will only affect the copy of the report, and nothing else.

**Command Line**
