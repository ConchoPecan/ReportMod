# ReportMod
Report Modifier POC - Change Your Reports For Third Parties Without Rewriting The Dang Thing Again

**Purpose:**
	ReportMod is a tool that takes a single report, copies it, and modifies the copied report so that it can be handed to third parties. The modifications it does to the copied report are done to redact or obscure information.
	ReportMod does not destroy or modify the original report. Nor is this POC feature-complete or beta-ready; this POC is meant only to showcase the application and purpose of this tool. To this end, efforts have been made to ensure safety of the original reports over speed, unusual report structure and supporting non-intended workflows. This tool is NOT intended to stay in this state. I personally hope to support it for some time to come, fixing every edge-case I can and adding many, many more features as this project matures.

**Requirements:**

	Python 3.7+
	python-docx
	pillow
	Element Tree

**To Install**

	git clone https://github.com/ConchoPecan/ReportMod.git <desired homepath for cloned repo>
	cd <path to local ReportMod repo>
	pip install -r requirements.txt

**Supported File Formats:**

	Docx
	html
	Future Formats: PDF, doc, odt, xlsx, xls

**Usage:**
	ReportMod.py [--html | --docx] <ifile/ifolder> -o <ofile/ofolder> [-e regex] [-d redactedtext] [-b] [-s] [-r]

**Arguments:**

	--html		Input html folder Flag – If --docx is not chosen, this is necessary
	--docx		Input docx file Flag – If --html is not chosen, this is necessary
	-i		The name of the initial report - necessary
	-o		Name of output folder or file – necessary
	-e		Regex String to search for and replace - Optional
	-s		String to replace matches - Default is ***** - Optional
	-b		Blur all images of the report - Optional
	-t		Shrink all images of the report - HTML only right now – Optional

**Current POC Limitations**

**Limited Error Handling** – Error handling is not robust – this tool is not yet guaranteed for fitness of duty or for widespread release. In the beta, there will be planned Robust error handling, with all known current error messages and types documented and will be handled with a specialized error class and procedures at the release of the beta.

**Multiple optional flags** – The capability to do multiple redactions or to do blur and shrinking efficiently is not yet implemented. This means that a single redaction should be handled differently than multiple, and the folder/file should only have to be read in once, rather than multiple times as it is done now. This is planned to be fixed in the beta.

**Word XML avoidance** – Right now, word's xml structures are largely left alone. This is currently being tackled, but it was determined to be too complex to be completely tackled before the POC. Indeed, some known issues are due to XML not being updated after images are modified. Current python libraries modify and change the xml only in certain ways, which means that to resolve these issues, word's xml libraries have to be manually loaded.

**No Auto-detection of input files** – docx and html files aren't automatically figured on the fly. You do have to specify what you are working with.

**No Config files** – Right now, to change what the program considers to be text files, code files, pictures or even picture sizes, you have to modify the source code. It is planned, yet not implemented, to have the tool read from configuration files to determine these things, as well as files/images to exclude.

**Upzipping is only done by scripting** – Right now, ReportMod supports unzipping both docx and html files... but only when using ReportMod as a library.

**No Doc, PDF, Odt and xlsx files are supported** – Regrettably, not all file types are yet supported. Doc files require separate libraries from docx files, so do xls and xlsx files. And pdfs... will require much more testing. While all these files are specified for the beta, they are not implemented or supported for the POC.

**No named groups for Regex** – In an announcement that'll surely annoy the one or two regex-nerds in the group (Hey fellow nerds!), named groups are not yet supported. This is a personal requirement.

**Reporting:**
	Even though this is just a POC, I am very interested in feedback. Please email me at bcleveland@nw3c.org if you run into any problems... even ones that you may suspect that I'm aware of. Chances are, you may be the first one to stumble on it. If you encounter any issues, please let me know. This allows me to update ReportMod not based on some personal goalposts but rather informed real world usage. I look forward to hearing from you!
