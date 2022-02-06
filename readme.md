This is short document how to write LaTeX documents using Word. Working principle is to write Word-document with some LaTeX-command that are used to include citations and references, figures, tables, and equations using LaTeX. Figure below illustrates the working principle. For more detailed information, see Word-document or converted pdf-file.

Simplified process is following. Firstly, write you document in Word with LaTeX-commands. Then, convert the Word-document to .tex using the Powershell script (worddoc_to_tex.ps1). If needed, upload (e.g. if using Overleaf) the converted .tex file to so that your main/root LaTeX document (WriteLaTeXusingWord.tex in this example) can find it. Then, you can build the final document e.g. using pdflatex. Flowcharts below illustrates the working principle.

![flowchart](figs/flowchart1.png)
![flowchart](figs/flowchart2.png)


# Acknowledgment
I warmly thank Juhana Ketola for his pro bono help related to PowerScipt. Thanks to my current and former colleagues, Sampo Tuukkanen, Mikko Peltokangas and Ville Rantanen helping with \LaTeX{}, especially introducing catchfilebetweentags,  developed by Florent Chervet. Special thanks for Florent, saved me a lot of time! Furthermore,  “tom’s” (sorry, couldn’t find author information, please share!) great [blog post in texblog](https://texblog.org/2012/12/04/keeping-things-organized-in-large-documents/) helped to implement catchfilebetweentags package.
