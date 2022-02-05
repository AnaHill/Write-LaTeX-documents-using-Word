# Dokumentin tiedot
$paper_name = 'WriteLaTeXusingWord' # input file
$Folder_path = $(get-location).Path + '\'; 
# Write-Output "Folder is: $Folder_path"s

$input_filename = $paper_name + '.docx'
$Fullpath_inputfile =  $Folder_path + $input_filename # docx
$osatiedosto = $Folder_path + 'partfiles/additional_parts.tex'

$output_filename_tex = $paper_name + '.tex'
$output_filename_pdf = $paper_name + '.pdf'

# temporal_naames
$out_draft_name_temp = 'temp.md' # poistetaan lopussa 
$out_draft_name = 'draft'
$out_draft_name_md = 'draft' + '.md'
$out_draft_name_tex = 'textfile' + '.tex'
$Fullpath_out_draft_name_md =  $Folder_path + $Fullpath_out_draft_name_md # md

#############################
## DOCX to MD
Write-Output ""
Write-Output "Converting Word-document to pdf (and LaTeX) document"
Write-Output "Folder is: $Folder_path"
Write-Output "Input: $input_filename -> $output_filename_pdf"
Write-Output "Step 1: converting .docx -> .md "
Write-Output ""
Write-Output "Converting $input_filename to $out_draft_name_md file "
# Convert .docx to temp.md
pandoc --atx-headers -s $Fullpath_inputfile --wrap=none -t markdown -o $out_draft_name_temp 
Write-Output ""
Write-Host "...temp file ($out_draft_name_temp) ready"
Write-Output ""
Write-Output "Modifying temp file"
# Modify temp file
(gc $out_draft_name_temp) | % {$_ -replace "\\textasciitilde{}", "~"} | sc $out_draft_name_temp
(gc $out_draft_name_temp) | % {$_ -replace "\\\\", "\"} | sc $out_draft_name_temp
(gc $out_draft_name_temp) | % {$_ -replace "\\\~", "~"} | sc $out_draft_name_temp
(gc $out_draft_name_temp) | % {$_ -replace "\\\$", "$"} | sc $out_draft_name_temp
(gc $out_draft_name_temp) | % {$_ -replace "\\\^", "^"} | sc $out_draft_name_temp
(gc $out_draft_name_temp) | % {$_ -replace "\\\[", "["} | sc $out_draft_name_temp
(gc $out_draft_name_temp) | % {$_ -replace "\\\]", "]"} | sc $out_draft_name_temp
(gc $out_draft_name_temp) | % {$_ -replace "\\\*", "*"} | sc $out_draft_name_temp
(gc $out_draft_name_temp) | % {$_ -replace "\\_", "_"} | sc $out_draft_name_temp
Write-Host "Replaced textasciitilde sign to correct in $out_draft_name_temp"
Write-Host "Removed extra \ in front of \,~,$,^,[,],*,_ in $out_draft_name_temp" 
# (gc $out_draft_name_temp) | % {$_ -replace "\\section{References}", ""} | sc $out_draft_name_temp
# Write-Host "Removed \section{References} from $out_draft_name_temp"

# Converting Abstract, References, and Acknowledgment to non-number headings
(gc $out_draft_name_temp) | % {$_ -replace "\# Abstract", "\section*{Abstract}"} | sc $out_draft_name_temp
(gc $out_draft_name_temp) | % {$_ -replace "\# Acknowledgment", "\section*{Acknowledgment}"} | sc $out_draft_name_temp
(gc $out_draft_name_temp) | % {$_ -replace "\# References", ""} | sc $out_draft_name_temp # delete reference heading if it comes from LaTeX / biblatex
# (gc $out_draft_name_temp) | % {$_ -replace "\# References", "\section*{References}"} | sc $out_draft_name_temp
# Write-Host "Converted Abstract and Acknowledgment to non-number headings"
Write-Host "Converted Abstract, References, and Acknowledgment to non-number headings"

Write-Host "Removing empty fig lines from $out_draft_name_temp and creating $out_draft_name_md, then deleting $out_draft_name_temp"
Get-Content $out_draft_name_temp | Where { $_ -notmatch "\!\[\]\(" } | Set-Content $out_draft_name_md
Write-Host "$out_draft_name_md created"
Write-Host "Removed empty fig lines from $out_draft_name_md"
Write-Host "Deleting $out_draft_name_temp"
Remove-Item  $out_draft_name_temp
Write-Host "...deleted"
Write-Host ""


##############################
## MD to TEX 
Write-Output "Converting $out_draft_name_md to .tex file "
# pandoc $out_draft_name_md -o $out_draft_name_tex  -f markdown-auto_identifiers -t latex
# pandoc $out_draft_name_md -o $out_draft_name_tex --wrap=preserve -f markdown-auto_identifiers -t latex
pandoc $out_draft_name_md -o $out_draft_name_tex --columns=120 -f markdown-auto_identifiers -t latex
Write-Host "$input_filename converted to $out_draft_name_tex"
Write-Output ""
Write-Output "Modifying .tex file"
(gc $out_draft_name_tex) | % {$_ -replace "\\textasciitilde{}", "~"} | sc $out_draft_name_tex
(gc $out_draft_name_tex) | % {$_ -replace ",height=\\textheight", ""} | sc $out_draft_name_tex
(gc $out_draft_name_tex) | % {$_ -replace "width=1.00000", "width=1"} | sc $out_draft_name_tex
(gc $out_draft_name_tex) | % {$_ -replace "begin{figure}", "begin{figure}[htb]"} | sc $out_draft_name_tex
(gc $out_draft_name_tex) | % {$_ -replace "\\section{References}", ""} | sc $out_draft_name_tex
(gc $out_draft_name_tex) | % {$_ -replace "\\tightlist", ""} | sc $out_draft_name_tex
Write-Host "Replaced textasciitilde sign to correct in $out_draft_name_tex"
Write-Host "Modified image size in figure enviroment $out_draft_name_tex"
Write-Host "Modified image size width"
Write-Host "Set positioning of all figures to [htb] in $out_draft_name_tex"
Write-Host "Removed \section{References} from $out_draft_name_tex"
Write-Host "Removed \tightlist from $out_draft_name_tex"
# Setting FloatBarrier to figure-environments
### NYT ei tartte jos figure - ympäristössä jo on
# (Get-Content $out_draft_name_tex) | 
    # Foreach-Object {
        # $_ # send the current line to output
        # if ($_ -match "\\end{figure}") 
        # {
        #   # Add Lines after the selected pattern 
            # "\FloatBarrier"
        # }
    # } | Set-Content $out_draft_name_tex
# Write-Host "Added \FloatBarrier to figure-environments in $out_draft_name_tex"

########
# jos Floatbarrier halutaan pois --> osat.tex
(gc $osatiedosto) | % {$_ -replace "^\\FloatBarrier", "%\FloatBarrier"} | sc $osatiedosto
# jos haluaa nuo takaisin
(gc $osatiedosto) | % {$_ -replace "^\%\\FloatBarrier", "\FloatBarrier"} | sc $osatiedosto


########
Write-Host ""
Write-Host ""
Write-Host "Draft tex file created"
### delete .md
Write-Host "Deleting $out_draft_name_md "
Remove-Item  $out_draft_name_md
Write-Host "...deleted"
Write-Host ""

##############################
## TEX to PDF
Write-Host "Creating pdf"
# bibtex.exe texbuildfiles/"BPM_sim_Draft" -include-directory=texbuildfiles
bibtex.exe texbuildfiles/$paper_name -include-directory=texbuildfiles
Write-Host "Bibliography created"
# pdflatex.exe -synctex=1 -interaction=nonstopmode --aux-directory=texbuildfiles $output_filename_tex

# pdflatex.exe -synctex=1 -interaction=nonstopmode --aux-directory=texbuildfiles $output_filename_tex
pdflatex.exe -synctex=1 -interaction=nonstopmode --aux-directory=texbuildfiles --output-directory=texbuildfiles $output_filename_tex
$output_fold_temp =  $Folder_path + 'texbuildfiles\' + $output_filename_pdf #
Write-Host "pdf created"

Write-Host "Moving $output_fold_temp"
Write-Host "to $Folder_path"
Move-Item -Path $output_fold_temp -Destination $Folder_path -Force
Invoke-Item $output_filename_pdf