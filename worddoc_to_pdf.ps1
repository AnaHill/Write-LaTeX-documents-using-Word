# create .docx -> .tex
. .\worddoc_to_tex.ps1
##############################
## build final pdf-document from .tex
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