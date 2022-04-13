# This script will convert 
    # \section --> \chapter
    # \section* --> \chapter* (non numbered, e.g Abstract)
    # \subsection --> \section
    # \subsubsection --> \subsection
    # Optional
        # \paragraph --> \subsubsection
        # \subparagraph --> \paragraph
    
$modified_file_name = 'main_text.tex' # give file name to be edited
(gc $modified_file_name) | % {$_ -replace "\\section{", "\chapter{"} | sc $modified_file_name
(gc $modified_file_name) | % {$_ -replace "\\section*{", "\chapter*{"} | sc $modified_file_name # abstract etc non numbered chapters
(gc $modified_file_name) | % {$_ -replace "\\subsection{", "\section{"} | sc $modified_file_name 
(gc $modified_file_name) | % {$_ -replace "\\subsubsection{", "\subsection{"} | sc $modified_file_name  
# optional: convert \paragraph to \subsubsection and \subparagraph to \paragraph
# (gc $modified_file_name) | % {$_ -replace "\\paragraph{", "\subsubsection{"} | sc $modified_file_name  # typically not used
# (gc $modified_file_name) | % {$_ -replace "\\subparagraph{", "\paragraph{"} | sc $modified_file_name  # typically not used
