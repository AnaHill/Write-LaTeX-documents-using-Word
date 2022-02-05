This is short document how to write LaTeX documents using Word. Working principle is to write Word-document with some LaTeX-command that are used to include citations and references, figures, tables, and equations using LaTeX. Figure below illustrates the working principle. For more detailed information, see Word-document or converted pdf-file.


```mermaid
flowchart TB%%LR
    %% idWord --> |convert to .md | markdown
    idWord --> |convert| markdown
    idWord("Word <br> (WriteLaTeXusingWord.docx)")
    markdown("draft.md")
    markdown --> |"edit"| markdown
    %% markdown --> |"convert (.tex) <br>and edit <br>-> main_text.tex"| tex_to_pdf 
    %% markdown --> |"convert to .tex <br>and edit --> main_text.tex"| idmain  
    markdown --> |"convert to .tex <br>and edit"| idmain  
    subgraph tex_to_pdf[LaTeX]
        %% direction TB%% LR
        idtex("Main Tex document<br> (WriteLaTeXusingWord.tex)")
        idmain["main text<br>(main_text.tex)"]
        %% id_add["additional parts <br> (figures,equations,...)<br>(.\partfiles\additional_parts.tex) "]
        id_add["additional parts"]
        id_bib["bibliography"]
        idtex <--> id_add
        idtex <--> id_bib
        idmain <--> idtex
    end
    %% tex_to_pdf --> |"create .pdf <br> pdflatex"| idpdf
    tex_to_pdf --> |"pdflatex"| idpdf
    %% tex_to_pdf --> | | idpdf
    idpdf("pdflatex -> <br> (WriteLaTeXusingWord.pdf)")
    
    %% Set the classes. Sub=includes subfunctions, 
    %% classDef aloitus fill:#0000f0,color:#ffffff,stroke:#331,stroke-width:4px
    %% classDef lopetus fill:#ff0000,color:#ffffff 
    %% classDef word fill:#21dd9b,color:#ffffff,stroke:#331,stroke-width:4px
    classDef word fill:#0000f0,color:#ffffff,stroke:#331,stroke-width:4px
    classDef tex fill:#21a1dd,color:#ffffff,stroke:#331,stroke-width:4px 
    classDef pdf fill:#ff0000,color:#ffffff,stroke:#331,stroke-width:4px  

    %% Assign the classes.   
    %% class idc1,idc2, aloitus;
    %% class idc3, lopetus;
    class idWord, word;
    class idtex, tex;
    class idpdf, pdf;
```
