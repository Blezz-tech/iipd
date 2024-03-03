mkdir bin
ls src/*.md | each { |it|
    print $it.name
    # let name = echo $it.name | path basename | str replace ".md" ".docx" 
    (pandoc $it.name
        -o ("bin/source.docx")
        --from markdown+four_space_rule
        --to docx+native_numbering
        # --reference-doc ./settings/custom-reference.docx
        --lua-filter ./settings/pagebreak.lua
        )
}