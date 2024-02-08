mkdir target
ls src/*.md | each { |it|
    print $it.name
    # let name = echo $it.name | path basename | str replace ".md" ".docx" 
    (pandoc $it.name
        -o ("target/source.docx")
        --from markdown
        --to docx+native_numbering
        --reference-doc ./settings/custom-reference.docx
        --lua-filter ./settings/pagebreak.lua
        )
}