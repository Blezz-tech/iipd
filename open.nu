def main [] {
    mkdir target
    ls src/*.md | each { |it|
        print $it.name
        let name = echo $it.name | path basename | str replace ".md" ".docx" 
        (pandoc $it.name
            -o ("target/" + $name)
            --from markdown
            --to docx
            --reference-doc ./custom-reference.docx)
    }
    
    let os = sys | get host | get name
    match $os {
        "Linux" => {
            print 'Linux'
            xdg-open ./target/Аналитический_отчет.docx
        },
        "Windows" => {
            print 'MS Windows'
            start ./target/Аналитический_отчет.docx
        },
        _ => {
            print 'Other OS'
        }
    }
}
