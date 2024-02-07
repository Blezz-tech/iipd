let targetDir = "./target/"
let targetFileName = "Аналитический_отчет.docx"
let targetFile = $targetDir + $targetFileName

def main [] {
    echo "Билдится Аналитический_отчет.md"
    mkdir $targetDir
    (pandoc $targetFileName
        -o $targetFile
        --from markdown
        --to docx
        --reference-doc ./custom-reference.docx)
    
    let os = sys | get host | get name
    match $os {
        "Linux" => {
            print 'Linux'
            xdg-open $targetFile
        },
        "Windows" => {
            print 'MS Windows'
            start $targetFile
        },
        _ => {
            print 'Other OS'
        }
    }
}
