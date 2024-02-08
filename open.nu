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