# Генератор, конвертатор и вообще на дуде игрец md2docx

## БАЗАААААА

1. pandoc - нельзя нормально декларативно настроить стили для fucking ворда
2. python-docx - списки приняли ислам, и нумерация у них одна на весь документ

## ЭТО УЖЕ ГЕНШТАБ

Объединим сии дав инструмента и получим:

1. pandoc+filter: md2docx (Просто генератор ворда, фильтр на разрывы страниц)
2. python-docx: docx2docx (Для +- меняемого декларативного изменения стилей)

## Траблы с python-docx

Изменение стиля `Заголовок 1`, ака `Заголовок первого уровня`, ака `Heading 1`

1. Просто нельзя обратиться через `styles["Heading 1"]`, ибо он в открыту.
   1. Мне говорит, что такого стиля не существует
   2. Но если генерировать docx с нуля, то существует. Что за траблы я не знаю. 
   3. Обошел сей закидон с помощью цикла и сопаставления с образцом (именем стиля), ака паттерн матчинг по хаскеллвски
2. Трабл с заголовками, они отказываются штатным образом ставить ~~православный~~ САТАНИНСКИЙ `Times New Roman`
   1. Для обхода данных траблов обратиться к [Пояснению за код](/docs/Пояснение_за_код.md)
3. Трабл с стилями списков. Я в душе не знаю как их изменить, ~~поэтому скитаюсь в поисках~~
   1. Когда найду как починить, напишу наверное `Пояснение_за_код_2.md`