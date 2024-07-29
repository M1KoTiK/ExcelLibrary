Возможности на данный момент: 
- Просмотр и изменение свойств документа Excel таких как: автор документа, автор последнего изменения, дата создания и редактирования документа.
- Просмотр таблиц которые есть в файле Excel (Наименование таблицы и другая информация о ней)

Пример того как будет выглядеть работа с библиотекой:
```cs
            var path = new FileInfo("C:\\path");
            var ex = new ExcelFile(path);
            Console.WriteLine(ex.DocInfo.Author);
            Console.WriteLine(ex.DocInfo.ModifyDate);
            ex.DocInfo.Author = "New Author";
            Console.WriteLine(ex.DocInfo.Author);
            foreach (var el in ex.Sheets)
            {
                Console.WriteLine(el.Name);
            }
```
Консольный вывод:

![image](https://github.com/user-attachments/assets/a1c6aa62-7a39-4dd9-8b88-d59faf529dfc)

На данный момент это все что может данная библиотека
