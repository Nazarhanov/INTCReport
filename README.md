# INTCReport

Python программа для преобразования отчета из .yaml в .docx.

### Запуск:

```
# dependencies
pip install -r requirements

# generate
python run.py <file-path>

#example
python run.py dist/oop-example/oop.yaml
```

### Технологии:

Python 3.8, Python-docx, Yaml

### Зачем?

Во время практики часто нужно писать отчет. При этом нужно соблюдать стандарты форматирования текста и изображений (две пустых строки перед каждой темой, слеженение за нумерованием изображений, изменение номеров страниц..). Это было больно, учитывая постоянные исправления отчетов.

### Что умеет?

1. Учитывает все стандарты форматирования отчета.
2. Нумерует все названия тем и изображений
3. Заполняет секцию "Содержимое", указывая темы и соответствующие им номера страниц.
4. Все что нужно, это лишь написать отчет-шаблон в формате yaml.

### Пример отчета:

На вход вскармливаем [oop.yaml](https://github.com/Nazarhanov/INTCReport/blob/master/dist/oop-example/oop.yaml)
На выходе получаем [oop.out.docx](https://view.officeapps.live.com/op/embed.aspx?src=http://nazarhanov.github.io/INTCReport/oop.out.docx)
