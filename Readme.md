# AutoOpenXML 
https://github.com/RuanVFachini/AutoOpenXML

## 1 Why does this lib exist?

### Using C# attributes annotations 'AutoOpenXML' was created to smooth the exportation and importation process of object collections to the worksheet, providing the correct formats without great cost of implementation.

## 2 How to use 

### 2.1 Class example:

<code>

    [ExportWorkSheet(Variables.WorksheetName)]
    internal class ModelDateTimeProperty
    {
        public string Name { get; set; }
        [ExportColumn(Variables.FieldName, 1)]
        public DateTime BirthDay { get; set; }

        public ModelDateTimeProperty() { }

        public ModelDateTimeProperty(string name, DateTime birthDay)
        {
            Name = name;
            BirthDay = birthDay;
        }
    }
</code>

### 2.2 Attributres syntax:

#### 2.2.1 Class Attributres:

<code>

    [ExportWorkSheet(
        {worksheetname : string}
    )]

</code>

#### 2.2.2 Property Attributres:

##### 2.2.2.1 To Export Column:

<code>

    [ExportColumn(
        {propertyLabel : string},
        {columnIndex : int},
        {numberFormat : string}
    )]
    
</code>

##### 2.2.2.2 Header Column Background:

<code>

    [ExportColumnHeaderBackgoundColor(
        {red: int},
        {green: int},
        {blue: int})]

</code>

### 2.3 Suported Types:

    -string
    -int
    -decimal
    -DateTime
    -int?
    -decimal?
    -DateTime?

## 3 Example

### 3.1 Model Class Exemple:

<code>


    [ExportWorkSheet("WorksheetName")]
    public class ModelOrderedProperties
    {
        [ExportColumn("Code", 2, "00000")]
        [ExportColumnHeaderBackgoundColor(0, 255, 0)]
        public int Id { get; set; }

        [ExportColumn("Name", 1)]
        public string Name { get; set; }

        [ExportColumn("Height", 4)]
        public decimal Height { get; set; }

        [ExportColumn("BirthDateLabel", 3)]
        public DateTime BrithDate { get; set; }

        public ModelOrderedProperties() { }

        public ModelOrderedProperties(
            int id,
            string name,
            decimal height,
            DateTime brithDate)
        {
            Id = id;
            Name = name;
            Height = height;
            BrithDate = brithDate;
        }
    }
</code>

### 3.1 Export Class Exemple:

<code>

    var stream = new ExportManagerBuilder<ModelOrderedProperties>()
                        .Init()
                        .SetData(Data.ToList())
                        .StartExportProcess();

    var workbook = new XLWorkbook(stream);

</code>

### 3.1 Import Class Exemple:

<code>

    var result = new ImportManagerBuilder<ModelDateTimeProperty>()
                    .OpenFile(StreamTestFile.GetStreamTestFile(), "Planilha1")
                    .Init()
                    .StartImportProcess();

</code>

## 4 Obs.:

* DateTime properties will be exported like DateTime Excel column format. For importation the same column type is required.
