# AutoOpenXml

## 1 Why this lib exists?

### AutoOpenXml Was created for make easy the exportation of objects collections to the worksheet with theirs correct formats and smallest costs of implementation using  C# attributes annotations in classes and theirs fields

## 2 How to use 

### 2.1 Class example:

<code>

    [ExportWorkSheet(Variables.WorksheetName)]
    internal class ModelDateTimeProperty
    {
        public string Name { get; set; }
        [ExportColumn(Variables.FieldName, 1, ColumnTypes.DateTime)]
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

<code>

    [ExportColumn(
        {propertyLabel : string},
        {columnIndex : int},
        {columnType : ColumnTypes}
    )]

</code>

### 2.3 Column Types:

<code>

    public enum ColumnTypes
    {
        Text,
        Int,
        Decimal,
        DateTime
    }

</code>

## 3 More Details, see test cases on test project