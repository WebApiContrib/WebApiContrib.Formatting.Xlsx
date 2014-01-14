Excel for Web API
=================

**Currently in beta.**

A smart, customisable Excel `MediaTypeFormatter` for ASP.NET Web API.


Features
--------

- Control over column names, order, number format and visibility of properties in generated Excel documents via custom attribute.
- Lots of formatting options:
  - freeze header rows;
  - add auto-filter;
  - customize header row and cell styles/heights; and
  - autofit column widths to data.
- Decent (and improving) unit test coverage.


Limitations
-----------

`ExcelMediaTypeFormatter` can only serialise collections that implement `IEnumerable<>` and does not work with nested/complex item types. Extending the range of types this formatter can work with is a priority for future work.


Setting it up
-------------

First, you will need to add a reference to **EPPlus** to your project—either [download it from CodePlex][epplus-codeplex] or [grab the package on NuGet][epplus-nuget].

Next, add the `ExcelMediaTypeFormatter` to the formatter collection in your Web API configuration. This will look something like:

```C#
config.Formatters.Add(new ExcelMediaTypeFormatter()); // Where config = System.Web.Http.HttpConfiguration.
```

You are now good to go forth and serialise; however, you may find the generated Excel output a tad boring. Enter the advanced formatter instantiation options!


### Advanced setup options

The `ExcelMediaTypeFormatter` provides a number of options for improving the appearance of generated Excel files.


#### `autoFit`
**Default `true`.** Fit columns to the maximum width of the data they contain.


#### `autoFilter`
**Default `false`.** Set the column headers up as an auto-filter on the data.


#### `freezeHeader`
**Default `false`.** Split the top row of cells so that the column headers stay at the top of the window while scrolling through data.


#### `headerHeight`
**Default `0` (i.e. not set).** Set the height of the column header row.

#### `cellHeight`
**Default `0` (i.e. not set).** Set the height of the data cells.


#### `cellStyle`
**Default `null`.** Can take an `Action<OfficeOpenXml.Style.ExcelStyle>` that specifies visual formatting options (such as fonts and borders) for all cells.


#### `cellStyle`
**Default `null`.** Can take an `Action<OfficeOpenXml.Style.ExcelStyle>` that specifies visual formatting options (such as fonts and borders) for only the column header cells.


#### Advanced setup example

```C#
var formatter = new ExcelMediaTypeFormatter(autoFit: true,
                                            autoFilter: true,
                                            freezeHeader: true,
                                            headerHeight: 20f,
                                            cellHeight: 18f,
                                            cellStyle: (ExcelStyle s) => s.WrapText = true,
                                            headerStyle: (ExcelStyle s) => s.Border.Bottom.Style = ExcelBorderStyle.Double
                                           );

config.Formatters.Add(formatter);
```


Controlling serialisation output with `ExcelAttribute`
------------------------------------------------------

You can control how data gets serialised into columns using `ExcelAttribute` on individual properties.


### Set column header names

Header names can be provided using the `Header` parameter.

```C#
[Excel(Header = "Column header")]
public string Value { get; set; }
```

### Set number format for data cells

The `NumberFormat` parameter allows you to provide a [custom Excel number format][number-format] to alter the number format used for data in a given column. For example, the following snippet will align numbers on the decimal point:

```C#
[Excel(NumberFormat = "???.???")]
public decimal Value { get; set; }
```


### Ignore a column

To prevent a property from appearing as a column in the generated Excel file, set the `Ignore` parameter to true.

```C#
[Excel(Ignore = true)]
public string Value { get; set; }
```


### Customise display order

By default, columns are ordered according to property order in the source, with properties in derived classes coming before those in base classes. However, sometimes you need more control over the order in which columns appear.

The `Order` parameter works similarly to [JSON.NET's `JsonPropertyAttribute.Order`][json-net]. Properties are serialised from lowest to highest `Order` parameter value, and properties with the same `Order` value are in source order. By default, all properties are assumed to have an `Order` value of -1.

Confusing? Here's an example showing all of these rules at work:

```C#
// This will be second, because it has an implicit Order of -1 and is the
// first item with that Order value.
public string Value1 { get; set; }

// This will be last, because it has the highest Order value.
[Excel(Order = 2)]
public string Value2 { get; set; }

// This will be second-to-last, because it has the second-highest Order value.
[Excel(Order = 1)]
public string Value3 { get; set; }

// This will be first, because it has the lowest Order value.
[Excel(Order = -2)]
public string Value4 { get; set; }

// This will be third, because it has an implicit Order of -1 and is the
// second item with that Order value.
public string Value5 { get; set; }
```


Future work
-----------

- Allow the generated Excel file name to be set as an attribute on objects.
- Extend the range of types that can be serialized to include non-enumerables and nested item types.
- Allow reading from Excel documents.
- Provide finer control over serialisation of individual properties (e.g. to serialise booleans as "yes" or "no").

Have an idea that you don't see here? Can't figure out how to do something? Would like to use this, but it doesn't cover your needs? Go ahead and open an issue!



<!-- References -->

[epplus-codeplex]:
  http://epplus.codeplex.com/

[epplus-nuget]:
  http://www.nuget.org/packages/EPPlus/

[number-format]:
  http://office.microsoft.com/en-001/excel-help/create-a-custom-number-format-HP010342372.aspx

[json-net]:
  http://james.newtonking.com/json/help/index.html?topic=html/JsonPropertyOrder.htm
