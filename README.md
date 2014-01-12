Excel for Web API
=================

**Currently in alpha.**

An Excel MediaTypeFormatter for ASP.NET Web API. Currently requires that classes be decorated with `DataMember` attributes to determine serialisation, though this will be changed in soon-to-come updates.


Basic syntax
------------

```C#
config.Formatters.Add(new ExcelMediaTypeFormatter()); // Where config = System.Web.Http.HttpConfiguration.
```

Advanced options
----------------

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


Notable issues
--------------

- Can only serialize classes decorated with `DataContract` and `DataMember` attributes.
  - This is a temporary workaround to allow ordering of columns that will soon be dispatched with.
- Incomplete unit test coverage.
- Bad documentation—sorry about that! :)

Future work
-----------

- Remove yucky dependency on `DataContract` attributes.
- Allow reading from Excel documents.
