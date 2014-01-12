Excel for Web API
=================

**Currently in alpha.**

An Excel MediaTypeFormatter for ASP.NET Web API.


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

- Incomplete unit test coverage.
- Bad documentation—sorry about that! :)

Future work
-----------

- Allow reading from Excel documents.
