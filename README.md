# DatasToExcel

[![Target Framework](https://img.shields.io/badge/%2ENet%20Core-3.0-green.svg?style=flat-square)](https://docs.microsoft.com/en-us/dotnet/core/about)
[![Nuget](https://img.shields.io/badge/Nuget-v1.0.1-blue.svg?style=flat-square)](https://www.nuget.org/packages/DatasToExcel/1.0.1)
[![Lincense](https://img.shields.io/badge/Lincense-MIT-orange.svg?style=flat-square)](https://github.com/Fei-Sheng-Wu/DatasToExcel/blob/1.0.1/LICENSE.txt)

> A 2D array to Excel file converter. Generate Microsoft Excel file from a 2D array in C#. Support making the first row headers and custom worksheet name. .Net Core 3.0 framework and depends on the Open Xml SDK.

## Dependencies

**.Net Core** >= 3.0  
**DocumentFormat.OpenXml** = 2.10.1

## Main Features

- [x] 2D array to Excel
- [x] Making the first row headers

## How to Use

Create a 2D array first.

```c#
string[,] datas = new string[,]
{
    { "Name", "Country", "Age", "Career" },
    { "Helen", "U.S.", "21", "Police" },
    { "Jucia", "Canada", "34", "Dancer" },
    { "Erik", "Canada", "13", "Student" },
    { "Bob", "British", "26", "Business person" },
    { "Nancy", "Russia", "64", "Fisherman" },
};
```

Convert it to a Excel file and save to a local file.

```c#
datas.GenerateExcel(filename); //First row as non-headers
datas.GenerateExcel(filename, true); //First row as headers
```

Or save it to a MemoryStream.

```c#
//Option 1
MemoryStream ms = datas.GenerateExcel(); //First row as non-headers
MemoryStream ms = datas.GenerateExcel(true); //First row as headers

//Option 2
using (MemoryStream ms = new MemoryStream())
{
    datas.GenerateExcel(ms); //First row as non-headers
    datas.GenerateExcel(ms, true); //First row as headers
}
```

## License

This project is under the [MIT License](https://github.com/Fei-Sheng-Wu/DatasToExcel/blob/1.0.1/LICENSE.txt).
