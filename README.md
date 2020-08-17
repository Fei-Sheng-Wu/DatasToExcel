# DatasToExcel

[![Target Framework](https://img.shields.io/badge/%2ENet%20Core-3.0-green.svg)](https://docs.microsoft.com/en-us/dotnet/core/about)
[![Nuget](https://img.shields.io/badge/Nuget-v1.0.0-blue.svg)](https://www.nuget.org/packages/DatasToExcel/1.0.0)
[![Lincense](https://img.shields.io/badge/Lincense-MIT-orange.svg)](https://github.com/Fei-Sheng-Wu/DatasToExcel/blob/1.0.0/LICENSE.txt)

> A 2D array to Excel file converter, support header first row and fixed headers. It uses .Net Core 3.0 as framework and only depends on the Open Xml SDK.

## Dependencies

**.Net Core** >= 3.0  
**DocumentFormat.OpenXml** = 2.10.1

## Main Features

- [x] 2D array to Excel
- [x] Header first row (optional)
- [x] Fixed first row (optional)

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

And convert it to Excel and save to a fixed path.

```c#
datas.GenerateExcel(filename); //Do not header first row
datas.GenerateExcel(filename, true); //Header first row
```

Or save it to a memory stream.

```c#
//Option 1
MemoryStream ms = datas.GenerateExcel(); //Do not header first row
MemoryStream ms = datas.GenerateExcel(true); //Header first row

//Option 2
using (MemoryStream ms = new MemoryStream())
{
    datas.GenerateExcel(ms); //Do not header first row
    datas.GenerateExcel(ms, true); //Header first row
}
```

## License

This project is under the [MIT License](https://github.com/Fei-Sheng-Wu/DatasToExcel/blob/1.0.0/LICENSE.txt).
