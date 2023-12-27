# PowerShell Chart Module

## Abstract
The PSChart is a platform-dependent PowerShell module created in early 2013 as part of reporting tools for monitoring the on-premise SharePoint farms. The module requires the Windows operating system and Windows PowerShell. The code of the module is frozen.

![PSChart](assets\PSChart.png)

## Overview
This PowerShell module provides a convenient way to create charts based on the `System.Windows.Forms.DataVisualization.Charting` library. It offers a range of features for creating and customizing charts, including support for multiple data series, interactive formatting, saving charts in various formats, exporting to the clipboard, and more.

The PSChart module leverages the .NET Charting library, so it supports all the ChartTypes that the Charting library supports. Please refer to the .NET Charting library documentation for a complete list of ChartTypes and details.

Here are some of the common ChartTypes:
- `Area`
- `Bar`
- `Column`
- `Doughnut`
- `FastLine`
- `FastPoint`
- `Line`
- `Pie`
- `Point`

The list of ValueTypes:
- `Double`: A double-precision floating-point number.
- `Single`: A single-precision floating-point number.
- `Int32`: A 32-bit signed integer.
- `Int64`: A 64-bit signed integer.
- `UInt32`: A 32-bit unsigned integer.
- `UInt64`: A 64-bit unsigned integer.
- `DateTime`: A date and time value.
- `Date`: A date value.
- `Time`: A time value.
- `String`: A string value.

Since the PSChart module is based on the .NET Charting library, it supports a wide range of formatting options provided by the library. Please refer to the .NET Charting library documentation for a complete list and details of each formatting option.

## Features

- Creation of Charts: The module allows for the creation of various types of charts such as line, bar, pie, etc. using the New-Chart function. 

- Data Series: It supports the creation of one or more data series using the New-ChartSeries function.

- Interactive Post Formatting: The module allows for interactive post formatting of series, charts, and chart areas.

- Saving and Exporting: Charts can be saved in .png and .emf formats, and can also be exported to the clipboard.

- Customizable ChartType and ValueType: The module allows for the selection of ChartType and ValueType.

- Documentation: The module includes a detailed readme.md file that provides usage examples and instructions for the module.

- ISE Snippets Support: The module supports ISE Snippets, which can be useful for creating and managing charts in the Integrated Scripting Environment (ISE).


## Requirements

- Operating System Windows 7 or higher
- Windows PowerShell 4 or higher
- Recommended for use with ISE (Integrated Scripting Environment)

## Installation
1. Place the module files in the following directory: 

```
$home\Documents\WindowsPowerShell\Modules\PSChart
```
2. Import the module by running the following command: 
```
Import-Module PSChart
```
3. Optionally, import the ISE snippet by running: 
```
Import-IseSnippet -Module PSChart
```

## Documentation

For detailed documentation run the following command:
```
Get-Help New-Chart
```

The parameters of the `New-Chart` cmdlet:

- DataSource – data source (collection of objects) for building the charts
- Series – series of the chart (several series can be placed on one chart) 
- SaveAs <FileName> - save to file (by default is using format png)
- Silent – do not show the chart dialog box
- Clipboard – place a copy of the image on the clipboard
- Enable3 – 3D graph.
- ChartTitle - Chart title 
- YTitle, XTitle – names of the axes

## Usage

There are only two functions needed to create a chart: 
- New-ChartSeries
- New-Chart

To get started, refer to the examples and documentation provided below.

Import the module and optionally import the ISE snippet:

```powershell
Import-Module PSChart
Import-IseSnippet -Module PSChart
```

Prepare the data source for the chart. 

Example: The data source is a collection of files and folders in the user's Documents directory. The first 10 items are selected and the `Name`, `FullName`, `Length`, `CreationTime`, `LastAccessTime`, `LastWriteTime`, and `Extension` properties are selected as the value members for the chart.

```powershell
$dataSource = Get-ChildItem -Recurse "$home\Documents\" |
    Where-Object { $_.PSIsContainer -ne $true } | 
    Select-Object -First 10 | 
    Select-Object name, fullname, Length, CreationTime, LastAccessTime, LastWriteTime, Extension
```

### Example - Creating a chart with multiple series

The code below demonstrates the usage of the `New-Chart` cmdlet to create pie charts.

```powershell
New-Chart -dataSource $dataSource -ValueMembers Name, Length -ChartType pie
New-Chart -dataSource $dataSource -ValueMembers Name, Length -ChartType pie -Enable3D
```

The example of creating a chart where the data passes through the pipeline.

```powershell
$dataSource | Group-Object Extension | ForEach-Object {
    New-Object psobject -Property @{
        Extension = $_.Name
        Size      = ($_.Group | Measure-Object Length -Sum).Sum
        Count     = ($_.Group | Measure-Object Length -Sum).Count
    }
} | Sort-Object Size -Descending | New-Chart -ChartType Bar -ValueMembers Extension,Size 

```

### Example - Creating a chart with selected data and specified value members

This code demonstrates the usage of the `New-Chart` cmdlet to create a bar chart. The chart is populated with data from the first 10 processes retrieved using the `Get-Process` cmdlet. The `Name` and `VirtualMemorySize` properties are used as the value members for the chart.

Explanation: selected data, specified the ValueMembers parameter (or its alias Properties), which determines which properties to use for the X, and Y axes. Chart type – optional (If ChartType is specified before the parameters containing the “,” sign), we will get the intelisence hint). 

```powershell
$datasource = Get-Process | Select-Object -First 6
New-Chart -DataSource $datasource -Properties Name,VirtualMemorySize
```
The similar examples, but with the use pipeline.

```powershell
Get-Process | New-Chart -ChartType spline -Properties Name,VirtualMemorySize
Get-Process | New-Chart -ChartType spline -ValueMembers Name,VirtualMemorySize
```

The complete example with the chart.

```powershell
New-Chart -ChartType Bar `
-dataSource (Get-Process | Select-Object -First 6) `
-ValueMembers Name, VirtualMemorySize
```
![Series1](assets\PSChartSeries1.png)


### Example - Creating a chart with multiple series

Creating a chart with multiple series. The source is data from the Get-Process cmdlet. 

```powershell
$datasource = Get-Process |select Name,VirtualMemorySize,WorkingSet,PeakWorkingSet | Select-Object -First 6
$Series1 = New-ChartSeries -XValueMember Name -YValueMembers VirtualMemorySize -ChartType Column
$Series2 = New-ChartSeries -XValueMember Name -YValueMembers WorkingSet -ChartType Column
$Series3 = New-ChartSeries -XValueMember Name -YValueMembers PeakWorkingSet -ChartType Spline
New-Chart -datasource $datasource -series $Series1,$Series2,$Series3
```

The code creates a chart with three series representing CPU usage data for different servers over time. The chart is saved as an Enhanced Metafile (EMF) image file.

```powershell
$dataSource = Import-Csv -Path C:\temp\data.csv
$Series1 = New-ChartSeries -XValueMember 'DayOfWeek' -YValueMembers 'Server1' -ChartType Spline
$Series2 = New-ChartSeries -XValueMember 'DayOfWeek' -YValueMembers 'Server2' -ChartType Spline
$Series3 = New-ChartSeries -XValueMember 'DayOfWeek' -YValueMembers 'Server3' -ChartType Spline
New-Chart -dataSource $dataSource -series $Series1, $Series2, $Series3 -SaveAs C:\temp\data.emf  -YTitle 'CPU Usage, %' -XTitle 'Time' -ChartTitle 'Average CPU Usage'
```

![Series2](assets\PSChartSeries2.png)


### Example - Creating a chart with custom data source

The code creates a pie chart by defining a data source. Each data point represents a category with a corresponding value. The chart series is then configured to use the `free` property as the X-axis value and the `used` property as the Y-axis value. Finally, the chart is created using the specified data source and series.

```powershell
$dataSource = New-Object System.Collections.ArrayList;
[void]$dataSource.Add([PSCustomObject]@{ free = 'foo'; used = 70; });
[void]$dataSource.Add([PSCustomObject]@{ free = 'zoo'; used = 30; });

$Series1 = New-ChartSeries -XValueMember 'free' -YValueMembers 'used' -ChartType Pie
New-Chart -dataSource $dataSource -series $Series1
```

### Example - Creating a chart with New-ChartSeries

The principle of using New-ChartSeries. 
New-ChartSeries – creates a series (several of them can be placed on one chart)

```powershell
$datasource = Get-Process | Select -First 3
$Series1 = New-ChartSeries -XValueMember Name -YValueMembers PeakWorkingSet
New-Chart -datasource $datasource -series $Series1
```

### Example - Saving a chart to a file

The chart is saved as a Portable Network Graphics (PNG) image file. The code will be run in silent mode (without showing the chart dialog box).

```powershell
New-Chart -datasource $datasource -series $Series1 -SaveAs C:\temp\abc.png -Silent
```
