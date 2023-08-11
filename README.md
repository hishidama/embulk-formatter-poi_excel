# Apache POI Excel formatter plugin for Embulk

Formats Excel files(xls, xlsx) for other file output plugins.  
This plugin uses Apache POI.

## Overview

* **Plugin type**: formatter
* Embulk 0.9 or earlier (refer to https://github.com/hishidama/embulk-formatter-excel-poi for 0.10 and later)


## Configuration

* **spread_sheet_version**: Excel file version. `EXCEL97` or `EXCEL2007`. (string, default: `EXCEL2007`)
* **sheet_name**: sheet name. (string, default: `Sheet1`)
* **column_options**: see bellow. (hash, default: `{}`)

### column_options

* **data_format**: data format of Cell. (string, default: `null`)

## Example

```yaml
in:
  type: any input plugin type
...
    columns:
    - {name: time,     type: timestamp}
    - {name: purchase, type: timestamp}

out:
  type: file	# any file output plugin type
  path_prefix: /tmp/embulk-example/excel-out/sample_
  file_ext: xls
  formatter:
    type: poi_excel
    spread_sheet_version: EXCEL97
    sheet_name: Sheet1
    column_options:
      time:     {data_format: "yyyy/mm/dd hh:mm:ss"}
      purchase: {data_format: "yyyy/mm/dd"}
```

### Note

The file name, file split or data order are decided by input/output plugin.  
If you'd like to process data and output Excel format, I think it's also one way to use [Asakusa Framework](http://www.asakusafw.com/) ([Excel Exporter](http://www.ne.jp/asahi/hishidama/home/tech/asakusafw/directio/excelformat.html>)).


## Install

```
$ embulk gem install embulk-formatter-poi_excel
```


## Build

```
$ ./gradlew package
```
