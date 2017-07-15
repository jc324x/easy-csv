# Easy CSV #

| Property      | Value      | Required | Process |
| ------------- | :--------  | :---:    | :-----: |
| process       | `string`   | yes      | all
| projectPath   | `string`   | yes      | all
| removeHeaders | `boolean`  | no       | all
| target        | `Object`   | yes      | expandSheet
| targets       | `Object[]` | yes      | exportSheets
| zipCSVs       | `boolean`  | no       | all
| zipName       | `string`   | no       | all

## expandSheet ##

```json
{
   "process":"expandSheet",
   "projectPath":"easy-csv-exports/jss-mutt",
   "removeHeaders":true,
   "target":{
      "sheet":"jss-mutt",
      "range":"A:J"
   },
   "zipCSVs":true,
   "zipName":"jss-mutt.zip"
}
```

### Input / Output ###

![alt text][sheet]
[sheet]: https://github.com/jcodesmn/easy-csv/blob/master/images/expandSheet/sheet.png "sheet"

## exportSheets ##

```json
{
   "projectPath":"easy-csv-exports/apple-school-manager",
   "process":"exportSheets",
   "targets":[
      {
         "sheet":"locations"
      },
      {
         "sheet":"students",
         "range":"A:J"
      },
      {
         "sheet":"courses",
         "range":"A:D"
      },
      {
         "sheet":"classes",
         "range":"A:E"
      },
      {
         "sheet":"rosters",
         "range":"A:C"
      },
      {
         "sheet":"staff",
         "range":"A:H"
      }
   ],
   "zipCSVs":true,
   "zipName":"Archive.zip"
}
```

## exportSpreadsheet ##

```json
{
   "projectPath":"easy-csv-exports/spreadsheet",
   "process":"exportSpreadsheet",
   "zipCSVs":true,
   "zipName":"Archive.zip"
}
```

`projectPath` | `string`  
Path to the output folder. Files in the path will be created if they don't already exist.

`removeHeaders` | `boolean`  

`target` | `string`  

`zipCSVs` | `boolean`  

`zipName` | `string`  

