# Easy CSV #

| Property      | Value                                                       | Process |
| ------------- | :------------------------------------------------------     | :-----: |
| process       | `string`  -> expandSheet / exportSheets / exportSpreadsheet | all     |
| projectPath   | `string`  -> path to output folder                          | all     |
| removeHeaders | `boolean` ->                                                | $1      |

#### expandSheet ####

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


#### exportSheets ####

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

#### exportSpreadsheet ####

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

