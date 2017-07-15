## Configuring `Config` ##

### Three Different Processes ###

#### expandSheet ####

```json
{
   "projectPath":"easy-csv-exports/jss-mutt",
   "process":"expandSheet",
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
