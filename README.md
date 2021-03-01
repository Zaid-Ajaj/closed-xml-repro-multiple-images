To reproduce the issue, run the console application
```
dotnet restore
dotnet run
```
You will get two Excel files, one called `SingleImage.xlsx` that contains on image fixed into the cell (expected behaviour) and another Excel file called `MultipleImages.xlsx` that seems to be corrupted