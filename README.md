# Overview:
This project contains two `Tekla Structures 2025` macros:

## ExportToExcel
Exports member information (ID, Name, Section, Start/End points, Comment) to an Excel file `Desktop\Model.xlsx`.

## ImportFromExcel
Reads connections from an Excel file `Desktop\Model.xlsx` and creates connections between members in Tekla.

## Install Software

- Install [Tekla Structures 2025 - Free Trial](https://download.tekla.com/tekla-structures/free-trial)
- Install [.NET Framework 4.8](https://dotnet.microsoft.com/en-us/download/dotnet-framework/thank-you/net48-web-installer)

## Steps to Use Macros

- Open desrired model e.g. `YourModelName` in `Tekla Structures 2025`
- Create 2 macros `ExportToExcel` and `ImportFromExcel` in Tekla Structures 2025
- Paste [`modeling`](docs/modeling) folder under `docs/modeling` in this repo to `C:\TeklaStructuresModels\<YourModelName>\macros\modeling`


## Run macros
- Go to `Edit` > `Components` > `Applications & Components` > `Search for your macro name`
- Double click `ExportToExcel` or `ImportFromExcel` in `Tekla Structures 2025` to run macros