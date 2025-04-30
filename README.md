# Overview:
This project contains two `Tekla Structures 2025` macros:

## ExportToExcel
Exports member information (ID, Name, Section, Start/End points, Comment) to an Excel file `Desktop\Model.xlsx`.
See the [Model.xlsx](#Modelxlsx) section for details on the exported data format.

## ImportFromExcel
Reads connections from an Excel file `Desktop\Connections.xlsx` and creates connections between members in Tekla.
See the [Connections.xlsx](#Connectionsxlsx) section for details on the exported data format.

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

## Excel Format

### Model.xlsx

| ID     | Name | Profile | StartX | StartY | StartZ | EndX | EndY | EndZ | Comment | Selected |
| ------ | ---- | ------- | ------ | ------ | ------ | ---- | ---- | ---- | ------- | -------- |
| GUID-1 | B1   | W200X42 | 0.0    | 0.0    | 0.0    | 4000 | 0.0  | 0.0  | Column  |          |
| GUID-2 | B2   | W250X53 | 0.0    | 0.0    | 3000   | 4000 | 0.0  | 3000 | Beam    |          |

---

#### Column Descriptions

- **MemberGUID**: Unique identifier of the Tekla part (used for connection mapping later).
- **Name**: Part name (`.Name` property).
- **Profile**: Cross-section or shape (e.g., W200X42).
- **StartX, StartY, StartZ**: Start point coordinates of the part.
- **EndX, EndY, EndZ**: End point coordinates of the part.
- **Comment**: Any custom comment from a UDA (optional).
- **Selected**: *(Initially blank)* – to be filled by the user in Excel later. If filled (e.g., "Yes"), the part is considered selected for connection creation in Macro 2.

### Connections.xlsx

| ID1    | ID2    | ConnectionType | AttributeFile | Selected |
| ------ | ------ | -------------- | ------------- | -------- |
| GUID-1 | GUID-2 | EndPlate       | MyEndPlate    | yes      |
| GUID-3 | GUID-4 | BasePlate      |               | yes      |
| ...    | ...    | ...            | ...           |          |

Only rows with a non-empty `Selected` field (e.g. `Yes`, `yes` or `X`) will be processed.  
