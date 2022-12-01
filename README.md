# ReStoRunT
ReStoRunT: simple Recording, Storing, Re-Running and Tracing changes in Excel files⋆
<<<<<<< HEAD
=======

## ReStoRunT Python Tools

This folder contains a few simple tools for ReStoRunT instrumenting Excel sheets, as well as analyzing ReStoRunT sheets with the purpose of building Excel Sheet transformers.

As VBA macros are often perceived as security risks, and as such they are prone to getting filtered from emails.


### `ReStoRunTify.py  `
Usage:

`
ReStoRunTify --infile f.xslx --outfile g.xslx
`

This tool reads the excel sheet from the infile and writes a transformation to the outfile. For each sheet in the infile, it adds a ReStoRunT sheet to the outfile. The full file is then written into outfile.

Sheets that are already ReStoRunT sheets will not receive a second ReStoRunT sheet.


### `IsolateReStoRunTsheet.py`

`
IsolateReStoRunTsheet --infile f.xslx  --tobeisolated "TestSheet" --outfile isolated.xslx
`

Take `ReStoRunT-TestSheet` from f.xslx and create a that contains just `ReStoRunT-TestSheet` and an empty `TestSheet`. We need the empty `TestSheet`, as without such a sheet all the references in `ReStoRunT-TestSheet` will be broken, yielding all kind of undesired effects when looking at this from within Excel. So... better keep a dummy sheet.


### `ApplyReStoRunTsheet.py`

`
ApplyReStoRunTsheet --infile f.xslx --sheetfile g.xslx --destinationsheet "Sheet 2" --outfile o.xslx
`
Takes the first ReStoRunT sheet in the sheetfile (`g.xslx`) and applies it to the sheet `--destinationsheet` `Sheet 2`, and then writes out the result to the `--outfile o.xslx`

---
## Contributors
### <ins>Wolfgang Müller</ins>

HITS gGmbH
Heidelberg Institute for Theoretical Studies, *Group Leader* Scientific Databases and Visualization

- Email: Wolfgang.Mueller@h-its.org
- OrcID: [0000-0002-4980-3512](https://orcid.org/0000-0002-4980-3512)
- [Web](https://www.h-its.org/de/people/priv-doz-dr-wolfgang-muller/)


### **Lukrécia Mertová** 
HITS gGmbH
Heidelberg Institute for Theoretical Studies, *PhD Student* Scientific Databases and Visualization

- Email: Lukrecia.Mertova@h-its.org
- OrcID: [0000-0002-7585-4479](https://orcid.org/0000-0002-7585-4479)
- [Web](https://www.h-its.org/people/lukrecia-mertova/)


*\*Supported by the Heidelberg Institute for Theoretical Studies and the Klaus Tschira Foundation, 
as well as MESI-STRAT. MESI-STRAT has received funding from the European Union’s Horizon 2020 research and innovation 
programme under grant agreement No 754688.*

