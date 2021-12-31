[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)   


## Hazop ⚙️⚙️⚙️ Formula [Student Thesis]

Tool to parse and verify Hazop studies. It works with linear data *excel* worksheets. Program interface can be extended to adopt other linear and hierarchical formats e. g. *csv*, *json*, *yaml*, *xml*, *rdf*.

The tool can be used to unify Hazop knowledge and create a unified ontology so that the data can be effectively stored and queried in a database.


### Works in shell

Use *bash*, *zsh*, *fish*, or shell of your choice.

**Requirements** 💀

- with asdf: `asdf install` (the specified go version will be installed [.tool-versions](.tool-versions)) 
- Go 1.17 

**Run**

- install program: `go install`
- help screen: `hazop-formula`
- run prompt: `hazop-formula prompt`


### How it works

The [settings](cfg.toml) specify all the constants and verification constraints for the data, e. g. name of the Hazop elements, data type, cell type, length, and range of the data. Documents in [data](data) are imported, parsed, and verified. The results are stored in [report](report). 

The report contains header and data information for each subject and peace of the information. The report is used to correct the input data. In addition, it aims to archive the better quality of the incoming information to avoid trivial errors that can lead to the loss of information.


### License

Available under [MIT License](LICENSE).