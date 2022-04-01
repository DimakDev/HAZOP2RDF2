[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)   

#### HAZOP2RDF2 ⚙️⚙️⚙️ [Thesis].

Hazop parser and modeling tool.

#### Works in shell

Use *bash*, *zsh*, *fish*, or shell of your choice.

#### Requirements 💀

- Go 1.17 (with asdf: `asdf install`) 

#### Run

- install: `go install .`
- help: `HAZOP2RDF2`
- prompt: `HAZOP2RDF2 prompt`

Run prompt and choose a Hazop document from [hazop dir](hazop) to proceed. The result is an RDF graph in `turtle` format saved in [graph dir](graph). See log information in the [report dir](report). 

[MIT License](LICENSE).