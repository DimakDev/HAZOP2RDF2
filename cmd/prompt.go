package main

import (
    "fmt"
    "io/ioutil"
    "log"
    "path/filepath"
    "strings"

    "github.com/manifoldco/promptui"
)

type Alias string

const (
    CMD_HEAD  Alias = "CMD_HEAD"
    CMD_TAIL        = "CMD_TAIL"
    CMD_FULL        = "CMD_FULL"
    CMD_INPUT       = "CMD_INPUT"
    CMD_DATA        = "CMD_DATA"
    CMD_GRAPH       = "CMD_GRAPH"
)

type Command struct {
    Name  string
    Alias Alias
    Desc  string
    File  string
}

type Prompt struct {
    PromptUI promptui.Select
    Commands []Command
}

func main() {
    prompt := &Prompt{}
    if err := prompt.generate(); err != nil {
        log.Fatal(err)
    }
    if err := prompt.run(); err != nil {
        log.Fatal(err)
    }
}

func (p *Prompt) generate() error {
    files, err := ioutil.ReadDir(datadir)
    if err != nil {
        return fmt.Errorf("Error reading `%s`: %v", datadir, err)
    }

    var commands []Command
    for _, f := range files {
        if strings.HasSuffix(f.Name(), dataext) {
            datapath := filepath.Join(datadir, f.Name())
            commands = append(commands,
                Command{
                    Name:  fmt.Sprintf("Head `%s`", f.Name()),
                    Alias: CMD_HEAD,
                    Desc:  fmt.Sprintf("Show first rows `%s`", f.Name()),
                    File:  datapath,
                },
                Command{
                    Name:  fmt.Sprintf("Tail `%s`", f.Name()),
                    Alias: CMD_TAIL,
                    Desc:  fmt.Sprintf("Show last rows `%s`", f.Name()),
                    File:  datapath,
                },
                Command{
                    Name:  fmt.Sprintf("Show full `%s`", f.Name()),
                    Alias: CMD_FULL,
                    Desc:  fmt.Sprintf("Show all rows `%s`", f.Name()),
                    File:  datapath,
                },
                Command{
                    Name:  fmt.Sprintf("Verificate input `%s`", f.Name()),
                    Alias: CMD_INPUT,
                    Desc:  fmt.Sprintf("Verificate input `%s`:\n  1. Check column name (regex, synonyms).\n  2. Check column type (number, text).\nVerification results will be saved in `%s`.", f.Name(), logpath),
                    File:  datapath,
                },
                Command{
                    Name:  fmt.Sprintf("Verificate data `%s`", f.Name()),
                    Alias: CMD_DATA,
                    Desc:  fmt.Sprintf("Verificate data `%s`:\n  1. Check value length.\n  2. Check value range.\n  3. Check optional and mandatory fields.\nVerification results will be written in `%s`.", f.Name(), logpath),
                    File:  datapath,
                },
                Command{
                    Name:  fmt.Sprintf("Create Hazop graph `%s`", f.Name()),
                    Alias: CMD_GRAPH,
                    Desc:  fmt.Sprintf("Create RDF graph from source `%s`", f.Name()),
                    File:  datapath,
                },
            )
        }
    }

    if len(commands) == 0 {
        return fmt.Errorf("`%s` is empty, no `%s` file(s) found", datadir, dataext)
    }

    templates := &promptui.SelectTemplates{
        Label:    "========== {{ . }} ==========",
        Active:   "➡️ {{ .Name | cyan }}",
        Inactive: "  {{ .Name | cyan }}",
        Selected: "➡️ {{ .Name | red }}",
        Details:  "---------- Description -------------\n{{ .Desc | faint }}",
    }

    searcher := func(input string, index int) bool {
        command := commands[index]
        name := strings.Replace(strings.ToLower(command.Name), " ", "", -1)
        input = strings.Replace(strings.ToLower(input), " ", "", -1)

        return strings.Contains(name, input)
    }

    p.PromptUI = promptui.Select{
        Label:     "Select Command",
        Items:     commands,
        Templates: templates,
        Size:      8,
        Searcher:  searcher,
    }

    p.Commands = commands

    return nil
}

func (p *Prompt) run() error {
    for {
        i, _, err := p.PromptUI.Run()
        if err != nil {
            return fmt.Errorf("Prompt failed: %v", err)
        }

        switch p.Commands[i].Alias {
        case CMD_HEAD:
            log.Println(CMD_HEAD)
        case CMD_TAIL:
            log.Println(CMD_TAIL)
        case CMD_FULL:
            log.Println(CMD_FULL)
        case CMD_INPUT:
            log.Println(CMD_INPUT)
        case CMD_DATA:
            log.Println(CMD_DATA)
        case CMD_GRAPH:
            log.Println(CMD_GRAPH)
        default:
            return fmt.Errorf("Error alias: `%s` not found", p.Commands[i].Alias)
        }
    }
    return nil
}
