package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"strings"

	"github.com/manifoldco/promptui"
)

const (
	data    = "./data"
	ext     = ".xlsx"
	rows    = 5
	format  = ".ttl"
	logfile = "verification.log"
)

type Command struct {
	Name  string
	Alias string
	Desc  string
}

type Prompt struct {
	PromptUI promptui.Select
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
	files, err := ioutil.ReadDir(data)
	if err != nil {
		return fmt.Errorf("Error reading `%s`: %v", data, err)
	}

	var commands []Command
	for _, f := range files {
		if strings.HasSuffix(f.Name(), ext) {
			commands = append(commands,
				Command{
					Name:  fmt.Sprintf("Import `%s`", f.Name()),
					Alias: "IMPORT",
					Desc:  fmt.Sprintf("Import `%s` from `%s` directory.", f.Name(), data),
				},
				Command{
					Name:  fmt.Sprintf("Show `%s`", f.Name()),
					Alias: "HEAD",
					Desc:  fmt.Sprintf("Show first %d rows of `%s`.", rows, f.Name()),
				},
				Command{
					Name:  fmt.Sprintf("Show `%s`", f.Name()),
					Alias: "TAIL",
					Desc:  fmt.Sprintf("Show last %d rows of `%s`.", rows, f.Name()),
				},
				Command{
					Name:  fmt.Sprintf("Show `%s`", f.Name()),
					Alias: "FULL",
					Desc:  fmt.Sprintf("Show all rows of `%s`.", f.Name()),
				},
				Command{
					Name:  fmt.Sprintf("Verificate `%s`", f.Name()),
					Alias: "INPUT",
					Desc:  fmt.Sprintf("Verificate input `%s` data:\n  1. Check column name (regex, synonyms).\n  2. Check column type (number, text).\nVerification log is saved in `%s`.", f.Name(), logfile),
				},
				Command{
					Name:  fmt.Sprintf("Verificate `%s`", f.Name()),
					Alias: "DATA",
					Desc:  fmt.Sprintf("Verificate `%s` data:\n  1. Check value length.\n  2. Check value range.\n  3. Check optional and mandatory fields.\nVerification log is saved in `%s`.", f.Name(), logfile),
				},
				Command{
					Name:  fmt.Sprintf("Generate report graph of `%s`", f.Name()),
					Alias: "GRAPH",
					Desc:  fmt.Sprintf("Generate report graph of `%s` in `%s`.", f.Name(), format),
				},
			)
		}
	}

	if len(commands) == 0 {
		return fmt.Errorf("`%s` is empty, no `%s` file(s) found", data, ext)
	}

	templates := &promptui.SelectTemplates{
		Label:    "========== {{ . }} ==========",
		Active:   "🎱{{ .Name | cyan }} ({{ .Alias | red }})",
		Inactive: "  {{ .Name | cyan }} ({{ .Alias | red }})",
		Selected: "🎱{{ .Name | red | cyan }}",
		Details: `
---------- Description ---------------------------
{{ .Desc | faint }}`,
	}

	searcher := func(input string, index int) bool {
		command := commands[index]
		name := strings.Replace(strings.ToLower(command.Name), " ", "", -1)
		input = strings.Replace(strings.ToLower(input), " ", "", -1)

		return strings.Contains(name, input)
	}

	p.PromptUI = promptui.Select{
		Label:     "Select Hazop Formula Command",
		Items:     commands,
		Templates: templates,
		Size:      8,
		Searcher:  searcher,
	}

	return nil
}

func (p *Prompt) run() error {
	_, _, err := p.PromptUI.Run()
	if err != nil {
		return fmt.Errorf("Prompt failed: %v", err)
	}
	return nil
	// fmt.Printf("You choose number %d: %s\n", i+1, commands[i].Name)
}
