/*
Copyright © 2021 Dmytro Kostiuk <X100@X100.LINK>

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
*/
package cmd

import (
    "fmt"
    "io/ioutil"
    "path/filepath"
    "strings"

    "github.com/dimakdev/hazop-formula/pkg/workbook"
    "github.com/manifoldco/promptui"
    "github.com/spf13/cobra"
)

var promptCmd = &cobra.Command{
    Use:   "prompt",
    Short: "Import, parse and verify Excel documents",
    Long:  "Import, parse and verify Excel documents",
    Run: func(cmd *cobra.Command, args []string) {
        if err := run(); err != nil {
            cmd.PrintErrln(err)
        }
    },
}

func init() {
    rootCmd.AddCommand(promptCmd)
}

type Command struct {
    Name        string
    Description string
    Datapath    string
}

func run() error {
    common := workbook.GetCommon()

    files, err := ioutil.ReadDir(common.DataDir)
    if err != nil {
        return fmt.Errorf("Error reading `%s`: %v", common.DataDir, err)
    }

    var commands []Command
    for _, f := range files {
        if strings.HasSuffix(f.Name(), common.DataExt) {
            datapath := filepath.Join(common.DataDir, f.Name())
            commands = append(commands,
                Command{
                    Name:        fmt.Sprintf("`%s`", f.Name()),
                    Description: fmt.Sprintf("Import, parse and verify `%s`", f.Name()),
                    Datapath:    datapath,
                },
            )
        }
    }

    if len(commands) == 0 {
        return fmt.Errorf("Directory `%s` is empty, no `%s` file(s) found", common.DataDir, common.DataExt)
    }

    templates := &promptui.SelectTemplates{
        Label:    "========== {{ . }} ==========",
        Active:   "⍈ {{ .Name }}",
        Inactive: "  {{ .Name }}",
        Selected: "⍈ {{ .Name }}",
        Details:  "---------- Description -------\n{{ .Description | faint }}",
    }

    searcher := func(input string, index int) bool {
        command := commands[index]
        name := strings.Replace(strings.ToLower(command.Name), " ", "", -1)
        input = strings.Replace(strings.ToLower(input), " ", "", -1)

        return strings.Contains(name, input)
    }

    prompt := promptui.Select{
        Label:     "Commands",
        Items:     commands,
        Templates: templates,
        Size:      8,
        Searcher:  searcher,
    }

    i, _, err := prompt.Run()
    if err != nil {
        return fmt.Errorf("Prompt failed: %v", err)
    }

    _, err = workbook.NewWorkbook(commands[i].Datapath)
    if err != nil {
        return err
    }

    return nil
}
