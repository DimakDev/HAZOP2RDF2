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
    "errors"
    "fmt"
    "io/ioutil"
    "log"
    "path/filepath"
    "strings"
    "sync"
    "time"

    "github.com/dimakdev/hazop-formula/pkg/exporter"
    "github.com/dimakdev/hazop-formula/pkg/workbook"
    "github.com/manifoldco/promptui"
    "github.com/spf13/cobra"
    "github.com/spf13/viper"
)

var (
    ErrReadingConfig     = errors.New("Error reading config file")
    ErrNoWorksheetsFound = errors.New("Error no worksheets found")
    ErrNoExcelFiles      = errors.New("Error no Excel files found")
    ErrReadingDirecotry  = errors.New("Error reading directory")
    ErrPromptFailed      = errors.New("Error prompt failed")
    InfoDescription      = "Import, parse and verify"
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

    viper.SetConfigName("config")
    viper.SetConfigType("toml")
    viper.AddConfigPath(".")

    if err := viper.ReadInConfig(); err != nil {
        log.Fatalf("%v: %v", ErrReadingConfig, err)
    }

    if err := viper.UnmarshalKey("hazop", &workbook.Hazop); err != nil {
        log.Fatalf("%v: %v", ErrReadingConfig, err)
    }

    if err := viper.UnmarshalKey("program", &program); err != nil {
        log.Fatalf("%v: %v", ErrReadingConfig, err)
    }

    if err := viper.UnmarshalKey("data", &data); err != nil {
        log.Fatalf("%v: %v", ErrReadingConfig, err)
    }
}

type Program struct {
    Author      string `mapstructure:"author"`
    Name        string `mapstructure:"name"`
    Description string `mapstructure:"description"`
    Help        string `mapstructure:"help"`
    Version     string `mapstructure:"version"`
}

type Data struct {
    DataDir             string `mapstructure:"data_dir"`
    DataExt             string `mapstructure:"data_ext"`
    ReportDir           string `mapstructure:"report_dir"`
    ReportExt           string `mapstructure:"report_ext"`
    GraphDir            string `mapstructure:"graph_dir"`
    GraphExt            string `mapstructure:"graph_ext"`
    BaseUri             string `mapstructure:"base_uri"`
    GraphTemplate       string `mapstructure:"graph_template"`
    ReportTemplateLong  string `mapstructure:"report_template_long"`
    ReportTemplateShort string `mapstructure:"report_template_short"`
}

type Command struct {
    Name        string
    Description string
    Datapath    string
}

var data Data
var program Program

func run() error {
    files, err := ioutil.ReadDir(data.DataDir)
    if err != nil {
        return fmt.Errorf("%v `%s` %v", ErrReadingDirecotry, data.DataDir, err)
    }

    var commands []Command
    for _, f := range files {
        if strings.HasSuffix(f.Name(), data.DataExt) {
            name := fmt.Sprintf("`%s`", f.Name())
            datapath := filepath.Join(data.DataDir, f.Name())
            description := fmt.Sprintf("%s `%s`", InfoDescription, f.Name())
            commands = append(commands,
                Command{
                    Name:        name,
                    Datapath:    datapath,
                    Description: description,
                },
            )
        }
    }

    if len(commands) == 0 {
        return fmt.Errorf("%v %s", ErrNoExcelFiles, data.DataDir)
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
        return fmt.Errorf("%v %v", ErrPromptFailed, err)
    }

    wb, err := readWorkbook(commands[i].Datapath)
    if err != nil {
        return err
    }

    if err := writeTemplateOutput(wb); err != nil {
        return err
    }

    return nil
}

func readWorkbook(fpath string) (*workbook.Workbook, error) {
    var wg sync.WaitGroup

    wb, err := workbook.ReadVerifyWorkbook(fpath, &wg)
    if err != nil {
        return nil, err
    }

    wg.Wait()

    if len(wb.Worksheets) == 0 {
        return nil, ErrNoWorksheetsFound
    }

    return wb, nil
}

func writeTemplateOutput(wb *workbook.Workbook) error {
    _, wbname := filepath.Split(wb.File.Path)
    fname := strings.TrimSuffix(wbname, filepath.Ext(wbname))
    rpath := filepath.Join(data.ReportDir, fname+data.ReportExt)
    gpath := filepath.Join(data.GraphDir, fname+data.GraphExt)

    exp := &exporter.Exporter{
        ReportPath:     rpath,
        GraphPath:      gpath,
        ProgramName:    program.Name,
        ProgramVersion: program.Version,
        DateTime:       time.Now().Format(time.UnixDate),
        BaseUri:        data.BaseUri + program.Name,
        Workbook:       wbname,
        Worksheets:     wb.Worksheets,
    }

    if err := exp.WriteToFile(gpath, data.GraphTemplate); err != nil {
        return err
    }

    if err := exp.WriteToFile(rpath, data.ReportTemplateLong); err != nil {
        return err
    }

    if err := exp.WriteToStdout(data.ReportTemplateShort); err != nil {
        return err
    }

    return nil
}
