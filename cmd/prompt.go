/*
Copyright © 2021 Dmytro Kostiuk <dmytro.kostiuk@mailbox.tu-dresden.de>

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
    "time"

    "github.com/dimakdev/HAZOP2RDF2/pkg/exporter"
    "github.com/dimakdev/HAZOP2RDF2/pkg/importer"
    "github.com/manifoldco/promptui"
    "github.com/spf13/cobra"
    "github.com/spf13/viper"
)

var (
    ErrReadingConfig     = errors.New("Error reading manifest file")
    ErrNoWorksheetsFound = errors.New("Error no worksheets found")
    ErrNoHazopFiles      = errors.New("Error no Hazop files found")
    ErrReadingDirecotry  = errors.New("Error reading directory")
    ErrPromptFailed      = errors.New("Error prompt failed")
    CommandDescription   = "Import, parse and verify"
)

var promptCmd = &cobra.Command{
    Use:   "prompt",
    Short: "Import, parse and verify Excel workbooks",
    Long:  "Import, parse and verify Excel workbooks",
    Run: func(cmd *cobra.Command, args []string) {
        if err := run(); err != nil {
            cmd.PrintErrln(err)
        }
    },
}

func init() {
    rootCmd.AddCommand(promptCmd)

    viper.SetConfigName("manifest")
    viper.SetConfigType("toml")
    viper.AddConfigPath(".")

    if err := viper.ReadInConfig(); err != nil {
        log.Fatalf("%v: %v", ErrReadingConfig, err)
    }

    if err := viper.UnmarshalKey("hazop", &importer.Hazop); err != nil {
        log.Fatalf("%v: %v", ErrReadingConfig, err)
    }

    if err := viper.UnmarshalKey("application", &application); err != nil {
        log.Fatalf("%v: %v", ErrReadingConfig, err)
    }

    if err := viper.UnmarshalKey("roots", &roots); err != nil {
        log.Fatalf("%v: %v", ErrReadingConfig, err)
    }
}

type Application struct {
    Author      string `mapstructure:"author"`
    Name        string `mapstructure:"name"`
    Description string `mapstructure:"description"`
    Version     string `mapstructure:"version"`
}

type Roots struct {
    HazopDir            string `mapstructure:"hazop_dir"`
    HazopExt            string `mapstructure:"hazop_ext"`
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
    Datapath    string
    Description string
}

var roots Roots
var application Application

func run() error {
    hazopFiles, err := ioutil.ReadDir(roots.HazopDir)
    if err != nil {
        return fmt.Errorf("%v `%s` %v", ErrReadingDirecotry, roots.HazopDir, err)
    }

    var commands []Command
    for _, f := range hazopFiles {
        if strings.HasSuffix(f.Name(), roots.HazopExt) {
            name := fmt.Sprintf("`%s`", f.Name())
            datapath := filepath.Join(roots.HazopDir, f.Name())
            description := fmt.Sprintf("%s `%s`", CommandDescription, f.Name())
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
        return fmt.Errorf("%v %s", ErrNoHazopFiles, roots.HazopDir)
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

    wb, err := importer.ImportWorkbook(commands[i].Datapath)
    if err != nil {
        return err
    }

    if len(wb.Worksheets) == 0 {
        return ErrNoWorksheetsFound
    }

    _, wbname := filepath.Split(wb.File.Path)
    fname := strings.TrimSuffix(wbname, filepath.Ext(wbname))
    rpath := filepath.Join(roots.ReportDir, fname+roots.ReportExt)
    gpath := filepath.Join(roots.GraphDir, fname+roots.GraphExt)

    e := &exporter.Exporter{
        ReportPath: rpath,
        GraphPath:  gpath,
        AppName:    application.Name,
        AppVersion: application.Version,
        DateTime:   time.Now().Format(time.UnixDate),
        BaseUri:    roots.BaseUri + application.Name,
        Workbook:   wbname,
        Worksheets: wb.Worksheets,
    }

    if err := e.ExportToFile(gpath, roots.GraphTemplate); err != nil {
        return err
    }

    if err := e.ExportToFile(rpath, roots.ReportTemplateLong); err != nil {
        return err
    }

    if err := e.ExportToStdout(roots.ReportTemplateShort); err != nil {
        return err
    }

    return nil
}
