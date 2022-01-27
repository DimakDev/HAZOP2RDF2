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

    "github.com/dimakdev/hazop-formula/pkg/report"
    "github.com/dimakdev/hazop-formula/pkg/workbook"
    "github.com/manifoldco/promptui"
    "github.com/spf13/cobra"
    "github.com/spf13/viper"
)

var (
    ErrReadingSettings        = errors.New("Error reading `settings.toml`")
    ErrNoValidWorksheetsFound = errors.New("Error no valid worksheets found")
    ErrDataDirectoryIsEmpty   = errors.New("Error data directory is empty, no Excel files found")
    ErrReadingDirecotry       = errors.New("Error reading directory")
    ErrPromptFailed           = errors.New("Error prompt failed")
    InfoDescription           = "Import, parse and verify"
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

    viper.SetConfigName("settings")
    viper.SetConfigType("toml")
    viper.AddConfigPath(".")

    if err := viper.ReadInConfig(); err != nil {
        log.Fatalf("%v: %v", ErrReadingSettings, err)
    }

    if err := viper.UnmarshalKey("hazop", &workbook.Hazop); err != nil {
        log.Fatalf("%v: %v", ErrReadingSettings, err)
    }

    if err := viper.UnmarshalKey("program", &program); err != nil {
        log.Fatalf("%v: %v", ErrReadingSettings, err)
    }

    if err := viper.UnmarshalKey("common", &common); err != nil {
        log.Fatalf("%v: %v", ErrReadingSettings, err)
    }
}

type Program struct {
    Name        string `mapstructure:"name"`
    Description string `mapstructure:"description"`
    Help        string `mapstructure:"help"`
    Version     string `mapstructure:"version"`
    Author      string `mapstructure:"author"`
}

type Common struct {
    DataDir        string `mapstructure:"data_dir"`
    DataExt        string `mapstructure:"data_ext"`
    ReportDir      string `mapstructure:"report_dir"`
    ReportExt      string `mapstructure:"report_ext"`
    TemplateFile   string `mapstructure:"template_file"`
    TemplateStdout string `mapstructure:"template_stdout"`
}

type Command struct {
    Name        string
    Description string
    Datapath    string
}

var common Common
var program Program

func run() error {
    files, err := ioutil.ReadDir(common.DataDir)
    if err != nil {
        return fmt.Errorf("%v `%s` %v",
            ErrReadingDirecotry,
            common.DataDir,
            err,
        )
    }

    var commands []Command
    for _, f := range files {
        if strings.HasSuffix(f.Name(), common.DataExt) {
            datapath := filepath.Join(common.DataDir, f.Name())
            commands = append(commands,
                Command{
                    Name: fmt.Sprintf("`%s`", f.Name()),
                    Description: fmt.Sprintf("%s `%s`",
                        InfoDescription,
                        f.Name(),
                    ),
                    Datapath: datapath,
                },
            )
        }
    }

    if len(commands) == 0 {
        return fmt.Errorf("%v %s %s",
            ErrDataDirectoryIsEmpty,
            common.DataDir,
            common.DataExt,
        )
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

    wb, err := runWorkbookRoutine(commands[i].Datapath)
    if err != nil {
        return err
    }

    if err := generateReport(wb); err != nil {
        return err
    }

    return nil
}

func runWorkbookRoutine(fpath string) (*workbook.Workbook, error) {
    var wg sync.WaitGroup

    wb, err := workbook.ReadVerifyWorkbook(fpath, &wg)
    if err != nil {
        return nil, err
    }

    wg.Wait()

    if len(wb.Worksheets) == 0 {
        return nil, ErrNoValidWorksheetsFound
    }

    return wb, nil
}

func generateReport(wb *workbook.Workbook) error {
    _, fname := filepath.Split(wb.File.Path)
    wbname := strings.TrimSuffix(fname, filepath.Ext(fname))
    rpath := filepath.Join(common.ReportDir, wbname+common.ReportExt)

    r := &report.Report{
        ReportPath:     rpath,
        ProgramName:    program.Name,
        ProgramVersion: program.Version,
        DateTime:       time.Now().Format(time.UnixDate),
        Workbook:       wbname,
        Worksheets:     wb.Worksheets,
    }

    if err := r.ReportToFile(rpath, common.TemplateFile); err != nil {
        return err
    }

    if err := r.ReportToStdout(common.TemplateStdout); err != nil {
        return err
    }

    return nil
}
