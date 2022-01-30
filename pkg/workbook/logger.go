package workbook

import (
    "errors"
)

var ErrUnknownPrefix = errors.New("Error unknown prefix")

type Logger struct {
    Warnings []string
    Errors   []string
    Info     []string
}

func (l *Logger) newMessage(msg string) error {
    switch msg[:4] {
    case "Erro":
        l.newError(msg)
    case "Info":
        l.newInfo(msg)
    case "Warn":
        l.newWarning(msg)
    default:
        return ErrUnknownPrefix
    }
    return nil
}

func (l *Logger) newWarning(msg string) {
    l.Warnings = append(l.Warnings, msg)
}

func (l *Logger) newError(msg string) {
    l.Errors = append(l.Errors, msg)
}

func (l *Logger) newInfo(msg string) {
    l.Info = append(l.Info, msg)
}
