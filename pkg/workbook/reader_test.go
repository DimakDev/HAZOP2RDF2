package workbook

import (
    "testing"

    "github.com/stretchr/testify/assert"
)

func TestReadXCellNames(t *testing.T) {
    assert := assert.New(t)

    cnames, err := readXCellNames(1, 1, 5)
    assert.Empty(err)
    assert.Equal(cnames, []string{"A1", "B1", "C1", "D1", "E1"})

    cnames, err = readXCellNames(0, 0, 0)
    assert.Empty(err)
    assert.Empty(cnames)

    cnames, err = readXCellNames(-1, 1, 1)
    assert.Error(err)
    assert.Empty(cnames)
}

func TestReadYCellNames(t *testing.T) {
    assert := assert.New(t)

    cnames, err := readYCellNames(1, 1, 5)
    assert.Empty(err)
    assert.Equal(cnames, []string{"A1", "A2", "A3", "A4", "A5"})

    cnames, err = readYCellNames(0, 0, 0)
    assert.Empty(err)
    assert.Empty(cnames)

    cnames, err = readYCellNames(-1, 1, 1)
    assert.Error(err)
    assert.Empty(cnames)
}

func TestReadXCoordinate(t *testing.T) {
    assert := assert.New(t)

    cnames, err := readXCoordinate("A2")
    assert.Empty(err)
    assert.Equal(cnames, 1)

    cnames, err = readXCoordinate("")
    assert.Error(err)
    assert.Equal(cnames, 0)
}

func TestReadYCoordinate(t *testing.T) {
    assert := assert.New(t)

    cnames, err := readYCoordinate("A2")
    assert.Empty(err)
    assert.Equal(cnames, 2)

    cnames, err = readYCoordinate("")
    assert.Error(err)
    assert.Equal(cnames, 0)
}

func TestReadCellNames(t *testing.T) {
    assert := assert.New(t)

    readerX := &reader{
        varDimension: readXCoordinate,
        fixDimension: readYCoordinate,
        cellNames:    readXCellNames,
    }

    readerY := &reader{
        varDimension: readYCoordinate,
        fixDimension: readXCoordinate,
        cellNames:    readYCellNames,
    }

    cnames, err := readerX.readCellNames("A1", 5)
    assert.Empty(err)
    assert.Equal(cnames, []string{"A1", "B1", "C1", "D1", "E1"})

    cnames, err = readerX.readCellNames("A1", 0)
    assert.Empty(err)
    assert.Empty(cnames)

    cnames, err = readerX.readCellNames("", 0)
    assert.Error(err)
    assert.Empty(cnames)

    cnames, err = readerY.readCellNames("A1", 5)
    assert.Empty(err)
    assert.Equal(cnames, []string{"A1", "A2", "A3", "A4", "A5"})

    cnames, err = readerY.readCellNames("A1", 0)
    assert.Empty(err)
    assert.Empty(cnames)

    cnames, err = readerY.readCellNames("", 0)
    assert.Error(err)
    assert.Empty(cnames)
}
