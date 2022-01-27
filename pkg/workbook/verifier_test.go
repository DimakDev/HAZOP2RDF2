package workbook

import (
    "testing"

    "github.com/stretchr/testify/assert"
)

func TestNewCellVerifier(t *testing.T) {
    assert := assert.New(t)

    var (
        err error
        ver verifier
    )

    ver, err = newCellVerifier(Hazop.CellType.String)
    assert.Empty(err)
    assert.Exactly(ver, verifierString{})

    ver, err = newCellVerifier(Hazop.CellType.Integer)
    assert.Empty(err)
    assert.Exactly(ver, verifierInteger{})

    ver, err = newCellVerifier(Hazop.CellType.Float)
    assert.Empty(err)
    assert.Exactly(ver, verifierFloat{})

    ver, err = newCellVerifier(4)
    assert.Error(err)
    assert.Empty(ver)
}

func TestVerifierString(t *testing.T) {
    assert := assert.New(t)

    var (
        err error
        val interface{}
        ver verifierString
    )

    val, err = ver.checkCellType("")
    assert.Empty(err)
    assert.Empty(val)

    val, err = ver.checkCellType("text")
    assert.Empty(err)
    assert.Equal(val, "text")

    err = ver.checkCellLength("text", 0, 4)
    assert.Empty(err)

    err = ver.checkCellLength("text", 0, 0)
    assert.Error(err)
}

func TestVerifierInteger(t *testing.T) {
    assert := assert.New(t)

    var (
        err error
        val interface{}
        ver verifierInteger
    )

    val, err = ver.checkCellType("text")
    assert.Error(err)
    assert.Empty(val)

    val, err = ver.checkCellType("0")
    assert.Empty(err)
    assert.Empty(val)

    val, err = ver.checkCellType("2")
    assert.Empty(err)
    assert.Equal(val, 2)

    err = ver.checkCellLength(2, 0, 4)
    assert.Empty(err)

    err = ver.checkCellLength(2, 0, 0)
    assert.Error(err)
}

func TestVerifierFloat(t *testing.T) {
    assert := assert.New(t)

    var (
        err error
        val interface{}
        ver verifierFloat
    )

    val, err = ver.checkCellType("text")
    assert.Error(err)
    assert.Empty(val)

    val, err = ver.checkCellType("0")
    assert.Empty(err)
    assert.Empty(val)

    val, err = ver.checkCellType("2.0")
    assert.Empty(err)
    assert.Equal(val, 2.0)

    err = ver.checkCellLength(float32(2), 0, 4)
    assert.Empty(err)

    err = ver.checkCellLength(float32(2), 0, 0)
    assert.Error(err)
}
