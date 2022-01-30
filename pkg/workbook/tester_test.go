package workbook

import (
    "testing"

    "github.com/stretchr/testify/assert"
)

func TestNewCellTester(t *testing.T) {
    assert := assert.New(t)

    var (
        err    error
        tester verifier
    )

    tester, err = newTester(Hazop.CellType.String)
    assert.Empty(err)
    assert.Exactly(tester, verifierString{})

    tester, err = newTester(Hazop.CellType.Integer)
    assert.Empty(err)
    assert.Exactly(tester, verifierInteger{})

    tester, err = newTester(Hazop.CellType.Float)
    assert.Empty(err)
    assert.Exactly(tester, verifierFloat{})

    tester, err = newTester(4)
    assert.Error(err)
    assert.Empty(tester)
}

func TestTesterString(t *testing.T) {
    assert := assert.New(t)

    var (
        err    error
        val    interface{}
        tester verifierString
    )

    val, err = tester.testValueType("")
    assert.Empty(err)
    assert.Empty(val)

    val, err = tester.testValueType("text")
    assert.Empty(err)
    assert.Equal(val, "text")

    err = tester.testValueLength("text", 0, 4)
    assert.Empty(err)

    err = tester.testValueLength("text", 0, 0)
    assert.Error(err)
}

func TestTesterInteger(t *testing.T) {
    assert := assert.New(t)

    var (
        err    error
        val    interface{}
        tester verifierInteger
    )

    val, err = tester.testValueType("text")
    assert.Error(err)
    assert.Empty(val)

    val, err = tester.testValueType("0")
    assert.Empty(err)
    assert.Empty(val)

    val, err = tester.testValueType("2")
    assert.Empty(err)
    assert.Equal(val, 2)

    err = tester.testValueLength(2, 0, 4)
    assert.Empty(err)

    err = tester.testValueLength(2, 0, 0)
    assert.Error(err)
}

func TestTesterFloat(t *testing.T) {
    assert := assert.New(t)

    var (
        err    error
        val    interface{}
        tester verifierFloat
    )

    val, err = tester.testValueType("text")
    assert.Error(err)
    assert.Empty(val)

    val, err = tester.testValueType("0")
    assert.Empty(err)
    assert.Empty(val)

    val, err = tester.testValueType("2.0")
    assert.Empty(err)
    assert.Equal(val, 2.0)

    err = tester.testValueLength(float32(2), 0, 4)
    assert.Empty(err)

    err = tester.testValueLength(float32(2), 0, 0)
    assert.Error(err)
}
