package importer

import (
    "testing"

    "github.com/stretchr/testify/assert"
)

func TestNewTester(tt *testing.T) {
    assert := assert.New(tt)

    var (
        err error
        t   tester
    )

    t, err = newTester(0)
    assert.Empty(err)
    assert.Exactly(t, testString{})

    t, err = newTester(1)
    assert.Empty(err)
    assert.Exactly(t, testInteger{})

    t, err = newTester(2)
    assert.Empty(err)
    assert.Exactly(t, testFloat{})

    t, err = newTester(5)
    assert.Error(err)
    assert.Empty(t)
}

func TestTestString(tt *testing.T) {
    assert := assert.New(tt)

    var (
        err error
        val interface{}
        t   testString
    )

    val, err = t.testCellType("")
    assert.Empty(err)
    assert.Empty(val)

    val, err = t.testCellType("txt")
    assert.Empty(err)
    assert.Equal(val, "txt")

    err = t.testCellLength("txt", 0, 4)
    assert.Empty(err)

    err = t.testCellLength("txt", 0, 0)
    assert.Error(err)
}

func TestTestInteger(tt *testing.T) {
    assert := assert.New(tt)

    var (
        err error
        val interface{}
        t   testInteger
    )

    val, err = t.testCellType("text")
    assert.Error(err)
    assert.Empty(val)

    val, err = t.testCellType("0")
    assert.Empty(err)
    assert.Empty(val)

    val, err = t.testCellType("2")
    assert.Empty(err)
    assert.Equal(val, 2)

    err = t.testCellLength(2, 0, 4)
    assert.Empty(err)

    err = t.testCellLength(2, 0, 0)
    assert.Error(err)
}

func TestTestFloat(tt *testing.T) {
    assert := assert.New(tt)

    var (
        err error
        val interface{}
        t   testFloat
    )

    val, err = t.testCellType("txt")
    assert.Error(err)
    assert.Empty(val)

    val, err = t.testCellType("0")
    assert.Empty(err)
    assert.Empty(val)

    val, err = t.testCellType("2.0")
    assert.Empty(err)
    assert.Equal(val, 2.0)

    err = t.testCellLength(float32(2), 0, 4)
    assert.Empty(err)

    err = t.testCellLength(float32(2), 0, 0)
    assert.Error(err)
}
