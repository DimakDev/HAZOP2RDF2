package workbook

import (
    "testing"

    "github.com/stretchr/testify/assert"
)

func TestNewChecker(t *testing.T) {
    assert := assert.New(t)

    var (
        err error
        c   checker
    )

    c, err = newChecker(STRING)
    assert.Empty(err)
    assert.Exactly(c, checkString{})

    c, err = newChecker(INTEGER)
    assert.Empty(err)
    assert.Exactly(c, checkInteger{})

    c, err = newChecker(FLOAT)
    assert.Empty(err)
    assert.Exactly(c, checkFloat{})

    c, err = newChecker(4)
    assert.Error(err)
    assert.Empty(c)
}

func TestCheckString(t *testing.T) {
    assert := assert.New(t)

    var (
        err error
        val interface{}
        c   checkString
    )

    val, err = c.checkValueType("")
    assert.Empty(err)
    assert.Empty(val)

    val, err = c.checkValueType("text")
    assert.Empty(err)
    assert.Equal(val, "text")

    err = c.checkValueLength("text", 0, 4)
    assert.Empty(err)

    err = c.checkValueLength("text", 0, 0)
    assert.Error(err)
}

func TestCheckInteger(t *testing.T) {
    assert := assert.New(t)

    var (
        err error
        val interface{}
        c   checkInteger
    )

    val, err = c.checkValueType("text")
    assert.Error(err)
    assert.Empty(val)

    val, err = c.checkValueType("0")
    assert.Empty(err)
    assert.Empty(val)

    val, err = c.checkValueType("2")
    assert.Empty(err)
    assert.Equal(val, 2)

    err = c.checkValueLength(2, 0, 4)
    assert.Empty(err)

    err = c.checkValueLength(2, 0, 0)
    assert.Error(err)
}

func TestCheckFloat(t *testing.T) {
    assert := assert.New(t)

    var (
        err error
        val interface{}
        c   checkFloat
    )

    val, err = c.checkValueType("text")
    assert.Error(err)
    assert.Empty(val)

    val, err = c.checkValueType("0")
    assert.Empty(err)
    assert.Empty(val)

    val, err = c.checkValueType("2.0")
    assert.Empty(err)
    assert.Equal(val, 2.0)

    err = c.checkValueLength(float32(2), 0, 4)
    assert.Empty(err)

    err = c.checkValueLength(float32(2), 0, 0)
    assert.Error(err)
}
