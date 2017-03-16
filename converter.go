package excel

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type Excel ole.IDispatch
type Workbooks ole.IDispatch
type Workbook ole.IDispatch
type Worksheet ole.IDispatch
type Cell ole.IDispatch

const (
	xlsDefaultVersion = -4143
	xlsVersion12      = 56
	ContentTypeXLS    = "application/vnd.ms-excel"
	ContentTypeXLSX   = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

func NewExcel() (*Excel, error) {
	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		return nil, err
	}
	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, err
	}
	return (*Excel)(excel), nil
}

func (e *Excel) Workbooks() (*Workbooks, error) {
	res, err := oleutil.GetProperty((*ole.IDispatch)(e), "Workbooks")
	if err != nil {
		return nil, err
	}
	workbooks := res.ToIDispatch()
	return (*Workbooks)(workbooks), nil
}

func (e *Excel) release() {
	(*ole.IDispatch)(e).Release()
}

func (e *Excel) Version() (int, error) {
	res, err := oleutil.GetProperty((*ole.IDispatch)(e), "Version")
	if err != nil {
		return 0, err
	}
	if res.ToString() == "12.0" {
		return xlsVersion12, nil
	} else {
		return xlsDefaultVersion, nil
	}
}

func (e *Excel) Quit() error {
	e.release()
	_, err := oleutil.CallMethod((*ole.IDispatch)(e), "Quit")
	return err
}

func (w *Workbooks) Workbook(filename string) (*Workbook, error) {
	res, err := oleutil.CallMethod((*ole.IDispatch)(w), "Open", filename)
	if err != nil {
		return nil, err
	}
	workbook := res.ToIDispatch()
	return (*Workbook)(workbook), nil
}

func (w *Workbooks) Close() error {
	_, err := oleutil.CallMethod((*ole.IDispatch)(w), "Close")
	return err
}

func (w *Workbook) Worksheet(id int) (*Worksheet, error) {
	res, err := oleutil.GetProperty((*ole.IDispatch)(w), "Worksheets", id)
	if err != nil {
		return nil, err
	}
	worksheet := res.ToIDispatch()
	return (*Worksheet)(worksheet), nil
}

func (w *Workbook) Saved(saved bool) error {
	_, err := oleutil.PutProperty((*ole.IDispatch)(w), "Saved", saved)
	return err
}

func (w *Workbook) SaveAs(filename string, version int) error {
	res, err := oleutil.CallMethod((*ole.IDispatch)(w), "SaveAs", filename, version, nil, nil)
	if err != nil {
		return err
	}
	res.ToIDispatch()
	return nil
}

func (w *Worksheet) Cell(x, y int) (*Cell, error) {
	res, err := oleutil.GetProperty((*ole.IDispatch)(w), "Cells", y, x)
	if err != nil {
		return nil, err
	}
	cell := res.ToIDispatch()
	return (*Cell)(cell), nil
}

func (c *Cell) Set(src interface{}) error {
	_, err := oleutil.PutProperty((*ole.IDispatch)(c), "Value", src)
	return err
}
