// +build windows

package wmi

import (
	"fmt"
	"regexp"
	"runtime"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func QueryMap(query string, connectServerArgs ...interface{}) ([]map[string]interface{}, error) {
	res := []map[string]interface{}{}

	queryLower := strings.ToLower(query)

	startCol := strings.Index(queryLower, "select ")
	if startCol == -1 {
		return res, fmt.Errorf("Invalid query, missing select")
	}

	startCol += 7

	endCol := strings.Index(queryLower, " from")
	if endCol == -1 {
		return res, fmt.Errorf("Invalid query, missing from")
	}

	strColumns := query[startCol:endCol]

	// Querystring is for example: "SELECT blah, blah2 FROM..." so everything between SELECT and FROM
	columns := regexp.MustCompile(`,\s*`).Split(strColumns, -1)

	lock.Lock()
	defer lock.Unlock()
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err != nil {
		oleCode := err.(*ole.OleError).Code()
		if oleCode != ole.S_OK && oleCode != S_FALSE {
			return res, err
		}
	}
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("WbemScripting.SWbemLocator")
	if err != nil {
		return res, err
	} else if unknown == nil {
		return res, ErrNilCreateObject
	}
	defer unknown.Release()

	wmi, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return res, err
	}
	defer wmi.Release()

	// service is a SWbemServices
	serviceRaw, err := oleutil.CallMethod(wmi, "ConnectServer", connectServerArgs...)
	if err != nil {
		return res, err
	}
	service := serviceRaw.ToIDispatch()
	defer serviceRaw.Clear()

	// result is a SWBemObjectSet
	resultRaw, err := oleutil.CallMethod(service, "ExecQuery", query)
	if err != nil {
		return res, err
	}
	result := resultRaw.ToIDispatch()
	defer resultRaw.Clear()

	enumProperty, err := result.GetProperty("_NewEnum")
	if err != nil {
		return res, err
	}
	defer enumProperty.Clear()

	enum, err := enumProperty.ToIUnknown().IEnumVARIANT(ole.IID_IEnumVariant)
	if err != nil {
		return res, err
	}
	if enum == nil {
		return res, fmt.Errorf("can't get IEnumVARIANT, enum is nil")
	}
	defer enum.Release()

	for itemRaw, length, err := enum.Next(1); length > 0; itemRaw, length, err = enum.Next(1) {
		if err != nil {
			return res, err
		}

		err := func() error {
			// item is a SWbemObject, but really a Win32_Process
			item := itemRaw.ToIDispatch()
			defer item.Release()

			m := make(map[string]interface{})
			for _, c := range columns {
				prop, err := oleutil.GetProperty(item, c)
				if err != nil {
					return err
				}
				m[c] = prop.Value()
			}

			res = append(res, m)
			return nil
		}()
		if err != nil {
			return res, err
		}
	}

	return res, nil
}
