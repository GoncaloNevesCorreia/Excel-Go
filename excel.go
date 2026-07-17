package excel

import (
	"bytes"
	"encoding/base64"
	"fmt"
	"math/big"
	"path/filepath"
	"reflect"
	"slices"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type Result[T any] struct {
	Success bool
	Data    []T
	Error   *FileError
}

type FileError struct {
	Code    string
	Message string
}

func Parse(blob string, filename string) Result[byte] {
	fileData, err := base64.StdEncoding.DecodeString(blob)

	if err != nil {
		return Result[byte]{
			Error: &FileError{
				Code:    "invalid_file_base64",
				Message: err.Error(),
			},
		}
	}

	ext := filepath.Ext(filename)

	if ext != ".xlsx" && ext != ".xlsm" {
		return Result[byte]{
			Error: &FileError{
				Code:    "invalid_file_format",
				Message: fmt.Sprintf("Extensão não suportada: %s", ext),
			},
		}
	}

	return Result[byte]{
		Success: true,
		Data:    fileData,
	}
}

func ReadBytes[T any](data []byte, sheets []string) Result[T] {
	reader := bytes.NewReader(data)

	f, err := excelize.OpenReader(reader)

	if err != nil {
		return Result[T]{
			Error: &FileError{
				Code:    "failed_open_reader",
				Message: err.Error(),
			},
		}
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	sheetName, err := chooseSheet(f, sheets)

	if err != nil {
		return Result[T]{
			Error: &FileError{
				Code:    "invalid_file",
				Message: err.Error(),
			},
		}
	}

	return readSheet[T](f, sheetName)
}

func ReadBook[T any](filePath string, sheets []string) Result[T] {
	f, err := excelize.OpenFile(filePath)

	if err != nil {
		return Result[T]{
			Error: &FileError{
				Code:    "failed_open_file",
				Message: err.Error(),
			},
		}
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	sheetName, err := chooseSheet(f, sheets)

	if err != nil {
		return Result[T]{
			Error: &FileError{
				Code:    "invalid_file",
				Message: err.Error(),
			},
		}
	}

	return readSheet[T](f, sheetName)
}

func WriteBook[T any](file string, sheet string, items []T) Result[T] {

	f := excelize.NewFile()

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	sw, err := f.NewStreamWriter(sheet)

	if err != nil {
		panic(err)
	}

	response := saveStruct(f, sw, items)

	if !response.Success {
		return response
	}

	// Save spreadsheet by the given path.
	if err := f.SaveAs(file); err != nil {
		return Result[T]{
			Error: &FileError{
				Code:    "failed_write_file",
				Message: err.Error(),
			},
		}
	}

	return response
}

func readSheet[T any](f *excelize.File, sheet string) Result[T] {
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows(sheet)

	if err != nil {
		return Result[T]{
			Error: &FileError{
				Code:    "failed_read_sheet",
				Message: err.Error(),
			},
		}
	}

	maxSize := len(rows[0])

	for i, row := range rows {
		rowSize := len(row)

		if rowSize >= maxSize {
			continue
		}

		padding := make([]string, maxSize-rowSize)

		rows[i] = append(rows[i], padding...)
	}

	return loadStruct[T](rows)
}

func chooseSheet(f *excelize.File, sheets []string) (string, error) {
	fileSheets := f.GetSheetList()

	for _, name := range sheets {
		if slices.Contains(fileSheets, name) {
			return name, nil
		}
	}

	return "", fmt.Errorf("O documento necessita de ter uma das seguintes folhas: %v", sheets)
}

func getColumns[T any]() ([]string, []any) {
	t := reflect.TypeFor[T]()

	if t.Kind() != reflect.Struct {
		panic(fmt.Errorf("%s must be a struct", t.Name()))
	}

	keys := []string{}
	columns := make([]any, 0)

	for field := range t.Fields() {
		if field.PkgPath != "" {
			continue
		}

		if jsonTag, ok := field.Tag.Lookup("json"); ok {

			jsonColumn, _, _ := strings.Cut(jsonTag, ",")

			if len(jsonColumn) > 0 {
				keys = append(keys, field.Name)
				columns = append(columns, jsonColumn)
			}
		}
	}

	if len(columns) == 0 {
		panic(fmt.Errorf("No Columns found in struct '%s'. Please include the missing json tags", t.Name()))
	}

	return keys, columns
}

func saveStruct[T any](f *excelize.File, sw *excelize.StreamWriter, items []T) Result[T] {

	if reflect.TypeFor[T]().Kind() != reflect.Struct {
		panic(fmt.Errorf("T must be a struct"))
	}

	keys, columns := getColumns[T]()

	cell := "A" + strconv.Itoa(1)

	if err := sw.SetRow(cell, columns); err != nil {

		return Result[T]{
			Error: &FileError{
				Code:    "failed_write_file",
				Message: err.Error(),
			},
		}
	}

	if len(items) == 0 {
		return Result[T]{
			Success: true,
		}
	}

	size := len(columns)

	row := make([]any, size)

	for i, item := range items {

		reflectValue := reflect.ValueOf(&item).Elem()

		row = row[:size]

		for j, key := range keys {

			fieldValue := reflectValue.FieldByName(key)

			valueType := fieldValue.Type()

			if isDuration(valueType) {

				duraration, err := durationCell(f, fieldValue.Interface().(time.Duration))

				if err != nil {
					return Result[T]{
						Error: &FileError{
							Code:    "parse_error",
							Message: fmt.Sprintf("Não foi possivel converter o valor '%v' da coluna '%s' para um time.Duration.", fieldValue, columns[j]),
						},
					}
				}

				row[j] = duraration

				continue
			}

			row[j] = fieldValue.Interface()
		}

		cell = "A" + strconv.Itoa(i+2)

		if err := sw.SetRow(cell, row); err != nil {
			return Result[T]{
				Error: &FileError{
					Code:    "failed_write_file",
					Message: err.Error(),
				},
			}
		}
	}

	if err := sw.Flush(); err != nil {
		return Result[T]{
			Error: &FileError{
				Code:    "failed_write_file",
				Message: err.Error(),
			},
		}
	}

	return Result[T]{
		Success: true,
	}
}

func loadStruct[T any](values [][]string) Result[T] {

	if reflect.TypeFor[T]().Kind() != reflect.Struct {
		panic(fmt.Errorf("T must be a struct"))
	}

	if len(values) == 0 {
		return Result[T]{
			Success: true,
			Data:    []T{},
		}
	}

	headers := values[0]

	headerIndex := make(map[string]int)

	for i, header := range headers {
		trimedHeader := strings.ToLower(strings.TrimSpace(header))

		if trimedHeader == "" {
			continue
		}

		headerIndex[trimedHeader] = i
	}

	result := make([]T, 0, len(values)-1)

	notFoundColumns := make([]string, 0)

	for _, row := range values[1:] {
		var item T

		v := reflect.ValueOf(&item).Elem()
		t := v.Type()

		for i := 0; i < t.NumField(); i++ {
			structField := t.Field(i)
			valueField := v.Field(i)

			if !valueField.CanSet() {
				continue
			}

			jsonTag := structField.Tag.Get("json")
			columnName, extra, _ := strings.Cut(jsonTag, ",")

			isOptional := extra == "omitempty"

			if columnName == "" {
				continue
			}

			columnIndex, exists := headerIndex[strings.ToLower(columnName)]

			if !exists {
				if isOptional {
					valueField.SetZero()
					continue
				}

				notFoundColumns = append(notFoundColumns, columnName)
				continue
			}

			if columnIndex >= len(row) {
				continue
			}

			value := row[columnIndex]

			valueType := valueField.Type()

			if valueType.Kind() == reflect.Pointer {
				valueType = valueType.Elem()

				valueField.Set(reflect.New(valueField.Type().Elem()))

				valueField = valueField.Elem()
			}

			// TODO: Adicionar converção automatica de strings no excel para:
			// DD/MM/YYYY or DD/MM/YYYY hh:mm:ss
			if isTime(valueType) {
				// TODO: NOT WORKING
				parsedTime, err := time.Parse("02:09:00", value)

				if err != nil {
					return Result[T]{
						Error: &FileError{
							Code:    "parse_error",
							Message: fmt.Sprintf("Não foi possivel converter o valor '%v' da coluna '%s' para um time.Time.", value, columnName),
						},
					}
				}

				valueField.Set(reflect.ValueOf(parsedTime))

				continue
			} else if isDuration(valueType) {
				parseDuration, err := parseHHMMSS(value)

				if err != nil {
					return Result[T]{
						Error: &FileError{
							Code:    "parse_error",
							Message: fmt.Sprintf("Não foi possivel converter o valor '%v' da coluna '%s' para um time.Duration.", value, columnName),
						},
					}
				}

				valueField.Set(reflect.ValueOf(parseDuration))

				continue
			}

			switch valueField.Kind() {
			case reflect.String:
				valueField.SetString(value)

			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
				if value == "" {
					valueField.SetInt(0)
					continue
				}

				valueInt, err := strconv.ParseInt(value, 10, bitSize(valueField))

				if err != nil {
					return Result[T]{
						Error: &FileError{
							Code:    "parse_error",
							Message: fmt.Sprintf("Não foi possivel converter o valor '%v' da coluna '%s' para um int.", value, columnName),
						},
					}
				}

				valueField.SetInt(valueInt)

			case reflect.Float32, reflect.Float64:
				if value == "" {
					valueField.SetFloat(0)
					continue
				}

				valueFloat, err := strconv.ParseFloat(value, bitSize(valueField))

				if err != nil {
					return Result[T]{
						Error: &FileError{
							Code:    "parse_error",
							Message: fmt.Sprintf("Não foi possivel converter o valor '%v' da coluna '%s' para um float.", value, columnName),
						},
					}
				}

				valueField.SetFloat(valueFloat)

			case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
				if value == "" {
					valueField.SetUint(0)
					continue
				}

				valueUint, err := strconv.ParseUint(value, 10, bitSize(valueField))

				if err != nil {
					return Result[T]{
						Error: &FileError{
							Code:    "parse_error",
							Message: fmt.Sprintf("Não foi possivel converter o valor '%v' da coluna '%s' para um uint.", value, columnName),
						},
					}
				}

				valueField.SetUint(valueUint)

			case reflect.Bool:
				if value == "" {
					valueField.SetBool(false)
					continue
				}

				valueBool, err := strconv.ParseBool(value)

				if err != nil {
					return Result[T]{
						Error: &FileError{
							Code:    "parse_error",
							Message: fmt.Sprintf("Não foi possivel converter o valor '%v' da coluna '%s' para um bool.", value, columnName),
						},
					}
				}

				valueField.SetBool(valueBool)

			default:

				if isBigFloat(valueType) {

					if value == "" {
						valueField.Set(reflect.ValueOf(big.NewFloat(0)))
						continue
					}

					precision := uint(len(value[strings.Index(value, ".")+1:]))

					valueFloat, _, err := big.ParseFloat(value, 10, precision*4, big.RoundingMode(big.Exact))

					if err != nil {
						return Result[T]{
							Error: &FileError{
								Code:    "parse_error",
								Message: fmt.Sprintf("Não foi possivel converter o valor '%v' da coluna '%s' para um BigFloat.", value, columnName),
							},
						}
					}

					valueField.Set(reflect.ValueOf(valueFloat))

				}
			}
		}

		if len(notFoundColumns) != 0 {
			return Result[T]{
				Error: &FileError{
					Code:    "parse_error",
					Message: fmt.Sprintf("As colunas %v não foram encontradas no ficheiro selecionado.", notFoundColumns),
				},
			}
		}

		result = append(result, item)
	}

	return Result[T]{
		Success: true,
		Data:    result,
	}
}

func bitSize(reflectValue reflect.Value) int {

	return int(reflectValue.Type().Size() * 8)
}

func isBigFloat(t reflect.Type) bool {
	tPtrBigFloat := reflect.TypeFor[*big.Float]()

	return t == tPtrBigFloat
}

func isTime(t reflect.Type) bool {
	tTimeDuration := reflect.TypeFor[time.Time]()

	return t == tTimeDuration
}

func isDuration(t reflect.Type) bool {
	tTimeDurationPtr := reflect.TypeFor[*time.Duration]()
	tTimeDuration := reflect.TypeFor[time.Duration]()

	return t == tTimeDuration || t == tTimeDurationPtr
}

func parseHHMMSS(input string) (time.Duration, error) {
	if input == "" {
		return 0, nil
	}

	var h, m, s int

	count := strings.Count(input, ":")

	switch count {
	case 2:
		_, err := fmt.Sscanf(input, "%d:%d:%d", &h, &m, &s)

		if err != nil {
			return 0, err
		}

	case 1:
		_, err := fmt.Sscanf(input, "%d:%d", &h, &m)

		if err != nil {
			return 0, err
		}

	default:
		_, err := fmt.Sscanf(input, "%d", &h)

		if err != nil {
			return 0, err
		}
	}

	return time.Duration(h)*time.Hour +
		time.Duration(m)*time.Minute +
		time.Duration(s)*time.Second, nil
}

func durationToExcelTime(d time.Duration) float64 {
	return float64(d) / float64(24*time.Hour)
}

func durationCell(f *excelize.File, d time.Duration) (excelize.Cell, error) {
	format := "hh:mm:ss"

	styleID, err := f.NewStyle(&excelize.Style{
		CustomNumFmt: &format,
	})

	if err != nil {
		return excelize.Cell{}, err
	}

	return excelize.Cell{
		StyleID: styleID,
		Value:   durationToExcelTime(d),
	}, nil
}
