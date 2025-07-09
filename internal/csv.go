package internal

import (
	"encoding/base64"
	"fmt"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

func (a *App) loadCSV(f FileData) (*excelize.File, string, error) {
	data, err := base64.StdEncoding.DecodeString(f.Data)
	if err != nil {
		return nil, "", fmt.Errorf("%s のデコードに失敗: %w", f.Name, err)
	}
	// xlsxファイルは保存せず、メモリ上で処理
	tmpFile, err := os.CreateTemp("", "tmpxlsx-*.xlsx")
	if err != nil {
		return nil, "", fmt.Errorf("一時ファイル作成失敗: %w", err)
	}
	defer os.Remove(tmpFile.Name())
	_, err = tmpFile.Write(data)
	if err != nil {
		tmpFile.Close()
		return nil, "", fmt.Errorf("一時ファイル書込失敗: %w", err)
	}

	tmpFile.Close()

	// Excelファイルを開く
	fx, err := excelize.OpenFile(tmpFile.Name())
	name := filepath.Base(f.Name)
	if err != nil {
		return nil, "", fmt.Errorf("%s のExcel読込に失敗: %w", f.Name, err)
	}
	return fx, name, nil
}

// シート名取得（最初のシート）

func (a *App) loadData(sheet string, fx *excelize.File) ([][]string, error) {

	// A1:AD48のデータ取得
	rows, err := fx.Rows(sheet)
	if err != nil {
		return nil, fmt.Errorf("範囲取得失敗: %w", err)
	}
	var tableData [][]string
	rowIdx := 0
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			return nil, fmt.Errorf("行取得失敗: %w", err)
		}
		if rowIdx >= 48 {
			break
		}
		// 30列分だけ取得
		rowData := make([]string, 30)
		for i := 0; i < 30; i++ {
			if i < len(row) {
				rowData[i] = row[i]
			} else {
				rowData[i] = ""
			}
		}
		tableData = append(tableData, rowData)
		rowIdx++
	}
	return tableData, nil
}
