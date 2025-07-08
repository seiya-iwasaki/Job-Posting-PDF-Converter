package internal

import (
	"embed"
	"fmt"
	"io"
	"os"
)

//go:embed fonts/*
var fontAssets embed.FS

// loadFont は埋め込みフォントを一時ファイルに展開し、ファイルとそのパスを返す。
// 呼び出し元で Close() と Remove() を行うこと。
func (a *App) loadFont() (*os.File, string, error) {
	tmpFontFile, err := os.CreateTemp("", "ipaexg-*.ttf")
	if err != nil {
		return nil, "", fmt.Errorf("一時フォントファイルの作成に失敗: %w", err)
	}

	// 埋め込みフォント読み込み
	ff, err := fontAssets.Open("fonts/ipaexg.ttf")
	if err != nil {
		tmpFontFile.Close()
		os.Remove(tmpFontFile.Name())
		return nil, "", fmt.Errorf("フォントの読み込みに失敗: %w", err)
	}
	defer ff.Close()

	_, err = io.Copy(tmpFontFile, ff)
	if err != nil {
		tmpFontFile.Close()
		os.Remove(tmpFontFile.Name())
		return nil, "", fmt.Errorf("一時フォントファイルへの書き込みに失敗: %w", err)
	}

	// 呼び出し元で明示的に Close & Remove すること
	return tmpFontFile, tmpFontFile.Name(), nil
}
