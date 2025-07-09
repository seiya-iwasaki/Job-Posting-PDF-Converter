package internal

import (
	"context"
	"fmt"
	"log"
	"os"
	"os/user"
	"path/filepath"
	"time"

	// "myapp/internal/pdf"

	"github.com/jung-kurt/gofpdf"
)

// FileData: フロントエンドから受け取るファイル情報
// DataはBase64エンコードされたファイル内容
type FileData struct {
	Name string `json:"name"`
	Data string `json:"data"`
}

// App struct
type App struct {
	ctx context.Context
}

// NewApp creates a new App application struct
func NewApp() *App {
	return &App{}
}

// startup is called when the app starts. The context is saved
// so we can call the runtime methods
func (a *App) Startup(ctx context.Context) {
	a.ctx = ctx
}

func GetDownloadsPath() (string, error) {
	usr, err := user.Current()
	if err != nil {
		return "", err
	}
	return filepath.Join(usr.HomeDir, "Downloads"), nil
}

type Table struct {
	pdf            *gofpdf.Fpdf
	x_i, y_i       float64   // 左上座標
	x_f, y_f       float64   // 右下座標
	Xs             []float64 // 各列のX座標
	Ys             []float64 // 各行のY座標
	font           string
	fontSize       float64
	default_H      float64 // デフォルトの行の高さ
	border         string
	rowNum         int        // 行数
	Rows           []Row      // 行のY座標とページ数
	Cells          []CellInfo // データ
	Rects          []RectInfo // セルの矩形情報
	Texts          []Text     // テキスト情報（タイトル用）
	margin         float64    // セルの余白
	initialpageNum int        // ページ数
	pageNum        int        // 描画準備時のページ数管理
	titleW         float64    // タイトルの幅
}

type CellInfo struct {
	x         float64 // X座標
	y         float64 // Y座標
	w         float64 // 幅
	h         float64 // 高さ
	pageNum   int     // ページ数
	border    string  // セルの枠線スタイル（"0" = なし, "1" = 枠線）
	col_i     int     // 開始列インデックス
	row_i     int     // 開始行インデックス
	col_f     int     // 終了列インデックス
	row_f     int     // 終了行インデックス
	text      string  // セルのテキスト
	align     string  // テキストの配置
	fill      bool    // 塗りつぶしフラグ
	fontSize  float64 // フォントサイズ
	link      string  // リンクURL
	LineWidth float64 // 線の太さ
}

type RectInfo struct {
	x         float64 // X座標
	y         float64 // Y座標
	w         float64 // 幅
	h         float64 // 高さ
	pageNum   int     // ページ数
	style     string  // スタイル（"F" = 塗りつぶし, "D" = 枠線）
	LineWidth float64 // 線の太さ
}

type Row struct {
	y       float64 // 行のY座標
	pageNum int     // ページ数
}

type Text struct {
	x       float64 // X座標
	y       float64 // Y座標
	text    string  // テキスト
	size    float64 // フォントサイズ
	pageNum int     // ページ数
}

// コンストラクタ
func NewTable(pdf *gofpdf.Fpdf, x_i, y_i, x_f, y_f float64, colNum int, rowNum int, font string, fontSize, default_H float64, border string) *Table {
	t := &Table{
		pdf:       pdf,
		x_i:       x_i,
		y_i:       y_i,
		x_f:       x_f,
		y_f:       y_f,
		font:      font,
		fontSize:  fontSize,
		default_H: default_H,
		border:    border,
		rowNum:    rowNum, // 行数は列数と同じ
	}

	t.titleW = 5.0 // タイトルの幅を設定
	t.margin = 20

	// データを初期化
	t.Rows = []Row{}
	t.Cells = []CellInfo{}
	t.Rects = []RectInfo{}
	t.Texts = []Text{}
	t.initialpageNum = pdf.PageNo() // 初期ページ番号を保存

	// 列のX座標を計算
	t.Xs = make([]float64, colNum+1)
	colWidth := (x_f - x_i) / float64(colNum)
	for i := 0; i <= colNum; i++ {
		t.Xs[i] = x_i + float64(i)*colWidth
	}

	// 行のY座標を入れていくところ
	t.Ys = []float64{}

	_, pageHeight := t.pdf.GetPageSize()
	if y_i > pageHeight {
		fmt.Print("[Render] y_i exceeds page height\n")
		y_i = 20.0
		pdf.AddPage() // 新しいページを作成
	}
	t.pageNum = pdf.PageNo()
	// 行のY座標を計算
	t.Ys = append(t.Ys, y_i)
	t.Rows = append(t.Rows, Row{
		y:       y_i,
		pageNum: pdf.PageNo(),
	})
	fmt.Print("[Render] Initialized table\n")

	return t
}

// コンストラクタ（付録用）
func NewAppendix(pdf *gofpdf.Fpdf, x_i, x_f, y_i float64, font string, fontSize, default_H float64, border string) *Table {
	t := &Table{
		pdf:       pdf,
		x_i:       x_i,
		y_i:       y_i,
		x_f:       x_f,
		y_f:       y_i, // 付録は1行のみなのでとりあえずy_fはy_iにしておく
		font:      font,
		fontSize:  fontSize,
		default_H: default_H,
		border:    border,
		rowNum:    1, // 付録は1行のみ
	}

	t.titleW = 5.0 // タイトルの幅を設定
	t.margin = 20

	// データを初期化
	t.Rows = []Row{}
	t.Cells = []CellInfo{}
	t.Rects = []RectInfo{}
	t.Texts = []Text{}
	t.initialpageNum = pdf.PageNo() // 初期ページ番号を保存
	t.pageNum = pdf.PageNo()

	t.Xs = []float64{x_i, x_f} // 左右のX座標
	t.Ys = []float64{y_i, y_i} // 上下のY座標

	return t
}

func (t *Table) GetBottomLine(pageNum int) float64 {
	maxY := 0.0

	// Cells から最大 y+h を探す
	for _, cell := range t.Cells {
		if cell.pageNum == pageNum {
			if yBottom := cell.y + cell.h; yBottom > maxY {
				maxY = yBottom
			}
		}
	}

	// Rects から最大 y+h を探す
	for _, rect := range t.Rects {
		if rect.pageNum == pageNum {
			if yBottom := rect.y + rect.h; yBottom > maxY {
				maxY = yBottom
			}
		}
	}

	bottom := maxY

	return bottom
}

func (t *Table) GetTopLine(pageNum int) float64 {
	minY := 1000.0 // 初期値は大きな値に設定

	// Cells から最小 y を探す
	for _, cell := range t.Cells {
		if cell.pageNum == pageNum {
			if yBottom := cell.y; yBottom < minY {
				minY = yBottom
			}
		}
	}

	// Rects から最小 y を探す
	for _, rect := range t.Rects {
		if rect.pageNum == pageNum {
			if yBottom := rect.y; yBottom < minY {
				minY = yBottom
			}
		}
	}

	return minY
}

func (t *Table) SetCell(col_i, row_i, col_f, row_f int, text string, align string, fill bool, fontSize float64, link string, lineWidth float64, rowH float64) {

	// 未生成のrow_iは無効
	if row_i < 0 || row_i > len(t.Ys)-1 {
		fmt.Print("[Render] Invalid table: row index out of range\n")
		return
	}
	// テキストの幅が列の幅を超えている場合は無効
	w := t.Xs[col_f] - t.Xs[col_i]
	if t.pdf.GetStringWidth(text) > w {
		fmt.Print("[Render] Invalid table: text width exceeds column width\n")
		return
	}

	unitSize := rowH
	if fontSize < 0 {
		fontSize = t.fontSize
	}

	_, pageHeight := t.pdf.GetPageSize()
	if t.Ys[row_i]+unitSize <= pageHeight-t.margin { // 現在のページに収まる場合

		for i := len(t.Ys); i < row_f; i++ { // 未生成の間の行を高さ0で仮設定する
			t.Ys = append(t.Ys, t.Ys[len(t.Ys)-1])
			t.Rows = append(t.Rows, Row{
				y:       t.Ys[len(t.Ys)-1],
				pageNum: t.pageNum,
			})
		}

		if row_f < len(t.Ys) { // Ys[row_f]が存在する = すでに決められた下端行がある
			if unitSize < t.Ys[row_f]-t.Ys[row_i] { // 現状の下端行のY座標が小さい場合
				unitSize = t.Ys[row_f] - t.Ys[row_i] // セル高さ継承
			} else {
				t.Ys[row_f] = t.Ys[row_i] + unitSize // 下端行のY座標を更新
				t.Rows[row_f] = Row{
					y:       t.Ys[row_i] + unitSize,
					pageNum: t.pageNum,
				}
			}
		} else { // Ys[row_f]が存在しない = 新しい下端行を追加
			t.Ys = append(t.Ys, t.Ys[row_i]+unitSize) // 行のY座標を設定
			t.Rows = append(t.Rows, Row{
				y:       t.Ys[row_i] + unitSize,
				pageNum: t.pageNum,
			})
		}

		t.Cells = append(t.Cells, CellInfo{
			x:         t.Xs[col_i],
			y:         t.Ys[row_i],
			w:         w,
			h:         unitSize,
			pageNum:   t.pageNum,
			col_i:     col_i,
			row_i:     row_i,
			col_f:     col_f,
			row_f:     row_f,
			text:      text,
			align:     align,
			fill:      fill,
			fontSize:  fontSize,
			link:      link,
			LineWidth: lineWidth, // デフォルトの線の太さ
			border:    "1",       // セルの枠線スタイル
		})
	} else { // 現在のページに収まらない場合
		fmt.Print("[Render] Current page exceeds page height\n")
	}
	fmt.Print("[Render] SetCell completed: ", text, " at (", col_i, ",", row_i, ") to (", col_f, ",", row_f, ")\n")
}

func (t *Table) SetMultiRowCell(col_i, row_i, col_f, row_f int, text string, align string, fill bool, fontSize float64, breakLines bool) {

	// 未生成のrow_iは無効
	fmt.Print("[Render] SetMultiRowCell called: ", text, " at (", col_i, ",", row_i, ") to (", col_f, ",", row_f, ")\n")
	fmt.Print("[Render] t.Ys: ", t.Ys, "\n")
	if row_i < 0 || row_i > len(t.Ys)-1 {
		fmt.Print("[Render] Invalid table: row index out of range\n")
		return
	}

	if fontSize < 0 {
		fontSize = t.fontSize
	}

	w := t.Xs[col_f] - t.Xs[col_i]

	// 行数を計算
	t.pdf.SetFontSize(fontSize)
	_, unitSize := t.pdf.GetFontSize()
	lines := SplitByMaxChars(t.pdf, text, w, fontSize)
	if len(lines) == 0 {
		lines = []string{""} // 空のセルを作成
	}

	// 余白を設定
	default_Margin := t.default_H - unitSize // セルの上下余白
	if row_f < len(t.Ys) {                   // row_fが存在する場合
		default_Margin = t.Ys[row_f] - t.Ys[row_i] - unitSize*float64(len(lines)) // セルの上下余白を計算
	}
	_, pageHeight := t.pdf.GetPageSize()

	// 2ページ以上にわたる場合は無効
	if t.Ys[row_i]+default_Margin+float64(len(lines))*unitSize > pageHeight*2 {
		fmt.Print("[Render] Invalid table: text height exceeds page height\n")
		return
	}

	// 下端の保存
	bottomY := 0.0

	// 何行目までページに収まるかを計算
	residue := pageHeight - 20 - t.Ys[row_i] - default_Margin
	contanableLines := int(residue / unitSize)
	if !breakLines && contanableLines < len(lines) { // breakLinesがfalseで収まらない場合はcontanableLines = 0で強制改行
		contanableLines = 0
	}
	if contanableLines > len(lines) { // ページに収まる行数がテキストの行数を超える場合
		contanableLines = len(lines)
	}
	fmt.Printf("[Render] Contanable lines: %d, Total lines: %d, Unit size: %.2f, Residue: %.2f\n", contanableLines, len(lines), unitSize, residue)
	for i := 0; i < contanableLines; i++ {
		lineY := t.Ys[row_i] + float64(i)*unitSize + default_Margin/2
		t.Cells = append(t.Cells, CellInfo{
			x:         t.Xs[col_i],
			y:         lineY,
			w:         w,
			h:         unitSize,
			pageNum:   t.pageNum,
			col_i:     col_i,
			row_i:     row_i,
			col_f:     col_f,
			row_f:     row_f,
			text:      lines[i],
			align:     align,
			fill:      fill,
			fontSize:  fontSize,
			link:      "",
			LineWidth: 0.1, // デフォルトの線の太さ
			border:    "0", // セルの枠線スタイル
		})
	}
	if contanableLines > 0 { // ページに収まる行がある場合
		if fill {
			t.Rects = append(t.Rects, RectInfo{
				x:         t.Xs[col_i],
				y:         t.Ys[row_i],
				w:         w,
				h:         float64(contanableLines)*unitSize + default_Margin,
				pageNum:   t.pageNum,
				style:     "F", // 塗りつぶし
				LineWidth: 0.0, // デフォルトの線の太さ
			})
		}
		t.Rects = append(t.Rects, RectInfo{
			x:         t.Xs[col_i],
			y:         t.Ys[row_i],
			w:         w,
			h:         float64(contanableLines)*unitSize + default_Margin,
			pageNum:   t.pageNum,
			style:     "D", // 枠線
			LineWidth: 0.1, // デフォルトの線の太さ
		})
	}

	for i := len(t.Ys) - 1; i < row_f-1; i++ { // 未生成の間の行を高さ0で仮設定する
		t.Ys = append(t.Ys, t.Ys[len(t.Ys)-1])
		t.Rows = append(t.Rows, Row{
			y:       t.Ys[len(t.Ys)-1],
			pageNum: t.pageNum,
		})
	}

	if len(lines) > contanableLines { // ページに収まらない場合
		t.Ys[row_i] = t.margin
		t.pageNum++ // ページを追加
		for i := contanableLines; i < len(lines); i++ {
			lineY := t.margin + float64(i-contanableLines)*unitSize + default_Margin/2
			t.Cells = append(t.Cells, CellInfo{
				x:         t.Xs[col_i],
				y:         lineY,
				w:         w,
				h:         unitSize,
				pageNum:   t.pageNum,
				col_i:     col_i,
				row_i:     row_i,
				col_f:     col_f,
				row_f:     row_f,
				text:      lines[i],
				align:     align,
				fill:      fill,
				fontSize:  fontSize,
				link:      "",
				LineWidth: 0.1, // デフォルトの線の太さ
				border:    "0", // セルの枠線スタイル
			})
		}
		// セルの矩形情報を追加
		if fill {
			t.Rects = append(t.Rects, RectInfo{
				x:         t.Xs[col_i],
				y:         t.margin,
				w:         w,
				h:         float64(len(lines)-contanableLines)*unitSize + default_Margin,
				pageNum:   t.pageNum,
				style:     "F", // 塗りつぶし
				LineWidth: 0.0, // デフォルトの線の太さ
			})
		}
		t.Rects = append(t.Rects, RectInfo{
			x:         t.Xs[col_i],
			y:         t.margin,
			w:         w,
			h:         float64(len(lines)-contanableLines)*unitSize + default_Margin,
			pageNum:   t.pageNum,
			style:     "D", // 枠線
			LineWidth: 0.1, // デフォルトの線の太さ
		})

		bottomY = t.margin + float64(len(lines)-contanableLines)*unitSize + default_Margin
	} else {
		bottomY = t.Ys[row_i] + float64(contanableLines)*unitSize + default_Margin
	}

	if row_f < len(t.Ys) { // t.Ys[row_f]が存在する場合
		if t.Ys[row_f] < t.Ys[row_i]+unitSize { // 現状の下端行のY座標が小さい場合
			t.Ys[row_f] = bottomY // 下端行のY座標を更新
			t.Rows[row_f] = Row{
				y:       bottomY,
				pageNum: t.pageNum,
			}
		}
	} else {
		// 行のY座標を更新
		t.Ys = append(t.Ys, bottomY) // 行のY座標を設定
		t.Rows = append(t.Rows, Row{
			y:       bottomY,
			pageNum: t.pageNum,
		})
	}
	fmt.Print("[Render] SetMultiRowCell completed: ", text, " at (", col_i, ",", row_i, ") to (", col_f, ",", row_f, ")\n")
}

func (t *Table) SetAppendix(text string, align string, fill bool, fontSize float64, breakLines bool) {

	if fontSize < 0 {
		fontSize = t.fontSize
	}

	w := t.x_f - t.x_i // 最大幅を使用

	// 高さを計算
	t.pdf.SetFontSize(fontSize)
	_, unitSize := t.pdf.GetFontSize()
	lines := SplitByMaxChars(t.pdf, text, w, fontSize)

	// 余白を設定
	default_Margin := t.default_H - unitSize // セルの上下余白

	t.y_f = t.y_i + float64(len(lines))*unitSize + default_Margin // 右下座標を更新
	t.Ys[1] = t.y_f
	t.Rows = append(t.Rows, Row{
		y:       t.y_i,
		pageNum: t.pageNum,
	})

	// 2ページ以上にわたる場合は無効
	_, pageHeight := t.pdf.GetPageSize()
	if t.y_i+default_Margin+float64(len(lines))*unitSize > pageHeight*2 {
		return
	}

	// 何行目までページに収まるかを計算
	residue := pageHeight - default_Margin - t.y_i
	contanableLines := int(residue / unitSize)
	if contanableLines < len(lines) { // breakLinesがfalseで収まらない場合はcontanableLines = 0で強制改行
		fmt.Print("[Render] Contanable lines is less than total lines at appendix\n")
	}
	for i := 0; i < len(lines); i++ {
		lineY := t.y_i + float64(i)*unitSize + default_Margin/2
		t.Cells = append(t.Cells, CellInfo{
			x:         t.x_i,
			y:         lineY,
			w:         w,
			h:         unitSize,
			pageNum:   t.pageNum,
			col_i:     0,
			row_i:     0,
			col_f:     1,
			row_f:     1,
			text:      lines[i],
			align:     align,
			fill:      fill,
			fontSize:  fontSize,
			link:      "",
			LineWidth: 0.1, // デフォルトの線の太さ
			border:    "0", // セルの枠線スタイル
		})
	}
}

func (t *Table) SetCellWithTitle(col_i, row_i, col_f, row_f int, text string, align string, fill bool, fontSize float64) {

	if fontSize < 0 {
		fontSize = t.fontSize
	}

	t.pdf.SetFontSize(fontSize)
	_, unitSize := t.pdf.GetFontSize()

	x := t.Xs[col_i] - t.titleW // タイトル用に左に5mm余白を追加
	y := t.Ys[row_i]
	w := t.Xs[col_f] - x
	h := t.Ys[row_f] - y

	// 現在のページ内に収まったテーブルか
	if t.Rows[row_i].pageNum == t.Rows[row_f].pageNum { // 現在のページに収まる場合
		t.Cells = append(t.Cells, CellInfo{
			x:         x,
			y:         y,
			w:         w,
			h:         h,
			pageNum:   t.Rows[row_i].pageNum,
			col_i:     col_i,
			row_i:     row_i,
			col_f:     col_f,
			row_f:     row_f,
			text:      text,
			align:     align,
			fill:      fill,
			fontSize:  fontSize,
			link:      "",
			LineWidth: 0.1, // デフォルトの線の太さ
			border:    "1", // セルの枠線スタイル
		})
	} else { // 別れている場合
		bottom := t.GetBottomLine(t.Rows[row_i].pageNum)
		text1 := text
		text2 := " "
		if bottom-t.Rows[row_i].y < unitSize && bottom-t.Rows[row_i].y < t.Rows[row_f].y-t.margin {

		}
		t.Cells = append(t.Cells, CellInfo{
			x:         x,
			y:         y,
			w:         w,
			h:         bottom - t.Rows[row_i].y,
			pageNum:   t.Rows[row_i].pageNum,
			col_i:     col_i,
			row_i:     row_i,
			col_f:     col_f,
			row_f:     row_f,
			text:      text1,
			align:     align,
			fill:      fill,
			fontSize:  fontSize,
			LineWidth: 0.1, // デフォルトの線の太さ
			link:      "",
			border:    "1", // セルの枠線スタイル
		})

		t.Cells = append(t.Cells, CellInfo{
			x:         x,
			y:         t.margin,
			w:         w,
			h:         t.Rows[row_f].y - t.margin,
			pageNum:   t.Rows[row_f].pageNum,
			col_i:     col_i,
			row_i:     row_f,
			col_f:     col_f,
			row_f:     row_f,
			text:      text2,
			align:     align,
			fill:      fill,
			fontSize:  fontSize,
			link:      "",
			border:    "1", // セルの枠線スタイル
			LineWidth: 0.1, // デフォルトの線の太さ
		})
	}
}

func (t *Table) SetTitle(text string) {
	runes := []rune(text)

	t.pdf.SetFontSize(t.fontSize)
	_, unitSize := t.pdf.GetFontSize()

	if t.Rows[0].pageNum == t.Rows[len(t.Rows)-1].pageNum { // 表全体がページに収まっている場合
		page := t.Rows[0].pageNum
		// 総高さ = 行数 × 行の高さ
		textH := float64(len(runes)) * unitSize
		startY := t.y_i + (t.Ys[len(t.Ys)-1]-t.y_i-textH)/2

		// 塗りつぶし背景
		t.Rects = append(t.Rects, RectInfo{
			x:         t.x_i - t.titleW,
			y:         t.y_i,
			w:         t.titleW,
			h:         t.Ys[len(t.Ys)-1] - t.y_i,
			pageNum:   page,
			style:     "F", // 塗りつぶし
			LineWidth: 0.0,
		})

		// 枠線
		t.Rects = append(t.Rects, RectInfo{
			x:         t.x_i - t.titleW,
			y:         t.y_i,
			w:         t.titleW,
			h:         t.Ys[len(t.Ys)-1] - t.y_i,
			pageNum:   page,
			style:     "D", // 枠線
			LineWidth: 0.3,
		})

		// 一文字ずつ中央揃えで描画
		x := t.x_i - t.titleW + t.titleW/2 // 横は中央固定
		for i, r := range runes {
			y := startY + float64(i)*unitSize
			t.Texts = append(t.Texts, Text{
				x:       x - t.pdf.GetStringWidth(string(r))/2,
				y:       y + unitSize*0.9,
				text:    string(r),
				size:    t.fontSize,
				pageNum: t.pageNum,
			})
		}
	} else { // ページに収まっていない場合
		textH := float64(len(runes)) * unitSize
		bottom := t.GetBottomLine(t.Rows[0].pageNum)
		if bottom-t.y_i >= textH { // タイトルがページの下端より上に入る場合
			startY := t.y_i + (bottom-t.y_i-textH)/2
			page := t.Rows[0].pageNum
			// 塗りつぶし背景
			t.Rects = append(t.Rects, RectInfo{
				x:         t.x_i - t.titleW,
				y:         t.y_i,
				w:         t.titleW,
				h:         bottom - t.y_i,
				pageNum:   page,
				style:     "F", // 塗りつぶし
				LineWidth: 0.0,
			})
			// 枠線
			t.Rects = append(t.Rects, RectInfo{
				x:         t.x_i - t.titleW,
				y:         t.y_i,
				w:         t.titleW,
				h:         bottom - t.y_i,
				pageNum:   page,
				style:     "D", // 枠線
				LineWidth: 0.3,
			})
			// 一文字ずつ中央揃えで描画
			x := t.x_i - t.titleW + t.titleW/2 // 横は中央固定
			for i, r := range runes {
				y := startY + float64(i)*unitSize
				t.Texts = append(t.Texts, Text{
					x:       x - t.pdf.GetStringWidth(string(r))/2,
					y:       y + unitSize*0.9,
					text:    string(r),
					size:    t.fontSize,
					pageNum: page,
				})
			}
		} else if t.Ys[len(t.Ys)-1]-t.margin >= textH { // タイトルが次のページに入る場合
			startY := t.margin + (t.Ys[len(t.Ys)-1]-t.margin-textH)/2
			page := t.Rows[len(t.Rows)-1].pageNum // 次のページに移動
			// 塗りつぶし背景
			t.Rects = append(t.Rects, RectInfo{
				x:         t.x_i - t.titleW,
				y:         t.margin,
				w:         t.titleW,
				h:         t.Ys[len(t.Ys)-1] - t.margin,
				pageNum:   page,
				style:     "F", // 塗りつぶし
				LineWidth: 0.0,
			})
			// 枠線
			t.Rects = append(t.Rects, RectInfo{
				x:         t.x_i - t.titleW,
				y:         t.margin,
				w:         t.titleW,
				h:         t.Ys[len(t.Ys)-1] - t.margin,
				pageNum:   page,
				style:     "D", // 枠線
				LineWidth: 0.3,
			})
			// 一文字ずつ中央揃えで描画
			x := t.x_i - t.titleW + t.titleW/2 // 横は中央固定
			for i, r := range runes {
				y := startY + float64(i)*unitSize
				t.Texts = append(t.Texts, Text{
					x:       x - t.pdf.GetStringWidth(string(r))/2,
					y:       y + unitSize*0.9,
					text:    string(r),
					size:    t.fontSize,
					pageNum: page,
				})
			}
		} else { // タイトルがページに収まらない場合
			contanableLines := int((bottom - t.y_i) / unitSize)
			page1 := t.Rows[0].pageNum
			page2 := t.Rows[len(t.Rows)-1].pageNum
			for i := 0; i < contanableLines; i++ {
				r := runes[i]
				y := t.y_i + float64(i)*unitSize + unitSize*0.9
				x := t.x_i - t.titleW + t.titleW/2 // 横は中央固定
				t.Texts = append(t.Texts, Text{
					x:       x - t.pdf.GetStringWidth(string(r))/2,
					y:       y,
					text:    string(r),
					size:    t.fontSize,
					pageNum: page1,
				})
				// 塗りつぶし背景
				t.Rects = append(t.Rects, RectInfo{
					x:         t.x_i - t.titleW,
					y:         t.y_i,
					w:         t.titleW,
					h:         bottom - t.y_i,
					pageNum:   page1,
					style:     "F", // 塗りつぶし
					LineWidth: 0.0,
				})
				// 枠線
				t.Rects = append(t.Rects, RectInfo{
					x:         t.x_i - t.titleW,
					y:         t.y_i,
					w:         t.titleW,
					h:         bottom - t.y_i,
					pageNum:   page1,
					style:     "D", // 枠線
					LineWidth: 0.3,
				})
			}
			// 次のページに移動
			for i := contanableLines; i < len(runes); i++ {
				r := runes[i]
				y := t.margin + float64(i-contanableLines)*unitSize + unitSize*0.9
				x := t.x_i - t.titleW + t.titleW/2 // 横は中央固定
				t.Texts = append(t.Texts, Text{
					x:       x - t.pdf.GetStringWidth(string(r))/2,
					y:       y,
					text:    string(r),
					size:    t.fontSize,
					pageNum: page2,
				})
				// 塗りつぶし背景
				t.Rects = append(t.Rects, RectInfo{
					x:         t.x_i - t.titleW,
					y:         t.margin,
					w:         t.titleW,
					h:         t.y_f - t.margin,
					pageNum:   page2,
					style:     "F", // 塗りつぶし
					LineWidth: 0.0,
				})
				// 枠線
				t.Rects = append(t.Rects, RectInfo{
					x:         t.x_i - t.titleW,
					y:         t.margin,
					w:         t.titleW,
					h:         t.y_f - t.margin,
					pageNum:   page2,
					style:     "D", // 枠線
					LineWidth: 0.3,
				})
			}
		}
	}
}

func (t *Table) Render(outLine bool) {
	defer func() {
		if r := recover(); r != nil {
			log.Println("Recovered from panic:", r)

		}
	}()
	fmt.Printf("[Render] initialpageNum=%d, pageNum=%d\n", t.initialpageNum, t.pageNum)
	for i := t.initialpageNum; i <= t.pageNum; i++ {
		if i > t.pdf.PageNo() {
			fmt.Printf("[Render] AddPage: i=%d, current PageNo=%d\n", i, t.pdf.PageNo())
			t.pdf.AddPage()
		}
		fmt.Printf("[Render] Render Rects: %d, Cells: %d, Texts: %d on page: %d\n", len(t.Rects), len(t.Cells), len(t.Texts), i)
		for _, rect := range t.Rects {
			if rect.pageNum == i && rect.style == "F" {
				fmt.Printf("[Render] Rect(F): page=%d x=%.2f y=%.2f w=%.2f h=%.2f LineWidth=%.2f\n", rect.pageNum, rect.x, rect.y, rect.w, rect.h, rect.LineWidth)
				t.pdf.SetXY(rect.x, rect.y)
				t.pdf.SetLineWidth(rect.LineWidth)
				t.pdf.Rect(rect.x, rect.y, rect.w, rect.h, rect.style)
			}
		}
		for _, cell := range t.Cells {
			if cell.pageNum == i {
				fmt.Printf("[Render] Cell: page=%d x=%.2f y=%.2f w=%.2f h=%.2f text=%s fontSize=%.2f align=%s fill=%v\n", cell.pageNum, cell.x, cell.y, cell.w, cell.h, cell.text, cell.fontSize, cell.align, cell.fill)
				t.pdf.SetXY(cell.x, cell.y)
				t.pdf.SetFont(t.font, "", cell.fontSize)
				t.pdf.SetLineWidth(cell.LineWidth)
				t.pdf.CellFormat(cell.w, cell.h, cell.text, cell.border, 0, cell.align, cell.fill, 0, cell.link)
			}
		}
		for _, text := range t.Texts {
			if text.pageNum == i {
				fmt.Printf("[Render] Text: page=%d x=%.2f y=%.2f text=%s size=%.2f\n", text.pageNum, text.x, text.y, text.text, text.size)
				t.pdf.SetFont(t.font, "", text.size)
				t.pdf.Text(text.x, text.y, text.text)
			}
		}
		for _, rect := range t.Rects {
			if rect.pageNum == i && rect.style == "D" {
				fmt.Printf("[Render] Rect(D): page=%d x=%.2f y=%.2f w=%.2f h=%.2f LineWidth=%.2f\n", rect.pageNum, rect.x, rect.y, rect.w, rect.h, rect.LineWidth)
				t.pdf.SetLineWidth(rect.LineWidth)
				t.pdf.Rect(rect.x, rect.y, rect.w, rect.h, rect.style)
			}
		}
		if outLine {
			bottom := t.GetBottomLine(i)
			top := t.GetTopLine(i)
			fmt.Printf("[Render] Outer Rect: page=%d x=%.2f y=%.2f w=%.2f h=%.2f\n", i, t.x_i-t.titleW, top, t.x_f-t.x_i+t.titleW, bottom-top)
			t.pdf.SetLineWidth(0.3)
			if i == t.initialpageNum && bottom > 0.0 { // 初期ページで、全体が1ページに収まっている場合
				if bottom-top > 0.0 {
					t.pdf.Rect(t.x_i-t.titleW, top, t.x_f-t.x_i+t.titleW, bottom-top, "D")
				}
			} else {
				if bottom-t.margin > 0.0 {
					t.pdf.Rect(t.x_i-t.titleW, t.margin, t.x_f-t.x_i+t.titleW, bottom-t.margin, "D")
				}
			}
		}
	}
}

func CalcTextHeight(pdf *gofpdf.Fpdf, text string, width float64, lineHeight float64) float64 {
	lines := pdf.SplitLines([]byte(text), width)
	return float64(len(lines)) * lineHeight
}

func GetMaxChars(pdf *gofpdf.Fpdf, runes []rune, start int, width float64, fontSize float64) int {
	accumWidth := 0.0
	for i := start; i < len(runes); i++ {
		ch := string(runes[i])
		pdf.SetFontSize(fontSize) // フォントサイズを設定
		w := pdf.GetStringWidth(ch)
		if accumWidth+w > width-1 { // 0.5は余白調整
			return i - start
		}
		accumWidth += w
	}
	return len(runes) - start
}

func SplitByMaxChars(pdf *gofpdf.Fpdf, text string, width float64, fontSize float64) []string {
	var lines []string
	runes := []rune(text)
	returnCheck := false

	for i := 0; i < len(runes); {

		// リターンで改行されていない場合は行頭スペースをスキップ（全角・半角）。リターンで改行されている場合は、そのスペースはスタイル上恣意的なものである可能性が高いのでスペースをスキップしない。
		if !returnCheck {
			for i < len(runes) && (runes[i] == ' ' || runes[i] == '　') {
				i++
			}
		}

		if i >= len(runes) {
			break
		}

		// 改行が maxChars より前にある場合は優先して分割
		end := i
		for end < len(runes) && runes[end] != '\n' {
			end++
		}

		// 改行位置か、文字幅に収まる最大長で切る
		maxChars := GetMaxChars(pdf, runes, i, width, fontSize)
		if i+maxChars > end {
			maxChars = end - i
		}

		lines = append(lines, string(runes[i:i+maxChars]))

		i += maxChars
		if i < len(runes) && runes[i] == '\n' {
			i++                // 改行文字スキップ
			returnCheck = true // 改行があったので、次の行頭スペースはスキップしない
		} else {
			returnCheck = false // 改行がなかったので、次の行頭スペースはスキップする
		}
	}

	return lines
}

// xlsxファイルをPDFディレクトリに保存し、A1:AD48をgofpdfでPDF出力
func (a *App) SaveXLSXsToPDFDir(files []FileData) error {

	defer func() {
		if r := recover(); r != nil {
			log.Println("Recovered from panic:", r)
		}
	}()

	// フォントファイルの読み込み
	fontFile, fontPath, err := a.loadFont()
	if err != nil {
		return err
	}
	defer fontFile.Close()
	defer os.Remove(fontPath)

	Dpath, err := GetDownloadsPath()
	if err != nil {
		return fmt.Errorf("ダウンロードパスの取得に失敗: %w", err)
	}

	for _, f := range files {

		fx, _, err := a.loadCSV(f)
		if err != nil {
			return fmt.Errorf("CSVファイルの読み込みに失敗: %w", err)
		}
		defer fx.Close()

		pdfPath := filepath.Join(Dpath, "求人票_"+time.Now().Format("20060102")+".pdf")

		// PDF生成
		pdf := gofpdf.New("P", "mm", "A4", os.TempDir())
		baseName := filepath.Base(fontPath)
		pdf.AddUTF8Font("IPA", "", baseName)
		pdf.SetAutoPageBreak(false, 0.0) // 自動改ページを無効化
		pdf.AddPage()

		sheets := fx.GetSheetList()
		if len(sheets) == 0 {
			return fmt.Errorf("%s: シートがありません", f.Name)
		}

		for index, sheet := range sheets {
			// CSVファイルの読み込み
			tableData, err := a.loadData(sheet, fx)
			if err != nil {
				return err
			}

			if index != 0 {
				pdf.AddPage()
			}
			// MARGIN CONFIG
			marginSide := 30.0
			marginTop := 13.0

			// TITLE
			pdf.SetFont("IPA", "", 20)
			title := "求人票"
			titleW := pdf.GetStringWidth(title)
			_, titleH := pdf.GetFontSize()
			pageW, _ := pdf.GetPageSize()
			x := (pageW - titleW) / 2
			y := marginTop
			pdf.SetXY(x, y)
			pdf.CellFormat(titleW, titleH, title, "", 0, "L", false, 0, "")

			// COMPANY NAME
			pdf.SetFontSize(7.0)
			company := "株式会社アーリー・バード・エージェント"
			companyW := pdf.GetStringWidth(company)
			_, companyH := pdf.GetFontSize() // サイズ取得は不要だが、GetFontSize()を呼び出しておく
			x = pageW - companyW - marginSide
			y = marginTop
			pdf.SetXY(x, y)
			pdf.CellFormat(companyW, companyH, company, "", 0, "L", false, 0, "")

			// TABLE A
			pdf.SetFillColor(153, 204, 255)
			w := 5.0
			offsetH := 2.0
			pdf.SetXY(marginSide, marginTop+titleH)
			pdf.SetFontSize(6.0)

			rowNum := 1
			rowH := 2.5
			ft := 6.0 // フォントサイズ
			dh := 4.5 // デフォルトのセル高さ
			tableID := NewTable(pdf, marginSide+w, marginTop+titleH+offsetH-rowH, pageW-marginSide, marginTop+titleH+offsetH, 9, rowNum, "IPA", ft, dh, "1")
			tableID.SetCell(7, 0, 8, 1, tableData[2][24], "C", true, 5.0, "", 0.3, rowH)
			tableID.SetCell(8, 0, 9, 1, tableData[2][27], "C", false, 5.0, "", 0.3, rowH)
			tableID.Render(false)
			fmt.Print("[Render] completed tableID\n")
			rowNum = 8
			rowH = 4.0

			table := NewTable(pdf, marginSide+w, marginTop+titleH+offsetH, pageW-marginSide, marginTop+titleH+offsetH+rowH, 9, rowNum, "IPA", ft, dh, "1")

			lw := 0.1 // セルの枠線の太さ
			table.SetCell(1, 0, 6, 1, tableData[3][2], "L", false, 5.0, "", lw, 2.5)
			table.SetCell(1, 1, 6, 2, tableData[4][2], "L", false, 10.0, "", lw, 7.0)
			table.SetCell(0, 0, 1, 2, tableData[3][1], "C", true, -1.0, "", lw, rowH)

			table.SetCell(6, 0, 7, 2, tableData[3][21], "C", true, -1.0, "", lw, rowH)
			table.SetMultiRowCell(7, 0, 9, 2, tableData[3][24], "L", false, -1.0, false)

			table.SetCell(0, 2, 1, 3, tableData[5][1], "C", true, -1.0, "", lw, rowH)
			table.SetCell(1, 2, 3, 3, tableData[5][2], "L", false, -1.0, "", lw, rowH)
			table.SetCell(3, 2, 4, 3, tableData[5][10], "C", true, -1.0, "", lw, rowH)
			table.SetCell(4, 2, 6, 3, tableData[5][13], "L", false, -1.0, "", lw, rowH)
			table.SetCell(6, 2, 7, 3, tableData[5][21], "C", true, -1.0, "", lw, rowH)
			table.SetCell(7, 2, 9, 3, tableData[5][24], "L", false, -1.0, "", lw, rowH)

			table.SetCell(0, 3, 1, 4, tableData[6][1], "C", true, -1.0, "", lw, rowH)
			table.SetCell(1, 3, 3, 4, tableData[6][2], "L", false, -1.0, "", lw, rowH)
			table.SetCell(3, 3, 4, 4, tableData[6][10], "C", true, -1.0, "", lw, rowH)
			table.SetCell(4, 3, 9, 4, tableData[6][13], "L", false, -1.0, "", lw, rowH)

			table.SetCell(0, 4, 1, 5, tableData[7][1], "C", true, -1.0, "", lw, rowH)
			table.SetCell(1, 4, 9, 5, tableData[7][2], "L", false, -1.0, "", lw, rowH)

			table.SetMultiRowCell(1, 5, 9, 6, tableData[8][2], "L", false, -1.0, false)
			table.SetCell(0, 5, 1, 6, tableData[8][1], "C", true, -1.0, "", lw, rowH)

			table.SetMultiRowCell(1, 6, 9, 7, tableData[9][2], "L", false, -1.0, true)
			table.SetCell(0, 6, 1, 7, tableData[9][1], "C", true, -1.0, "", lw, rowH)

			table.SetMultiRowCell(1, 7, 9, 8, tableData[10][2], "L", false, -1.0, true)
			table.SetCell(0, 7, 1, 8, tableData[10][1], "C", true, -1.0, "", lw, rowH)

			table.SetTitle(tableData[3][0])
			table.Render(true)

			// TABLE B
			pdf.SetLineWidth(0.1)
			currentH := table.Ys[len(table.Ys)-1] + offsetH
			rowNum = 10
			rowH = 4.5

			table2 := NewTable(pdf, marginSide+w, currentH, pageW-marginSide, currentH+rowH, 9, rowNum, "IPA", ft, dh, "1")

			table2.SetCell(0, 0, 1, 1, tableData[12][1], "C", true, -1.0, "", lw, rowH)
			table2.SetCell(1, 0, 9, 1, tableData[12][2], "L", false, -1.0, "", lw, rowH)

			table2.SetMultiRowCell(1, 1, 9, 2, tableData[13][2], "L", false, -1.0, true)
			table2.SetCell(0, 1, 1, 2, tableData[13][1], "C", true, -1.0, "", lw, rowH)

			table2.SetCell(0, 2, 1, 4, tableData[14][1], "C", true, -1.0, "", lw, rowH*2)
			table2.SetCell(1, 2, 3, 4, tableData[14][2], "L", false, -1.0, "", lw, rowH*2)
			table2.SetCell(3, 2, 4, 3, tableData[14][10], "C", true, -1.0, "", lw, rowH)
			table2.SetCell(4, 2, 6, 3, tableData[14][13], "L", false, -1.0, "", lw, rowH)
			table2.SetCell(6, 2, 7, 3, tableData[14][21], "C", true, -1.0, "", lw, rowH)
			table2.SetCell(7, 2, 9, 3, tableData[14][24], "L", false, -1.0, "", lw, rowH)

			table2.SetCell(3, 3, 4, 4, tableData[15][10], "C", true, -1.0, "", lw, rowH)
			table2.SetCell(4, 3, 9, 4, tableData[15][13], "L", false, -1.0, "", lw, rowH)

			table2.SetMultiRowCell(1, 4, 9, 5, tableData[16][2], "L", false, -1.0, false)
			table2.SetCell(0, 4, 1, 5, tableData[16][1], "C", true, -1.0, "", lw, rowH)

			table2.SetMultiRowCell(1, 5, 9, 6, tableData[17][2], "L", false, -1.0, false)
			table2.SetCell(0, 5, 1, 6, tableData[17][1], "C", true, -1.0, "", lw, rowH)

			table2.SetMultiRowCell(1, 6, 9, 7, tableData[18][2], "L", false, -1.0, false)
			table2.SetCell(0, 6, 1, 7, tableData[18][1], "C", true, -1.0, "", lw, rowH)

			table2.SetCell(0, 7, 1, 8, tableData[19][1], "C", true, -1.0, "", lw, rowH)
			table2.SetCell(1, 7, 3, 8, tableData[19][2], "L", false, -1.0, "", lw, rowH)
			table2.SetCell(3, 7, 4, 8, tableData[19][10], "C", true, -1.0, "", lw, rowH)
			table2.SetCell(4, 7, 9, 8, tableData[19][13], "L", false, -1.0, "", lw, rowH)

			table2.SetCell(0, 8, 1, 9, tableData[20][1], "C", true, -1.0, "", lw, rowH)
			table2.SetCell(1, 8, 3, 9, tableData[20][2], "L", false, -1.0, "", lw, rowH)
			table2.SetCell(3, 8, 4, 9, tableData[20][10], "C", true, -1.0, "", lw, rowH)
			table2.SetCell(4, 8, 6, 9, tableData[20][13], "L", false, -1.0, "", lw, rowH)
			table2.SetCell(6, 8, 7, 9, tableData[20][21], "C", true, -1.0, "", lw, rowH)
			table2.SetCell(7, 8, 9, 9, tableData[20][24], "L", false, -1.0, "", lw, rowH)

			table2.SetMultiRowCell(1, 9, 9, 10, tableData[21][2], "L", false, -1.0, true)
			table2.SetCell(0, 9, 1, 10, tableData[21][1], "C", true, -1.0, "", lw, rowH)

			table2.SetTitle(tableData[12][0])
			table2.Render(true)

			// TABLE C
			pdf.SetLineWidth(0.1)
			currentH = table2.Ys[len(table2.Ys)-1] + offsetH
			rowNum = 2

			table3 := NewTable(pdf, marginSide+w, currentH, pageW-marginSide, currentH+rowH, 9, rowNum, "IPA", ft, dh, "1")

			table3.SetCell(0, 0, 1, 1, tableData[23][1], "C", true, -1.0, "", lw, rowH)
			table3.SetCell(1, 0, 3, 1, tableData[23][2], "L", false, -1.0, "", lw, rowH)
			table3.SetCell(3, 0, 4, 1, tableData[23][10], "C", true, -1.0, "", lw, rowH)
			table3.SetCell(4, 0, 9, 1, tableData[23][13], "L", false, -1.0, "", lw, rowH)

			table3.SetMultiRowCell(1, 1, 9, 2, tableData[24][2], "L", false, -1.0, true)
			table3.SetMultiRowCell(0, 1, 1, 2, tableData[24][1], "C", true, -1.0, true)

			table3.SetTitle(tableData[23][0])
			table3.Render(true)

			// TABLE D
			pdf.SetLineWidth(0.1)
			currentH = table3.Ys[len(table3.Ys)-1] + offsetH
			rowNum = 6

			table4 := NewTable(pdf, marginSide+w, currentH, pageW-marginSide, currentH+rowH, 9, rowNum, "IPA", ft, dh, "1")

			table4.SetCell(0, 0, 1, 1, tableData[26][1], "C", true, -1.0, "", lw, rowH)
			table4.SetCell(1, 0, 3, 1, tableData[26][2], "L", false, -1.0, "", lw, rowH)
			table4.SetCell(3, 0, 4, 1, tableData[26][10], "C", true, -1.0, "", lw, rowH)
			table4.SetCell(4, 0, 6, 1, tableData[26][13], "L", false, -1.0, "", lw, rowH)
			table4.SetCell(6, 0, 7, 1, tableData[26][21], "C", true, -1.0, "", lw, rowH)
			table4.SetCell(7, 0, 9, 1, tableData[26][24], "L", false, -1.0, "", lw, rowH)

			table4.SetCell(0, 1, 1, 2, tableData[27][1], "C", true, -1.0, "", lw, rowH)
			table4.SetCell(1, 1, 3, 2, tableData[27][2], "L", false, -1.0, "", lw, rowH)
			table4.SetCell(3, 1, 4, 2, tableData[27][10], "C", true, -1.0, "", lw, rowH)
			table4.SetCell(4, 1, 6, 2, tableData[27][13], "L", false, -1.0, "", lw, rowH)
			table4.SetCell(6, 1, 7, 2, tableData[27][21], "C", true, -1.0, "", lw, rowH)
			table4.SetCell(7, 1, 9, 2, tableData[27][24], "L", false, -1.0, "", lw, rowH)

			table4.SetMultiRowCell(1, 2, 9, 3, tableData[28][2], "L", false, -1.0, true)
			table4.SetCell(0, 2, 1, 3, tableData[28][1], "C", true, -1.0, "", lw, rowH)

			table4.SetMultiRowCell(1, 3, 9, 4, tableData[29][2], "L", false, -1.0, true)
			table4.SetCell(0, 3, 1, 4, tableData[29][1], "C", true, -1.0, "", lw, rowH)

			table4.SetMultiRowCell(1, 4, 9, 5, tableData[30][2], "L", false, -1.0, true)
			table4.SetCell(0, 4, 1, 5, tableData[30][1], "C", true, -1.0, "", lw, rowH)

			table4.SetMultiRowCell(1, 5, 6, 6, tableData[31][2], "L", false, -1.0, true)
			table4.SetCell(0, 5, 1, 6, tableData[31][1], "C", true, -1.0, "", lw, rowH)
			table4.SetCell(6, 5, 7, 6, tableData[31][21], "C", true, -1.0, "", lw, rowH)
			table4.SetCell(7, 5, 9, 6, tableData[31][24], "L", false, -1.0, "", lw, rowH)

			table4.SetTitle(tableData[26][0])
			table4.Render(true)

			// TABLE E
			pdf.SetLineWidth(0.1)
			currentH = table4.Ys[len(table4.Ys)-1] + offsetH
			rowNum = 3

			table5 := NewTable(pdf, marginSide+w, currentH, pageW-marginSide, currentH+rowH, 9, rowNum, "IPA", ft, dh, "1")

			table5.SetMultiRowCell(1, 0, 6, 1, tableData[33][2], "L", false, -1.0, false)
			table5.SetCell(0, 0, 1, 1, tableData[33][1], "C", true, -1.0, "", lw, rowH)
			table5.SetCell(6, 0, 7, 1, tableData[33][21], "C", true, -1.0, "", lw, rowH)
			table5.SetCell(7, 0, 9, 1, tableData[33][24], "L", false, -1.0, "", lw, rowH)

			table5.SetMultiRowCell(1, 1, 9, 2, tableData[34][2], "L", false, -1.0, false)
			table5.SetCell(0, 1, 1, 2, tableData[34][1], "C", true, -1.0, "", lw, rowH)

			table5.SetMultiRowCell(1, 2, 9, 3, tableData[35][2], "L", false, -1.0, false)
			table5.SetCell(0, 2, 1, 3, tableData[35][1], "C", true, -1.0, "", lw, rowH)

			table5.SetTitle(tableData[33][0])
			table5.Render(true)

			// TABLE F
			currentH = table5.Ys[len(table5.Ys)-1] + offsetH
			rowNum = 1

			table6 := NewTable(pdf, marginSide+w, currentH, pageW-marginSide, currentH+rowH, 9, rowNum, "IPA", ft, dh, "1")

			table6.SetMultiRowCell(1, 0, 9, 1, tableData[37][2], "L", false, -1.0, false)
			table6.SetCellWithTitle(0, 0, 1, 1, tableData[37][0], "C", true, -1.0)

			table6.Render(false)

			// // Appendix
			currentH = table6.Ys[len(table6.Ys)-1] + offsetH

			table7 := NewAppendix(pdf, marginSide, pageW-marginSide, currentH, "IPA", ft, dh, "0")
			table7.SetAppendix(tableData[41][0], "L", false, -1.0, true)

			table7.Render(false)

			if index == 0 && len(sheets) == 1 {
				pdfPath = filepath.Join(Dpath, "求人票_"+tableData[4][2]+"_"+tableData[12][2]+".pdf")
			}
		}

		err = pdf.OutputFileAndClose(pdfPath)
		if err != nil {
			return fmt.Errorf("%s のPDF出力に失敗: %w", f.Name, err)
		}
		fmt.Printf("PDFファイルを保存しました: %s\n", pdfPath)
	}
	return nil
}
