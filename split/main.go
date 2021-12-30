package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"path"
	"path/filepath"
	"runtime"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func main() {
	start()
	// confirm("测试")
	// test()
	showMemory()
}

func getDir() string {
	file := ""
	files := GetFilesFromDir(".", ".xlsx")
	if len(files) > 0 {
		file = files[0]
	}
input1:
	fmt.Println("请输入文件名. 默认为" + file)
	fmt.Scanln(&file)
	_, err := os.Stat(file)
	if err != nil {
		fmt.Println("文件不存在" + file)
		file = ""
		goto input1
	}
	return file
}
func getNeedHead() int {
	res := 1
	need := "Y"
	fmt.Println("是否保留表头? Y/N 默认Y")
	fmt.Scanln(&need)
	if need == "n" || need == "N" {
		res = 0
	}
	return res
}
func getIndexRow() int {
input2:
	res := 1
	rowIndex := "1"
	fmt.Println("请选择分页列?请输入ABCD表示列  默认A")
	fmt.Scanln(&rowIndex)
	res, err := excelize.ColumnNameToNumber(rowIndex)
	if nil != err {
		fmt.Println("请输入有效的列")
		goto input2
	}
	return res - 1
}
func confirm() {
	key := ""
	fmt.Println("请输入回车确认,或输入ctrl+c终止")
	fmt.Scanln(&key)
}
func start() {
	file := getDir()
	needHead := getNeedHead()
	indexRow := getIndexRow()
	data := Read(file)
	files := split(data, indexRow, needHead)
	fmt.Println("结果已保存至:")
	fmt.Println(files)
}
func showMemory() {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	// fmt.Printf("%+v\n", m)
	fmt.Printf("os %dM\n", m.Sys/1024/1024)
}

func split(data [][]string, indexRow int, needHead int) []string {
	count := len(data)
	if count == 0 {
		fmt.Println("表格内容为空")
		return nil
	}
	var title []string
	if needHead == 1 {
		title = data[0]
	}
	excels := make(map[string][][]string)
	for i := needHead; i < count; i++ {
		row := data[i]
		capital, ok := excels[row[indexRow]]
		if !ok {
			capital = make([][]string, 0)
			if needHead == 1 {
				capital = append(capital, title)
			}
		}
		excels[row[indexRow]] = append(capital, row)
	}
	if needHead == 1 {
		fmt.Println("表头为:")
		fmt.Println(title)
	}
	fmt.Println("索引列为:" + string(title[indexRow]))
	fmt.Println("预计生成" + strconv.Itoa(len(excels)) + "个文件")
	confirm()
	files := make([]string, 0)
	for key, excel := range excels {
		filename := Write(excel, key, "")
		files = append(files, filename)
	}
	return files
}

func convertToInterface(arr []string) []interface{} {
	res := make([]interface{}, 0)
	for _, v := range arr {
		res = append(res, v)
	}
	return res
}
func Write(data [][]string, filename string, p string) string {
	f := excelize.NewFile()
	streamWriter, err := f.NewStreamWriter("Sheet1")
	if err != nil {
		fmt.Println(err)
		return ""
	}
	for i, row := range data {
		cell, _ := excelize.CoordinatesToCellName(1, i+1)
		if err := streamWriter.SetRow(cell, convertToInterface(row)); err != nil {
			fmt.Println(err)
		}
	}
	streamWriter.Flush()
	if p == "" {
		p, _ = GetCurrentPath()
	}
	result := p + "/" + filename + ".xlsx"

	if err := f.SaveAs(result); err != nil {
		fmt.Println(err)
	}
	return result
}

func Read(filepath string) [][]string {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		fmt.Println(err)
		return nil
	}
	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil
	}
	rows, err := f.GetRows(sheets[0])
	if err != nil {
		fmt.Println(err)
		return nil
	}
	return rows

}

// 获取当前路径
func GetCurrentPath() (string, error) {
	// path, err := os.Executable()
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))
	if err != nil {
		fmt.Println(err)
		return "", err
	}
	// dir := filepath.Dir(path)
	return dir, nil
}

//根据路径找到该路径下的所有文件，并返回
func GetFilesFromDir(dirpath string, filter string) []string {
	var files = make([]string, 0, 100)

	//从目录path下找到所有的目录文件
	allDir, err := ioutil.ReadDir(dirpath)
	//如果有错误发生就返回nil，不再进行查找
	if err != nil {
		return files
	}
	//遍历获取到的每一个目录信息
	for _, dir := range allDir {
		if dir.IsDir() {
			//如果还是目录，就继续向下进行递归查找,并追加到返回切片中
			GetFilesFromDir(dirpath+"/"+dir.Name(), filter)
			continue
		}
		filetype := path.Ext(path.Base(dir.Name()))
		if filetype != filter {
			continue
		}
		//如果不是目录,就读取该文件并加入到files中
		fileName := dirpath + "/" + dir.Name()
		files = append(files, fileName)
	}
	return files
}
