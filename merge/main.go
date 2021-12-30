package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"path"
	"path/filepath"
	"runtime"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	start()
	showMemory()
}

func getDir() string {
input1:
	inputdir := ""
	dir := "./"
	fmt.Println("请输入路径： 默认为当前路径" + dir)
	fmt.Scanln(&inputdir)
	if inputdir == "" {
		inputdir = dir
	}
	_, err := os.Stat(inputdir)
	if err != nil {
		fmt.Println("无效的路径" + inputdir)
		inputdir = ""
		goto input1
	}
	return inputdir
}
func getNeedHead() int {
	res := 1
	need := "Y"
	fmt.Println("是否过滤表头? Y/N 默认Y")
	fmt.Scanln(&need)
	if need == "n" || need == "N" {
		res = 0
	}
	return res
}
func start() {
	dir := getDir()
	needHead := getNeedHead()

	result := GetFilesFromDir(dir, ".xlsx")

	file := merge(result, needHead)
	fmt.Println("结果已保存至:" + file)
}
func showMemory() {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	// fmt.Printf("%+v\n", m)
	fmt.Printf("os %dM\n", m.Sys/1024/1024)
}

func merge(files []string, needHead int) string {
	f := excelize.NewFile()
	streamWriter, err := f.NewStreamWriter("Sheet1")
	if err != nil {
		fmt.Println(err)
	}
	rowIndex := 1
	for i, file := range files {
		data := Read(file)
		if nil == data {
			continue
		}
		for j, row := range data {
			if needHead == 1 && j == 0 && i > 0 {
				// 只有第一张表的第一行会被写入,其他表的第一行都跳过
				continue
			}
			cell, _ := excelize.CoordinatesToCellName(1, rowIndex)
			// fmt.Println(cell, convertToInterface(row))
			if err := streamWriter.SetRow(cell, convertToInterface(row)); err != nil {
				fmt.Println(err)
			}
			rowIndex++
		}
	}
	streamWriter.Flush()
	result := setResultPath(".")
	f.SaveAs(result)
	return result
}

func convertToInterface(arr []string) []interface{} {
	res := make([]interface{}, 0)
	for _, v := range arr {
		res = append(res, v)
	}
	return res
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

// 设置输出路径
func setResultPath(path string) string {
	currentTime := time.Now().Format("2006-01-02-15-04-05")
	return "./" + currentTime + "result.xlsx"
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
