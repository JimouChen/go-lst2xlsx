# go-lst2xlsx

- A efficient  tool implemented by go to convert lists to excel/列表转excel的高效工具


<div align="center">

<img width="360" alt="image" src="https://github.com/user-attachments/assets/b7284537-09a8-42c3-986f-f29013407f45">

</div>

## 使用场景

- 一个json文件(./data/lst.json)，内容如下：

```json
[
  {
    "name": "Tom",
    "age": 22,
    "addr": "A City"
  },
  {
    "name": "Peter",
    "age": 42,
    "addr": "B City"
  },
  {
    "name": "Gogo",
    "age": 18,
    "addr": "C City"
  }
]
```

- 要把上述的数据转成excel，并且可以指定列名的顺序。json文件可能非常大，转换后效果如下：

<img width="232" alt="image" src="https://github.com/user-attachments/assets/57847081-b1e9-4621-abf2-e15faa71479c">


## 使用例子

### 开箱即用(MacOS环境)

- 帮助说明
```bash
❯ ./lst2xlsx_mac -h
Usage of ./lst2xlsx:
  -jp string
        Path to the JSON file
  -ord string
        Column order
  -s string
        Sheet name (default "Sheet1")
  -sp string
        Path to save the Excel file (default "1730988813.xlsx")
```

- 转excel
```bash
./lst2xlsx_mac -jp ./data/lst.json -sp ./data/data1.xlsx -s sheet2 -ord '["name", "addr", "age"]'
```

### 自行编译后使用

#### 本地环境MacOS
```bash
go build -ldflags="-s -w" -o lst2xlsx_mac main.go
```

#### Windows
```bash
GOOS=windows GOARCH=amd64 go build -ldflags="-s -w" -o lst2xlsx_win.exe main.go
```

#### Linux
```bash
GOOS=linux GOARCH=amd64 go build -ldflags="-s -w" -o lst2xlsx_linux main.go
```
