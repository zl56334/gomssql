package main

import (
	"database/sql"
	"fmt"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	_ "github.com/mattn/go-adodb"
	"github.com/urfave/cli"
)

type Mssql struct {
	*sql.DB
	dataSource string
	database   string
	windows    bool
	sa         SA
}

type SA struct {
	user   string
	passwd string
	port   int
}

var help = func() {
	fmt.Println("催记系统辅助工具")
	fmt.Println("-o updateSeNoByCaseId 执行催收系统人员变更快捷操作，commonCase 处理公债，exchangeCase 换手，默认为updateSeNoByCaseId")
	fmt.Println("-f 对应导入文件路径")
	fmt.Println("类似格式为 ‘.\\TestSqlServer.exe -o commonCase -f c:\\test.xlsx’")
}

func (m *Mssql) Open() (err error) {
	var conf []string
	conf = append(conf, "Provider=SQLOLEDB")
	conf = append(conf, "Data Source="+m.dataSource)
	conf = append(conf, "Initial Catalog="+m.database)

	// Integrated Security=SSPI 这个表示以当前WINDOWS系统用户身去登录SQL SERVER服务器
	// (需要在安装sqlserver时候设置)，
	// 如果SQL SERVER服务器不支持这种方式登录时，就会出错。
	if m.windows {
		conf = append(conf, "integrated security=SSPI")
	} else {
		conf = append(conf, "user id="+m.sa.user)
		conf = append(conf, "password="+m.sa.passwd)
		conf = append(conf, "port="+fmt.Sprint(m.sa.port))
	}

	m.DB, err = sql.Open("adodb", strings.Join(conf, ";"))
	if err != nil {
		return err
	}
	return nil
}

func DoQuery(sqlInfo string, args ...interface{}) ([]map[string]interface{}, error) {
	db := Mssql{
		// 如果数据库是默认实例（MSSQLSERVER）则直接使用IP，命名实例需要指明。
		dataSource: "10.27.42.9",
		database:   "ccds",
		// windows: true 为windows身份验证，false 必须设置sa账号和密码
		windows: false,
		sa: SA{
			user:   "sa",
			passwd: "sa123abc()",
			port:   1433,
		},
	}
	// 连接数据库
	err := db.Open()
	if err != nil {
		fmt.Println("sql open:", err)
		return nil, err
	}

	rows, err := db.Query(sqlInfo, args...)
	if err != nil {
		return nil, err
	}
	columns, _ := rows.Columns()
	columnLength := len(columns)
	cache := make([]interface{}, columnLength) //临时存储每行数据
	for index := range cache {                 //为每一列初始化一个指针
		var a interface{}
		cache[index] = &a
	}
	var list []map[string]interface{} //返回的切片
	list = make([]map[string]interface{}, 0)
	for rows.Next() {
		_ = rows.Scan(cache...)

		item := make(map[string]interface{})
		for i, data := range cache {
			item[columns[i]] = *data.(*interface{}) //取实际类型
		}
		list = append(list, item)
	}
	_ = rows.Close()
	db.Close()
	fmt.Println("doq list:", list)
	return list, nil
}

func DoExec(sqlInfo string, args ...interface{}) (int64, error) {
	db := Mssql{
		// 如果数据库是默认实例（MSSQLSERVER）则直接使用IP，命名实例需要指明。
		dataSource: "10.27.42.9",
		database:   "ccds",
		// windows: true 为windows身份验证，false 必须设置sa账号和密码
		windows: false,
		sa: SA{
			user:   "sa",
			passwd: "sa123abc()",
			port:   1433,
		},
	}
	// 连接数据库
	err := db.Open()
	if err != nil {
		fmt.Println("sql open:", err)
		return 0, err
	}

	rows, err := db.Exec(sqlInfo, args...)
	if err != nil {
		return 0, err
	}
	db.Close()
	return rows.RowsAffected()
}

func readXlsx(fpath string) [][]string {
	f, err := excelize.OpenFile(fpath)
	if err != nil {
		fmt.Println(err)
	}
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows("Sheet1")
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		// fmt.Println()
	}
	return rows
}

func doUpdateSeNoByCaseId(cellRows [][]string) {
	fmt.Print(cellRows)
	if cellRows[0][0] != "案件ID" {
		fmt.Println("未找到 案件ID 列")
		return
	}
	if cellRows[0][1] != "员工ID" {
		fmt.Println("未找到 员工ID 列")
		return
	}

	db := Mssql{
		// 如果数据库是默认实例（MSSQLSERVER）则直接使用IP，命名实例需要指明。
		dataSource: "10.27.42.9",
		database:   "ccds",
		// windows: true 为windows身份验证，false 必须设置sa账号和密码
		windows: false,
		sa: SA{
			user:   "sa",
			passwd: "sa123abc()",
			port:   1433,
		},
	}
	// 连接数据库
	err := db.Open()
	if err != nil {
		fmt.Println("sql open:", err)
		return
	}
	defer db.Close()
	for k, v := range cellRows {
		if k != 0 {
			sql := "UPDATE bank_case set cas_se_no = '#{cas_se_no}' where cas_id = '#{cas_id}'"
			sql = strings.Replace(sql, "#{cas_id}", v[0], -1)
			sql = strings.Replace(sql, "#{cas_se_no}", v[1], -1)
			fmt.Println(sql)
			rows, err := db.Query(sql)

			if err != nil {
				fmt.Println("query: ", err)
				fmt.Println(rows)
				return
			} else {
				fmt.Println("已完成：", k)
			}
		} else {
			fmt.Println("开始执行更改")
		}

	}
}

func doCommonCase(cellRows [][]string) {
	fmt.Print(cellRows)
	if cellRows[0][0] != "批次号" {
		fmt.Println("未找到 批次号 列")
		return
	}
	for k, v := range cellRows {
		if k != 0 {
			sql := "SELECT cbat_id FROM case_bat WHERE cbat_code = '#{cbat_code}'"
			sql = strings.Replace(sql, "#{cbat_code}", v[0], -1)
			fmt.Println(sql)
			rows, err := DoQuery(sql)

			for i := 0; i < len(rows); i++ {
				data := rows[i]
				fmt.Println("cbat_id:", data["cbat_id"])
				cbat_id := fmt.Sprint((data["cbat_id"].(int64)))

				sql = "SELECT cas_id,cas_code,cas_m,cas_se_no,cas_num FROM bank_case WHERE cas_cbat_id = '#{cas_cbat_id}'"
				sql = strings.Replace(sql, "#{cas_cbat_id}", cbat_id, -1)

				fmt.Println("sql:", sql)

				caseRows, caseerr := DoQuery(sql)

				for j := 0; j < len(caseRows); j++ {
					caseRow := caseRows[j]
					fmt.Println("cas_id:", caseRow["cas_id"])
					fmt.Println("cas_code:", caseRow["cas_code"])
					fmt.Println("cas_m:", caseRow["cas_m"])
					fmt.Println("cas_se_no:", caseRow["cas_se_no"])
					for l := 0; l < len(caseRows); l++ {
						caseRowT := caseRows[l]
						if caseRowT["cas_id"] != caseRow["cas_id"] && caseRowT["cas_se_no"] != caseRow["cas_se_no"] && caseRowT["cas_num"] == caseRow["cas_num"] {
							upsql := "UPDATE bank_case SET cas_se_no = '#{cas_se_no}' WHERE cas_id = '#{cas_id}'"
							if caseRowT["cas_m"].(float64) < caseRow["cas_m"].(float64) {
								upsql = strings.Replace(sql, "#{cas_se_no}", fmt.Sprint((caseRow["cas_se_no"])), -1)
								upsql = strings.Replace(sql, "#{cas_id}", fmt.Sprint((caseRowT["cas_id"])), -1)

								upCaseRows, upCaseerr := DoExec(upsql)
								if upCaseerr != nil {
									fmt.Println("query: ", upCaseerr)
									fmt.Println(upCaseRows)
									return
								}

								fmt.Println("UPDATE sql:", sql)
							} else if caseRowT["cas_m"].(float64) >= caseRow["cas_m"].(float64) {
								upsql = strings.Replace(sql, "#{cas_se_no}", fmt.Sprint((caseRowT["cas_se_no"])), -1)
								upsql = strings.Replace(sql, "#{cas_id}", fmt.Sprint((caseRow["cas_id"])), -1)

								upCaseRows, upCaseerr := DoExec(sql)
								if upCaseerr != nil {
									fmt.Println("query: ", upCaseerr)
									fmt.Println(upCaseRows)
									return
								}

								fmt.Println("UPDATE sql:", sql)
							}
						}
					}
				}

				if caseerr != nil {
					fmt.Println("query: ", caseerr)
					fmt.Println(rows)
					return
				}

			}
			if err != nil {
				fmt.Println("query: ", err)
				fmt.Println(rows)
				return
			} else {
				fmt.Println("已完成：", k)
			}
		} else {
			fmt.Println("开始执行更改")
		}

	}

}

func doExchangeCase(cellRows [][]string) {
	fmt.Println(cellRows)
	if cellRows[0][0] != "员工ID" {
		fmt.Println("未找到 员工ID 列")
		return
	}
	if cellRows[0][1] != "批次" {
		fmt.Println("未找到 批次 列")
		return
	}
	if cellRows[0][2] != "是否留案" {
		fmt.Println("未找到 是否留案 列")
		return
	}
	if cellRows[0][3] != "留案Code" {
		fmt.Println("未找到 留案Code 列")
		return
	}
	var stayCaseCodeListForAll []string
	casSeNoForAll := "'0'"
	stayCaseCodeStringForAll := "'0'"
	for k, v := range cellRows {
		if k != 0 {
			if v[2] == "是" {
				stayCaseCodes := strings.Fields(v[3])
				stayCaseCodeListForAll = append(stayCaseCodeListForAll, stayCaseCodes...)
				for _, stayCaseCode := range stayCaseCodes {
					stayCaseCodeStringForAll = stayCaseCodeStringForAll + ",'" + stayCaseCode + "'"
				}
			}
			casSeNoForAll = casSeNoForAll + ",'" + v[0] + "'"
		}
	}

	for k, v := range cellRows {
		if k != 0 {
			sql := "SELECT bc.cas_id,  bc.cas_code,  bc.cas_se_no FROM dbo.bank_case AS bc INNER JOIN dbo.case_bat AS cb ON  bc.cas_cbat_id = cb.cbat_id WHERE  bc.cas_se_no = '#{cas_se_no}' AND cb.cbat_code = '#{cbat_code}' AND bc.cas_code NOT IN (#{stayCaseCodeStringForAll})"
			sql = strings.Replace(sql, "#{cas_se_no}", v[0], -1)
			sql = strings.Replace(sql, "#{cbat_code}", v[1], -1)
			sql = strings.Replace(sql, "#{stayCaseCodeStringForAll}", stayCaseCodeStringForAll, -1)
			fmt.Println(sql)
			rows, err := DoQuery(sql)
			fmt.Println("reRowsNum: ", len(rows))

			if err != nil {
				fmt.Println("query: ", err)
				fmt.Println(rows)
				return
			}
			changeNum := len(rows)
			cellRows[k] = append(cellRows[k], fmt.Sprint(changeNum))

			fmt.Println(cellRows[k])
		}
	}

	completeCaseList := "'0'"
	for k, v := range cellRows {
		if k != 0 && v[4] != "0" {
			changeCsaIdSql := "SELECT TOP(#{changeNum}) bc.cas_id FROM bank_case as bc INNER JOIN case_bat as cb ON bc.cas_cbat_id = cbat_id WHERE cas_se_no <> '#{cas_se_no}' AND cas_se_no IN (#{in_cas_se_no}) AND cb.cbat_code = '#{cbat_code}' AND bc.cas_code NOT IN (#{not_in_cas_codes}) AND bc.cas_id NOT IN (#{not_in_cas_id})"
			changeCsaIdSql = strings.Replace(changeCsaIdSql, "#{changeNum}", v[4], -1)
			changeCsaIdSql = strings.Replace(changeCsaIdSql, "#{cas_se_no}", v[0], -1)
			changeCsaIdSql = strings.Replace(changeCsaIdSql, "#{cbat_code}", v[1], -1)
			changeCsaIdSql = strings.Replace(changeCsaIdSql, "#{not_in_cas_codes}", stayCaseCodeStringForAll, -1)
			changeCsaIdSql = strings.Replace(changeCsaIdSql, "#{not_in_cas_id}", completeCaseList, -1)
			changeCsaIdSql = strings.Replace(changeCsaIdSql, "#{in_cas_se_no}", casSeNoForAll, -1)
			fmt.Println(changeCsaIdSql)

			changeCsaIdRows, changeCsaIderr := DoQuery(changeCsaIdSql)
			fmt.Println("changeCsaIdRows :", changeCsaIdRows)
			for i := 0; i < len(changeCsaIdRows); i++ {
				data := changeCsaIdRows[i]
				fmt.Println("data :", data)
				fmt.Println("data['cas_id'] :", data["cas_id"])
				updateSql := "UPDATE bank_case SET cas_se_no = '#{cas_se_no}' WHERE cas_id = '#{cas_id}'"
				updateSql = strings.Replace(updateSql, "#{cas_se_no}", fmt.Sprint(v[0]), -1)
				updateSql = strings.Replace(updateSql, "#{cas_id}", fmt.Sprint(data["cas_id"]), -1)

				fmt.Println("sql:", updateSql)

				updateRows, caseerr := DoExec(updateSql)

				if caseerr != nil {
					fmt.Println("query: ", caseerr)
					return
				}
				fmt.Println(updateRows)
				completeCaseList = completeCaseList + ",'" + fmt.Sprint(data["cas_id"]) + "'"
				fmt.Println("completeCaseList:", completeCaseList)
			}
			if changeCsaIderr != nil {
				fmt.Println("query: ", changeCsaIderr)
				fmt.Println(changeCsaIdRows)
				return
			}
		}
	}
}

func main() {
	app := cli.NewApp()
	app.Name = "催收辅助"
	app.Usage = "辅助催收系统调整数据"
	app.Version = "0.1.0"

	app.Flags = []cli.Flag{
		cli.StringFlag{
			Name:  "operation, o",
			Value: "updateSeNoByCaseId",
			Usage: "operation",
		},
		cli.StringFlag{
			Name:  "fPath, f",
			Value: "fPath",
			Usage: "fPath for operation",
		},
	}

	app.Action = func(c *cli.Context) error {
		operation := c.String("operation")
		fPath := c.String("fPath")
		if operation == "updateSeNoByCaseId" {
			rows := readXlsx(fPath)
			fmt.Println("Result rows:", rows)
			doUpdateSeNoByCaseId(rows)
		} else if operation == "commonCase" {
			rows := readXlsx(fPath)
			fmt.Println("Result rows:", rows)
			doCommonCase(rows)
		} else if operation == "exchangeCase" {
			rows := readXlsx(fPath)
			fmt.Println("Result rows:", rows)
			doExchangeCase(rows)
		} else {
			help()
		}
		return nil
	}

	app.Run(os.Args)
}
