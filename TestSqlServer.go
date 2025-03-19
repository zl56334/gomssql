package main

import (
	"database/sql"
	"fmt"
	"os"
	"strconv"

	"strings"

	_ "github.com/mattn/go-adodb"
	"github.com/urfave/cli"
	"github.com/xuri/excelize/v2"
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

// var db = Mssql{
// 	// 如果数据库是默认实例（MSSQLSERVER）则直接使用IP，命名实例需要指明。
// 	dataSource: "192.168.1.2",
// 	database:   "ccds",
// 	// windows: true 为windows身份验证，false 必须设置sa账号和密码
// 	windows: false,
// 	sa: SA{
// 		user:   "sa",
// 		passwd: "A64988329b",
// 		port:   1433,
// 	},
// }

// var db = Mssql{
// 	// 如果数据库是默认实例（MSSQLSERVER）则直接使用IP，命名实例需要指明。
// 	dataSource: "127.0.0.1",
// 	database:   "ccds",
// 	// windows: true 为windows身份验证，false 必须设置sa账号和密码
// 	windows: false,
// 	sa: SA{
// 		user:   "sa",
// 		passwd: "sa123abc()",
// 		port:   1433,
// 	},
// }

// 石家庄
var db = Mssql{
	// 如果数据库是默认实例（MSSQLSERVER）则直接使用IP，命名实例需要指明。
	dataSource: "10.27.44.10",
	database:   "ccds",
	// windows: true 为windows身份验证，false 必须设置sa账号和密码
	windows: false,
	sa: SA{
		user:   "sa",
		passwd: "sa123abc()",
		port:   1433,
	},
}

// var db = Mssql{
// 	// 如果数据库是默认实例（MSSQLSERVER）则直接使用IP，命名实例需要指明。
// 	dataSource: "10.27.152.99",
// 	database:   "ccds",
// 	// windows: true 为windows身份验证，false 必须设置sa账号和密码
// 	windows: false,
// 	sa: SA{
// 		user:   "sa",
// 		passwd: "sa123abc()",
// 		port:   1433,
// 	},
// }

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
	// 连接数据库
	err := db.Open()
	if err != nil {
		fmt.Println("sql open:", err)
		return 0, err
	}

	fmt.Println("sql open IP :", db.dataSource)

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
	// if cellRows[0][0] != "案件ID" {
	// 	fmt.Println("未找到 案件ID 列")
	// 	return
	// }
	// if cellRows[0][1] != "员工ID" {
	// 	fmt.Println("未找到 员工ID 列")
	// 	return
	// }
	for k, v := range cellRows {
		if k != 0 {
			sql_select_seno := "SELECT se_no FROM sal_emp WHERE se_no = '#{cas_se_no}'"
			sql_select_seno = strings.Replace(sql_select_seno, "#{cas_se_no}", v[1], -1)
			sql_select_seno_re_rows, sql_select_seno_re_err := DoQuery(sql_select_seno)
			if sql_select_seno_re_err != nil {
				fmt.Println("query: ", sql_select_seno_re_err)
			} else {
				if len(sql_select_seno_re_rows) == 0 {
					reStr := k + 1
					fmt.Println("无法找到第 " + fmt.Sprint(reStr) + " 行的员工ID: " + v[1])
				} else {
					sql_update_casById := "UPDATE bank_case set cas_se_no = '#{cas_se_no}' where cas_id = '#{cas_id}'"
					sql_update_casById = strings.Replace(sql_update_casById, "#{cas_id}", v[0], -1)
					sql_update_casById = strings.Replace(sql_update_casById, "#{cas_se_no}", v[1], -1)
					fmt.Println(sql_update_casById)
					rows, err := DoExec(sql_update_casById)

					if err != nil {
						fmt.Println("query: ", err)
						fmt.Println(rows)
						return
					} else {
						fmt.Println("已完成：", k)
					}
				}
			}
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
			if len(rows) == 0 {
				fmt.Println("没有找到对应的案件批次号： " + fmt.Sprint(v[0]))
				return
			}

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
					sql_select_cas_max_num := "SELECT cas_se_no FROM bank_case WHERE cas_m = (SELECT MAX(cas_m) FROM bank_case WHERE cas_cbat_id = '#{cas_cbat_id}' AND cas_num = '#{cas_num}')"
					sql_select_cas_max_num = strings.Replace(sql_select_cas_max_num, "#{cas_cbat_id}", cbat_id, -1)
					sql_select_cas_max_num = strings.Replace(sql_select_cas_max_num, "#{cas_cbat_id}", fmt.Sprint(caseRow["cas_m"]), -1)

					sql_select_cas_max_num_re_rows, sql_select_cas_max_num_err := DoQuery(sql)

					if len(sql_select_cas_max_num_re_rows) != 0 {
						re_se_no := sql_select_cas_max_num_re_rows[0]
						sql_update_seNo := "UPDATE bank_case SET cas_se_no = '#{cas_se_no}' WHERE cas_num = '#{cas_num}'"
						sql_update_seNo = strings.Replace(sql_update_seNo, "#{cas_se_no}", fmt.Sprint(re_se_no), -1)
						sql_update_seNo = strings.Replace(sql_update_seNo, "#{cas_num}", fmt.Sprint(caseRow["cas_num"]), -1)

						_, caseerr := DoExec(sql_update_seNo)

						if caseerr != nil {
							fmt.Println("query: ", caseerr)
						} else {
							fmt.Println("已完成", fmt.Sprint(j)+"/"+fmt.Sprint(len(caseRows)))
						}
					}

					if sql_select_cas_max_num_err != nil {
						fmt.Println("query: ", sql_select_cas_max_num_err)
						fmt.Println(sql_select_cas_max_num_re_rows)
					}
				}

				if caseerr != nil {
					fmt.Println("query: ", caseerr)
					fmt.Println(rows)
				}

			}
			if err != nil {
				fmt.Println("query: ", err)
				fmt.Println(rows)
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
				}
				fmt.Println(updateRows)
				completeCaseList = completeCaseList + ",'" + fmt.Sprint(data["cas_id"]) + "'"
				fmt.Println("completeCaseList:", completeCaseList)
			}
			if changeCsaIderr != nil {
				fmt.Println("query: ", changeCsaIderr)
				fmt.Println(changeCsaIdRows)
			}
		}
	}
}

func doExchangeCaseTemp(cellRows [][]string) {
	fmt.Println(cellRows)

	if cellRows[0][0] != "批次" {
		fmt.Println("未找到 批次 列")
		return
	}
	if cellRows[0][1] != "留案Code" {
		fmt.Println("未找到 留案Code 列")
		return
	}

	stayCaseCodeStringForAll := ""
	for k, v := range cellRows {
		if k != 0 {
			if k == 1 {
				stayCaseCodeStringForAll = "'" + v[1] + "'"
			} else {
				stayCaseCodeStringForAll = stayCaseCodeStringForAll + ",'" + v[1] + "'"
			}
		}
	}
	fmt.Println("stayCaseCodeStringForAll:", stayCaseCodeStringForAll)

	cbat_code_sql := ""
	for k, v := range cellRows {
		if k != 0 {
			if k == 1 {
				cbat_code_sql = "'" + v[0] + "'"
			} else {
				cbat_code_sql = cbat_code_sql + ",'" + v[0] + "'"
			}
		}
	}

	fmt.Println("cbat_code_sql:", cbat_code_sql)

	if cbat_code_sql == "" {
		return
	}
	sql := "SELECT bc.cas_se_no,COUNT(bc.cas_id) AS cas_change_num FROM dbo.bank_case AS bc INNER JOIN dbo.case_bat AS cb ON  bc.cas_cbat_id = cb.cbat_id WHERE cb.cbat_code IN (#{cbat_code_sql})"
	sql = strings.Replace(sql, "#{cbat_code_sql}", cbat_code_sql, -1)
	if stayCaseCodeStringForAll != "" {
		sql = sql + " AND bc.cas_code NOT IN (" + stayCaseCodeStringForAll + ") GROUP BY bc.cas_se_no"
	} else {
		sql = sql + " GROUP BY bc.cas_se_no"
	}

	fmt.Println("cas_se_no_list sql: ", sql)

	cas_se_no_list, err := DoQuery(sql)
	if err != nil {
		fmt.Println("query: ", err)
		fmt.Println(cas_se_no_list)
	}

	casSeNoForAll := ""
	for i, cas_se_no := range cas_se_no_list {
		if i == 0 {
			casSeNoForAll = "'" + fmt.Sprint(cas_se_no["cas_se_no"]) + "'"
		} else {
			casSeNoForAll = casSeNoForAll + ",'" + fmt.Sprint(cas_se_no["cas_se_no"]) + "'"
		}
	}

	fmt.Println("casSeNoForAll:", casSeNoForAll)

	sql = "SELECT bc.cas_id,bc.cas_se_no AS cas_change_num FROM dbo.bank_case AS bc INNER JOIN dbo.case_bat AS cb ON  bc.cas_cbat_id = cb.cbat_id WHERE cb.cbat_code IN (#{cbat_code_sql})"
	sql = strings.Replace(sql, "#{cbat_code_sql}", cbat_code_sql, -1)
	if stayCaseCodeStringForAll != "" {
		sql = sql + " AND bc.cas_code NOT IN (" + stayCaseCodeStringForAll + ")"
	}

	fmt.Println("sql:", sql)

	cas_id_se_no_list, err := DoQuery(sql)
	if err != nil {
		fmt.Println("query: ", err)
		fmt.Println(cas_id_se_no_list)
	}

	fmt.Println("cas_id_se_no_list:", cas_id_se_no_list)

	for _, cas_id_se_no := range cas_id_se_no_list {
		for _, cas_se_no := range cas_se_no_list {
			if cas_id_se_no["cas_se_no"] != cas_se_no["cas_se_no"] && cas_se_no["cas_change_num"].(int) > 0 {
				sql = "UPDATE bank_case SET cas_se_no = '#{cas_se_no}' WHERE cas_id = '#{cas_id}'"
				sql = strings.Replace(sql, "#{cas_se_no}", fmt.Sprint(cas_se_no["cas_se_no"]), -1)
				sql = strings.Replace(sql, "#{cas_id}", fmt.Sprint(cas_id_se_no["cas_id"]), -1)
				cas_se_no["cas_change_num"] = cas_se_no["cas_change_num"].(int) - 1
				break
			}
		}
		fmt.Println("sql:", sql)

		updateRows, caseerr := DoExec(sql)

		if caseerr != nil {
			fmt.Println("query: ", caseerr)
		}
		fmt.Println(updateRows)
	}
}

func clearRemark6() {
	sql := "UPDATE bank_case SET cas_remark6 = REPLACE(cas_remark6, '再委托|', '') WHERE cas_remark6 LIKE '%再委托|%'"
	DoExec(sql)
	fmt.Println("执行结束")
}

func insertPhoneInfo(rows [][]string) {
	for k, v := range rows {
		if k != 0 && len(v) == 4 {
			sql := "INSERT INTO dbo.phone_list (phl_state, phl_name, phl_num, phl_cas_id, phl_cat, phl_count, phl_remark, phl_isdel, phl_isnew, phl_upd_time ) SELECT 1,'#{phl_name}','#{phl_num}',cas_id,'第三方',0,'#{phl_remark}',NULL,1,NULL FROM bank_case WHERE cas_code = '#{cas_code}' and cas_state='0'"

			sql = strings.Replace(sql, "#{phl_name}", v[0], -1)
			sql = strings.Replace(sql, "#{phl_num}", v[1], -1)
			sql = strings.Replace(sql, "#{phl_remark}", v[2], -1)
			sql = strings.Replace(sql, "#{cas_code}", v[3], -1)
			fmt.Println(sql)
			DoExec(sql)
		}
	}

	sql := "DELETE FROM dbo.phone_list WHERE phl_id NOT IN (SELECT MAX ( phl_id ) FROM dbo.phone_list GROUP BY phl_name, phl_num, phl_cas_id, phl_cat, phl_count, phl_remark)"
	fmt.Println(sql)
	DoExec(sql)
}

func withdrawalCase(rows [][]string) {
	for k, v := range rows {
		if k != 0 && len(v) == 1 {
			sql := "UPDATE bank_case SET cas_state = '3' WHERE cas_code = '#{cas_code}' AND cas_state != '3'"
			casCode := strings.Replace(strings.Replace(v[0], " ", "", -1), "'", "", -1)
			sql = strings.Replace(sql, "#{cas_code}", casCode, -1)
			fmt.Println(sql)
			DoExec(sql)
		}
	}
}

func deleteBankCaseOnID(rows [][]string) {
	for k, v := range rows {
		if k != 0 && len(v) == 1 {
			sql := "DELETE FROM bank_case WHERE cas_id = '#{cas_id}'"
			casID := strings.Replace(strings.Replace(v[0], " ", "", -1), "'", "", -1)
			sql = strings.Replace(sql, "#{cas_id}", casID, -1)
			fmt.Println(sql)
			DoExec(sql)
		}
	}
}

func supplementCustomersHomePhoneNumber(rows [][]string) {
	for k, v := range rows {
		if k != 0 && len(v) == 2 {
			sql := "UPDATE bank_case SET cas_hom_pho = '#{cas_hom_pho}' WHERE cas_code = '#{cas_code}' AND cas_state != '3'"
			casCode := strings.Replace(strings.Replace(v[0], " ", "", -1), "'", "", -1)
			casHomPho := strings.Replace(strings.Replace(v[1], " ", "", -1), "'", "", -1)
			sql = strings.Replace(sql, "#{cas_code}", casCode, -1)
			sql = strings.Replace(sql, "#{cas_hom_pho}", casHomPho, -1)
			fmt.Println(sql)
			DoExec(sql)
		}
	}
}

func retrieveTableContentByNameUploadItToDatabase(fpath string) {
	f, err := excelize.OpenFile(fpath)
	if err != nil {
		fmt.Println(err)
	}

	// 创建一个新的Excel文件
	reF := excelize.NewFile()
	// rows, err := f.GetRows("Sheet1")
	for sheetNums, sheetName := range f.GetSheetList() {
		fmt.Println(sheetName)
		rows, err := f.GetRows(sheetName)
		if err != nil {
			fmt.Println(err)
			continue
		}

		customerName := ""
		overdueStage := ""
		regAdd := ""
		for reRowsNum, row := range rows {
			for colCellNum, colCell := range row {
				// 判断字符串是否包含 通讯地址
				if strings.Contains(colCell, "通讯地址") {
					fmt.Println("找到通信地址列")
					// 拆分字符串 删除字符串 通讯地址 前的内容
					colCell = strings.Replace(colCell,strings.Split(colCell, "通讯地址")[0], "", -1)
					regAdd = colCell
				}

				// 判断字符串是否等于 姓名
				if colCell == "姓名" {
					fmt.Println("找到姓名列")
					// 找到姓名列，获取姓名列的下一行
					nextRow := rows[reRowsNum+1]
					// 获取姓名列的下一行的第一个单元格
					firstCell := nextRow[0]
					// 获取姓名列的下一行的第二个单元格
					customerName = firstCell
				}
				// 判断字符串是否等于 逾期阶段
				if colCell == "逾期阶段" {
					fmt.Println("找到逾期阶段列")
					// 找到逾期阶段列，获取逾期阶段列的下一行
					nextRow := rows[reRowsNum+1]
					// 获取逾期阶段列的下一行的对应的单元格
					nextCell := nextRow[colCellNum]

					overdueStage = nextCell
				}

			}
		}
		//导出customerName overdueStage regAdd到 excel文件至当前目录
		reF.SetCellValue("Sheet1", "A"+strconv.Itoa(sheetNums+1), customerName)
		reF.SetCellValue("Sheet1", "B"+strconv.Itoa(sheetNums+1), overdueStage)
		reF.SetCellValue("Sheet1", "C"+strconv.Itoa(sheetNums+1), regAdd)

		fmt.Print(customerName)
		fmt.Print(overdueStage)
		fmt.Print(regAdd)
	}
	// 保存文件
	if err := reF.SaveAs("result.xlsx"); err != nil {
		fmt.Println(err)
	}

	// for k, v := range cellRows {
	// 	if k != 0 {
	// 		sql_select_seno := "SELECT se_no FROM sal_emp WHERE se_no = '#{cas_se_no}'"
	// 		sql_select_seno = strings.Replace(sql_select_seno, "#{cas_se_no}", v[1], -1)
	// 		sql_select_seno_re_rows, sql_select_seno_re_err := DoQuery(sql_select_seno)
	// 		if sql_select_seno_re_err != nil {
	// 			fmt.Println("query: ", sql_select_seno_re_err)
	// 		} else {
	// 			if len(sql_select_seno_re_rows) == 0 {
	// 				reStr := k + 1
	// 				fmt.Println("无法找到第 " + fmt.Sprint(reStr) + " 行的员工ID: " + v[1])
	// 			} else {
	// 				sql_update_casById := "UPDATE bank_case set cas_se_no = '#{cas_se_no}' where cas_id = '#{cas_id}'"
	// 				sql_update_casById = strings.Replace(sql_update_casById, "#{cas_id}", v[0], -1)
	// 				sql_update_casById = strings.Replace(sql_update_casById, "#{cas_se_no}", v[1], -1)
	// 				fmt.Println(sql_update_casById)
	// 				rows, err := DoExec(sql_update_casById)

	// 				if err != nil {
	// 					fmt.Println("query: ", err)
	// 					fmt.Println(rows)
	// 					return
	// 				} else {
	// 					fmt.Println("已完成：", k)
	// 				}
	// 			}
	// 		}
	// 	}
	// }
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
		} else if operation == "clearRemark6" {
			clearRemark6()
		} else if operation == "insertPhoneInfo" {
			rows := readXlsx(fPath)
			fmt.Println("Result rows:", rows)
			insertPhoneInfo(rows)
		} else if operation == "withdrawalCase" {
			rows := readXlsx(fPath)
			fmt.Println("Result rows:", rows)
			withdrawalCase(rows)
		} else if operation == "deleteBankCaseOnID" {
			rows := readXlsx(fPath)
			fmt.Println("Result rows:", rows)
			deleteBankCaseOnID(rows)
		} else if operation == "supplementCustomersHomePhoneNumber" {
			rows := readXlsx(fPath)
			fmt.Println("Result rows:", rows)
			supplementCustomersHomePhoneNumber(rows)
		} else if operation == "retrieveTableContentByNameUploadItToDatabase" {
			retrieveTableContentByNameUploadItToDatabase(fPath)
		} else {
			help()
		}
		return nil
	}

	app.Run(os.Args)
}

// func main() {
// 	// rows := readXlsx("E:\\work\\goproject\\gomssql\\moban\\huanshou1.xlsx")
// 	// fmt.Println("Result rows:", rows)
// 	// doExchangeCaseTemp(rows)
// 	retrieveTableContentByNameUploadItToDatabase("E:\\work\\goproject\\src\\code.fylan.com\\servergroup\\gomssql\\294.xlsx")
// }
