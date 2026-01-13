/*
Copyright © 2026 NAME lichenliang <1031809056@qq.com>
*/
package cmd

import (
	"fmt"
	"math"
	"os"
	"sort"
	"strconv"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var Name string

type Student struct {
	ID              string
	Name            string
	Gender          string
	ObjectiveScore  float64
	SubjectiveScore float64
	TotalDuifenyi   float64
	DailyScore      float64
	FinalScore      float64
}

var rootCmd = &cobra.Command{
	Use:   "handle",
	Short: "\nhandle scores in excel to obtain a normal distribution.",
	Run: func(cmd *cobra.Command, args []string) {
		if Name == "" {
			fmt.Println("请使用 -n 参数指定文件名")
			return
		}

		students, err := readExcel(Name)
		if err != nil {
			fmt.Printf("读取文件失败: %v\n", err)
			return
		}

		adjustScores(students)

		// 生成新文件名
		outputName := generateOutputFilename(Name)
		err = writeExcel(outputName, students)
		if err != nil {
			fmt.Printf("写入文件失败: %v\n", err)
			return
		}

		fmt.Printf("成绩调整完成！已保存到: %s\n", outputName)
		printStatistics(students)
	},
}

func Execute() {
	err := rootCmd.Execute()
	if err != nil {
		os.Exit(1)
	}
}

func init() {
	rootCmd.Flags().StringVarP(&Name, "name", "n", "", "handle file name")
}

func readExcel(filename string) ([]*Student, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		return nil, err
	}

	var students []*Student
	for i, row := range rows {
		if i == 0 || len(row) < 7 {
			continue
		}

		objScore, _ := strconv.ParseFloat(row[3], 64)
		subjScore, _ := strconv.ParseFloat(row[4], 64)
		duifenyiTotal, _ := strconv.ParseFloat(row[5], 64)
		dailyScore, _ := strconv.ParseFloat(row[6], 64)

		student := &Student{
			ID:              row[0],
			Name:            row[1],
			Gender:          row[2],
			ObjectiveScore:  objScore,
			SubjectiveScore: subjScore,
			TotalDuifenyi:   duifenyiTotal,
			DailyScore:      dailyScore,
		}
		student.FinalScore = duifenyiTotal*0.4 + dailyScore*0.6
		students = append(students, student)
	}

	return students, nil
}

func adjustScores(students []*Student) {
	// 第一步：确保所有人总分 >= 60
	for _, s := range students {
		if s.FinalScore < 60 {
			requiredDaily := (60 - s.TotalDuifenyi*0.4) / 0.6
			if requiredDaily > s.DailyScore {
				s.DailyScore = requiredDaily
				s.FinalScore = s.TotalDuifenyi*0.4 + s.DailyScore*0.6
			}
		}
	}

	// 第二步：调整为正态分布
	targetMean := 75.0
	targetStdDev := 8.0

	// 对学生按当前总分排序
	sort.Slice(students, func(i, j int) bool {
		return students[i].FinalScore < students[j].FinalScore
	})

	// 使用正态分布调整分数
	for i, s := range students {
		// 计算该学生在正态分布中的位置（百分位）
		percentile := float64(i+1) / float64(len(students)+1)

		// 使用逆正态分布计算目标分数
		zScore := inverseNormalCDF(percentile)
		targetScore := targetMean + zScore*targetStdDev

		// 确保目标分数在合理范围内
		if targetScore < 60 {
			targetScore = 60
		}
		if targetScore > 100 {
			targetScore = 100
		}

		// 如果目标分数高于当前分数，调整平时分
		if targetScore > s.FinalScore {
			requiredDaily := (targetScore - s.TotalDuifenyi*0.4) / 0.6
			if requiredDaily > s.DailyScore && requiredDaily <= 100 {
				s.DailyScore = requiredDaily
				s.FinalScore = s.TotalDuifenyi*0.4 + s.DailyScore*0.6
			}
		}
	}
}

func calculateMean(students []*Student) float64 {
	sum := 0.0
	for _, s := range students {
		sum += s.FinalScore
	}
	return sum / float64(len(students))
}

func calculateStdDev(students []*Student, mean float64) float64 {
	sumSquares := 0.0
	for _, s := range students {
		diff := s.FinalScore - mean
		sumSquares += diff * diff
	}
	return math.Sqrt(sumSquares / float64(len(students)))
}

func inverseNormalCDF(p float64) float64 {
	// 使用近似算法计算标准正态分布的逆累积分布函数
	if p <= 0 {
		return -10
	}
	if p >= 1 {
		return 10
	}

	// Beasley-Springer-Moro 算法的简化版本
	a := []float64{-3.969683028665376e+01, 2.209460984245205e+02,
		-2.759285104469687e+02, 1.383577518672690e+02,
		-3.066479806614716e+01, 2.506628277459239e+00}
	b := []float64{-5.447609879822406e+01, 1.615858368580409e+02,
		-1.556989798598866e+02, 6.680131188771972e+01,
		-1.328068155288572e+01}
	c := []float64{-7.784894002430293e-03, -3.223964580411365e-01,
		-2.400758277161838e+00, -2.549732539343734e+00,
		4.374664141464968e+00, 2.938163982698783e+00}
	d := []float64{7.784695709041462e-03, 3.224671290700398e-01,
		2.445134137142996e+00, 3.754408661907416e+00}

	pLow := 0.02425
	pHigh := 1 - pLow

	var q, r, result float64

	if p < pLow {
		q = math.Sqrt(-2 * math.Log(p))
		result = (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q + c[5]) /
			((((d[0]*q+d[1])*q+d[2])*q+d[3])*q + 1)
	} else if p <= pHigh {
		q = p - 0.5
		r = q * q
		result = (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r + a[5]) * q /
			(((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r + 1)
	} else {
		q = math.Sqrt(-2 * math.Log(1-p))
		result = -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q + c[5]) /
			((((d[0]*q+d[1])*q+d[2])*q+d[3])*q + 1)
	}

	return result
}

func generateOutputFilename(inputName string) string {
	ext := ".xls"
	if strings.HasSuffix(inputName, ".xlsx") {
		ext = ".xlsx"
	}
	baseName := strings.TrimSuffix(inputName, ext)
	return baseName + "_adjusted" + ext
}

func writeExcel(filename string, students []*Student) error {
	f := excelize.NewFile()
	defer f.Close()

	// 创建表头
	headers := []string{"学号", "学生", "性别", "客观题得分", "主观题得分", "对分易总分", "平时分", "总分"}
	for i, header := range headers {
		col := string(rune('A' + i))
		f.SetCellValue("Sheet1", col+"1", header)
	}

	// 写入学生数据
	for i, s := range students {
		row := i + 2
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), s.ID)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), s.Name)
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), s.Gender)
		f.SetCellValue("Sheet1", fmt.Sprintf("D%d", row), s.ObjectiveScore)
		f.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), s.SubjectiveScore)
		f.SetCellValue("Sheet1", fmt.Sprintf("F%d", row), s.TotalDuifenyi)
		f.SetCellValue("Sheet1", fmt.Sprintf("G%d", row), fmt.Sprintf("%.0f", s.DailyScore))
		f.SetCellValue("Sheet1", fmt.Sprintf("H%d", row), fmt.Sprintf("%.1f", s.FinalScore))
	}

	return f.SaveAs(filename)
}

func printStatistics(students []*Student) {
	mean := calculateMean(students)
	stdDev := calculateStdDev(students, mean)

	fmt.Printf("\n统计信息:\n")
	fmt.Printf("平均分: %.2f\n", mean)
	fmt.Printf("标准差: %.2f\n", stdDev)
	fmt.Printf("最低分: %.2f\n", students[0].FinalScore)
	fmt.Printf("最高分: %.2f\n", students[len(students)-1].FinalScore)
}
