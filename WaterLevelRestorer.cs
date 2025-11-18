using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Diagnostics;
using System.Text;
using OfficeOpenXml;

namespace WaterLevelCompleter
{
    /// <summary>
    /// 水位数据恢复器，用于将省略了个位和十位的水位数据恢复完整
    /// </summary>
    public class WaterLevelRestorer
    {
        // 用于存储每个月第一天的水位值，作为该月的基准值
        private Dictionary<int, double> _monthlyBaseLevels = new Dictionary<int, double>();
        // 用于存储每个单元格的前一天完整值
        private Dictionary<string, double> _previousDayValues = new Dictionary<string, double>();
        // 处理状态跟踪
        public int ProcessedCount { get; private set; } = 0;
        public int SkippedCount { get; private set; } = 0;
        // 错误信息收集
        public List<string> ProcessingErrors { get; private set; } = new List<string>();
        
        // 配置参数
        private const double MIN_WATER_LEVEL = 0.0;      // 最小水位值
        private const double MAX_WATER_LEVEL = 100.0;    // 最大水位值
        private const int MAX_TEXT_LENGTH = 5;           // 最大文本长度
        
        /// <summary>
        /// 恢复Excel单元格区域中的水位数据
        /// </summary>
        /// <param name="range">需要处理的Excel单元格区域</param>
        /// <exception cref="ArgumentNullException">当range为空时抛出</exception>
        /// <exception cref="ArgumentException">当range无效时抛出</exception>
        public void RestoreWaterLevels(ExcelRange range)
        {
            // 参数验证
            if (range == null)
                throw new ArgumentNullException(nameof(range), "Excel单元格区域不能为空");
            
            // 检查区域是否有效
            if (range.Start.Row > range.End.Row || range.Start.Column > range.End.Column)
                throw new ArgumentException("无效的单元格区域范围", nameof(range));
            
            if (range.Worksheet == null)
                throw new ArgumentException("无效的工作表", nameof(range));
            // 重置状态
            ProcessedCount = 0;
            SkippedCount = 0;
            ProcessingErrors.Clear();
            
            // 清空之前的数据
            _monthlyBaseLevels.Clear();
            _previousDayValues.Clear();
            
            try
            {
                // 获取区域的维度
                int startRow = range.Start.Row;
                int endRow = range.End.Row;
                int startColumn = range.Start.Column;
                int endColumn = range.End.Column;
                
                // 验证区域大小
                if ((endRow - startRow + 1) > 1000 || (endColumn - startColumn + 1) > 100)
                {
                    throw new ArgumentException("处理区域过大，可能导致性能问题");
                }
                
                // 遍历每个单元格，处理水位数据
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        try
                        {
                            // 不再跳过区域内的任何单元格，用户指定的区域即为完整的数据区域
                            
                            // 验证单元格坐标
                            if (row < 1 || col < 1 || 
                                row > range.Worksheet.Dimension.Rows || 
                                col > range.Worksheet.Dimension.Columns)
                            {
                                SkippedCount++;
                                ProcessingErrors.Add($"单元格[{row},{col}]超出工作表范围");
                                continue;
                            }
                            
                            // 处理当前单元格
                            ExcelRange cell = range.Worksheet.Cells[row, col];
                            if (cell == null)
                            {
                                SkippedCount++;
                                continue;
                            }
                            
                            int month = GetMonthFromColumn(col, range);
                            int day = GetDayFromRow(row, range);
                            
                            // 验证月份和日期的有效性
                            if (month < 1 || month > 12 || day < 1 || day > GetDaysInMonth(month, 2020))
                            {
                                SkippedCount++;
                                ProcessingErrors.Add($"单元格[{row},{col}]的日期信息无效: {month}月{day}日");
                                continue;
                            }
                            
                            // 生成月份的唯一键，用于跟踪前一天的值
                            string monthKey = $"{month}_{col}";
                            
                            // 恢复单元格值
                            if (RestoreCellValue(cell, month, day, monthKey))
                            {
                                ProcessedCount++;
                            }
                            else
                            {
                                SkippedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            SkippedCount++;
                            string errorMsg = $"处理单元格[{row},{col}]时发生错误: {ex.Message}";
                            ProcessingErrors.Add(errorMsg);
                            Debug.WriteLine(errorMsg);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ProcessingErrors.Add($"处理数据区域时发生严重错误: {ex.Message}");
                Debug.WriteLine($"处理失败: {ex}");
                throw;
            }
            
            Debug.WriteLine($"成功处理了 {ProcessedCount} 个单元格的数据，跳过了 {SkippedCount} 个单元格");
        }
        
        /// <summary>
        /// 恢复单个单元格的水位值
        /// </summary>
        /// <param name="cell">Excel单元格</param>
        /// <param name="month">月份</param>
        /// <param name="day">日期</param>
        /// <param name="monthKey">月份的唯一键，用于跟踪前一天的值</param>
        /// <returns>是否成功恢复了数据</returns>
        private bool RestoreCellValue(ExcelRange cell, int month, int day, string monthKey)
        {
            try
            {
                // 安全获取单元格文本
                string cellValue = cell.Text?.Trim() ?? string.Empty;
                
                // 如果单元格为空或文本过长，跳过
                if (string.IsNullOrEmpty(cellValue) || cellValue.Length > MAX_TEXT_LENGTH)
                {                    
                    return false;
                }
                
                // 检查是否是完整的数值格式（包含小数点）
                if (cellValue.Contains("."))
                {
                    // 尝试解析为double
                    if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double fullValue) &&
                        IsValidWaterLevel(fullValue))
                    {
                        // 对于每月第一天，保存为基准值
                        if (day == 1 && !_monthlyBaseLevels.ContainsKey(month))
                        {
                            _monthlyBaseLevels[month] = fullValue;
                        }
                        
                        // 保存当前值作为下一天的参考
                        _previousDayValues[monthKey] = fullValue;
                        
                        // 已经是完整值，不需要修改
                        return false;
                    }
                }
                
                // 判断是否是省略的数值（只有小数部分）
                // 例如："17" 表示小数部分 .17
                if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedValue) &&
                    parsedValue >= 0 && parsedValue <= 999) // 限制数值范围，避免异常大的数字
                {
                    // 获取前一天的完整值作为基准
                    double previousValue;
                    if (_previousDayValues.TryGetValue(monthKey, out previousValue))
                    {
                        // 提取前一天值的整数部分（个位和十位）
                        int integerPart = (int)Math.Floor(previousValue);
                        
                        // 处理小数部分
                        double decimalPart;
                        string decimalStr = cellValue;
                        
                        // 根据输入长度确定小数位数
                        if (decimalStr.Length == 1 && parsedValue < 10)
                        {
                            // 一位数，补零，如 "7" 变为 ".70"
                            decimalPart = parsedValue / 100.0;
                        }
                        else
                        {
                            // 两位数或更多，如 "17" 变为 ".17"
                            decimalPart = parsedValue / 100.0;
                        }
                        
                        // 构造完整的数值（整数部分 + 小数部分）
                        double restoredValue = integerPart + decimalPart;
                        
                        // 检查恢复后的值是否合理
                        if (IsValidWaterLevel(restoredValue))
                        {
                            // 保存恢复后的值到单元格
                            cell.Value = restoredValue;
                            
                            // 更新前一天的值，用于下一行的恢复
                            _previousDayValues[monthKey] = restoredValue;
                            
                            return true;
                        }
                    }// 对于第一天，如果没有前一天但有月基准值
                    else if (day == 1 && _monthlyBaseLevels.TryGetValue(month, out double monthlyBase))
                    {
                        Debug.WriteLine($"使用月基准值恢复{month}月{day}日的值");
                        // 对于第一天，如果没有前一天但有月基准值
                        int integerPart = (int)Math.Floor(monthlyBase);
                        double decimalPart = parsedValue / 100.0;
                        double restoredValue = integerPart + decimalPart;
                        
                        if (IsValidWaterLevel(restoredValue))
                        {
                            cell.Value = restoredValue;
                            _previousDayValues[monthKey] = restoredValue;
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"处理单元格 [{cell.Address}] 时发生错误: {ex.Message}";
                ProcessingErrors.Add(errorMsg);
                Debug.WriteLine(errorMsg);
            }
            
            return false;
        }
        
        /// <summary>
        /// 从列索引获取月份
        /// </summary>
        /// <param name="column">列索引</param>
        /// <param name="range">Excel单元格区域</param>
        /// <returns>月份（1-12）</returns>
        private int GetMonthFromColumn(int column, ExcelRange range)
        {    
            try
            {
                // 对于用户指定的区域（如C6:N36），直接按列计算月份
                // 第1列对应1月，第2列对应2月，...，第12列对应12月
                int month = column - range.Start.Column + 1;
                
                // 确保月份在1-12之间
                return Math.Min(12, Math.Max(1, month));
            }
            catch (Exception ex)
            {
                ProcessingErrors.Add($"获取月份时出错: {ex.Message}");
                Debug.WriteLine($"获取月份时出错: {ex.Message}");
                // 出错时返回默认值
                return Math.Max(1, Math.Min(12, column - range.Start.Column + 1));
            }
        }
        
        /// <summary>
        /// 从行索引获取日期
        /// </summary>
        /// <param name="row">行索引</param>
        /// <param name="range">Excel单元格区域</param>
        /// <returns>日期（1-31）</returns>
        private int GetDayFromRow(int row, ExcelRange range)
        {    
            try
            {
                // 对于用户指定的区域（如C6:N36），直接按行计算日期
                // 第1行对应1日，第2行对应2日，...，第31行对应31日
                int day = row - range.Start.Row + 1;
                
                // 确保日期至少为1
                return Math.Max(1, day);
            }
            catch (Exception ex)
            {
                ProcessingErrors.Add($"获取日期时出错: {ex.Message}");
                // 出错时返回默认值
                return Math.Max(1, Math.Min(31, row - range.Start.Row + 1));
            }
        }
        
        /// <summary>
        /// 验证水位值是否合理
        /// </summary>
        private bool IsValidWaterLevel(double value)
        {
            // 根据水文数据的实际情况设置合理的水位范围
            return value >= MIN_WATER_LEVEL && value <= MAX_WATER_LEVEL;
        }
        
        /// <summary>
        /// 获取指定月份的天数
        /// </summary>
        /// <param name="month">月份</param>
        /// <param name="year">年份</param>
        /// <returns>天数</returns>
        private int GetDaysInMonth(int month, int year)
        {            
            try
            {
                return DateTime.DaysInMonth(year, month);
            }
            catch
            {
                // 出错时返回默认值
                return month switch
                {
                    2 => 29, // 假设是闰年2月
                    4 or 6 or 9 or 11 => 30,
                    _ => 31
                };
            }
        }
        
        /// <summary>
        /// 获取处理过程中的错误信息摘要
        /// </summary>
        /// <returns>错误信息摘要字符串</returns>
        public string GetErrorSummary()
        {            
            if (ProcessingErrors.Count == 0)
                return "无错误";
            
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"共发现 {ProcessingErrors.Count} 个错误:");
            
            // 只显示前10个错误
            int maxErrorsToShow = Math.Min(10, ProcessingErrors.Count);
            for (int i = 0; i < maxErrorsToShow; i++)
            {
                sb.AppendLine($"- {ProcessingErrors[i]}");
            }
            
            if (ProcessingErrors.Count > maxErrorsToShow)
            {
                sb.AppendLine($"... 还有 {ProcessingErrors.Count - maxErrorsToShow} 个错误未显示");
            }
            
            return sb.ToString();
        }
    }
}