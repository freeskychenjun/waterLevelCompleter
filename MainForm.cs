using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Diagnostics;

namespace WaterLevelCompleter
{
    public partial class MainForm : Form
    {
        private ExcelPackage? _excelPackage = null;
        private string? _currentFilePath = null;

        public MainForm()
        {
            InitializeUI();
        }

        private void InitializeUI()
        {
            // 设置表单属性
            this.Text = "水位数据补齐工具";
            this.Size = new Size(550, 520);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            
            // 创建表格布局面板
            TableLayoutPanel layoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 9,
                Padding = new Padding(10),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
            };
            
            // 设置列样式
            layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
            layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40F));
            
            // 设置行样式
            for (int i = 0; i < layoutPanel.RowCount; i++)
            {
                layoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
            }
            layoutPanel.RowStyles[7] = new RowStyle(SizeType.Percent, 100F); // 状态框占剩余空间
            
            // 创建文件选择按钮
            Button btnSelectFile = new Button
            {
                Text = "选择Excel文件",
                Dock = DockStyle.Fill
            };
            btnSelectFile.Click += BtnSelectFile_Click;
            layoutPanel.Controls.Add(btnSelectFile, 0, 0);
            
            // 创建文件路径显示框
            txtFilePath = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Multiline = true,
                ScrollBars = ScrollBars.Horizontal
            };
            layoutPanel.SetColumnSpan(txtFilePath, 2);
            layoutPanel.Controls.Add(txtFilePath, 1, 0);
            
            // 创建工作表标签
            Label lblSheet = new Label
            {
                Text = "选择工作表:",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            layoutPanel.Controls.Add(lblSheet, 0, 1);
            
            // 创建工作表下拉列表
            cmbSheetNames = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbSheetNames.SelectedIndexChanged += CmbSheetNames_SelectedIndexChanged;
            layoutPanel.SetColumnSpan(cmbSheetNames, 2);
            layoutPanel.Controls.Add(cmbSheetNames, 1, 1);
            
            // 创建数据区域标签
            Label lblRange = new Label
            {
                Text = "数据区域范围:",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            layoutPanel.Controls.Add(lblRange, 0, 2);
            
            // 创建数据区域输入框
            txtRange = new TextBox
            {
                Dock = DockStyle.Fill,
                Text = "C6:N36" // 默认范围修改为C6:N36
            };
            layoutPanel.Controls.Add(txtRange, 1, 2);
            
            // 创建数据区域提示标签
            Label rangeTipLabel = new Label
            {
                Text = "(多个区域用逗号分隔，例如: C6:N36,D7:E30)",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.Gray,
                Font = new Font(lblRange.Font, FontStyle.Italic)
            };
            layoutPanel.SetColumnSpan(rangeTipLabel, 2);
            layoutPanel.Controls.Add(rangeTipLabel, 1, 3);
            
            // 创建帮助按钮
            Button btnHelpRange = new Button
            {
                Text = "?",
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat
            };
            btnHelpRange.Click += BtnHelpRange_Click;
            layoutPanel.Controls.Add(btnHelpRange, 2, 2);
            
            // 创建预览数据按钮
            Button btnPreviewData = new Button
            {
                Text = "预览数据",
                Dock = DockStyle.Fill
            };
            btnPreviewData.Click += BtnPreviewData_Click;
            layoutPanel.Controls.Add(btnPreviewData, 0, 3);
            
            // 创建进度条标签
            Label lblProgress = new Label
            {
                Text = "处理进度:",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            layoutPanel.Controls.Add(lblProgress, 0, 4);
            
            // 创建进度条
            progressBar1 = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };
            layoutPanel.SetColumnSpan(progressBar1, 2);
            layoutPanel.Controls.Add(progressBar1, 1, 4);
            
            // 创建处理按钮
            btnProcess = new Button
            {
                Text = "处理数据",
                Dock = DockStyle.Fill,
                Enabled = false,
                BackColor = Color.LightSeaGreen,
                ForeColor = Color.White
            };
            btnProcess.Click += BtnProcessFile_Click;
            layoutPanel.SetColumnSpan(btnProcess, 3);
            layoutPanel.Controls.Add(btnProcess, 0, 5);
            
            // 在同步处理模式下不需要取消按钮
            
            // 创建状态显示文本框
            txtStatus = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Courier New", 9F)
            };
            layoutPanel.SetColumnSpan(txtStatus, 3);
            layoutPanel.Controls.Add(txtStatus, 0, 7);
            
            // 创建状态栏
            statusStrip1 = new StatusStrip();
            toolStripStatusLabel1 = new ToolStripStatusLabel { Text = "就绪" };
            statusStrip1.Items.Add(toolStripStatusLabel1);
            
            // 设置错误日志按钮
            btnErrorLog.Text = "错误日志";
            btnErrorLog.Click += BtnErrorLog_Click;
            statusStrip1.Items.Add(btnErrorLog);
            
            // 将布局面板添加到表单
            this.Controls.Add(layoutPanel);
            this.Controls.Add(statusStrip1);
            
            // 设置初始状态
            UpdateStatus("欢迎使用水位数据恢复工具！\n\n功能说明：\n1. 选择包含水位数据的Excel文件\n2. 选择目标工作表\n3. 指定数据区域范围\n4. 点击处理按钮开始恢复省略的水位数值");
            
            // 添加工具提示
            toolTip1 = new ToolTip();
            toolTip1.SetToolTip(btnSelectFile, "选择需要处理的Excel文件");
            toolTip1.SetToolTip(cmbSheetNames, "选择包含水位数据的工作表");
            toolTip1.SetToolTip(txtRange, "输入数据区域范围，例如：A2:N32");
            toolTip1.SetToolTip(btnProcess, "开始处理水位数据");
            toolTip1.SetToolTip(btnPreviewData, "预览选中的数据区域");
            toolTip1.SetToolTip(btnHelpRange, "查看数据区域格式帮助");
        }

        private void BtnSelectFile_Click(object? sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*";
                openFileDialog.Title = "选择水位数据Excel文件";
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _currentFilePath = openFileDialog.FileName;
                    txtFilePath.Text = _currentFilePath;
                    
                    try
                    {
                        // 加载Excel文件
                        LoadExcelFile(_currentFilePath);
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"错误：无法加载Excel文件。{ex.Message}");
                        ClearSheetSelection();
                    }
                }
            }
        }

        private void LoadExcelFile(string filePath)
        {
            try
            {
                // 设置EPPlus许可模式
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                
                // 关闭之前打开的包
                if (_excelPackage != null)
                {
                    _excelPackage.Dispose();
                }
                
                // 打开新的Excel文件
                _excelPackage = new ExcelPackage(new FileInfo(filePath));
                
                // 填充工作表下拉列表
                cmbSheetNames.Items.Clear();
                foreach (ExcelWorksheet worksheet in _excelPackage.Workbook.Worksheets)
                {
                    cmbSheetNames.Items.Add(worksheet.Name);
                }
                
                // 默认选中第一个工作表
                if (cmbSheetNames.Items.Count > 0)
                {
                    cmbSheetNames.SelectedIndex = 0;
                }
                
                UpdateStatus($"已成功加载Excel文件：{Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                UpdateStatus($"加载Excel文件时发生错误：{ex.Message}");
                throw;
            }
        }

        private void CmbSheetNames_SelectedIndexChanged(object? sender, EventArgs e)
        {
            btnProcess.Enabled = cmbSheetNames.SelectedIndex >= 0 && !string.IsNullOrEmpty(_currentFilePath);
            if (btnProcess.Enabled)
            {
                UpdateStatus($"已选择工作表：{cmbSheetNames.SelectedItem}");
            }
        }

        private void BtnHelpRange_Click(object? sender, EventArgs e)
        {
            MessageBox.Show("请输入需要处理的数据区域范围，例如：A2:N32\n\n" +
                            "其中：\n" +
                            "- A2 是数据区域的左上角单元格\n" +
                            "- N32 是数据区域的右下角单元格\n\n" +
                            "可以一次输入多个数据区域\n" +
                            "例如：B4:M34,B46:M76,B88:M118,B130:M160\n" +
                            "请确保选择的区域包含完整的水位数据。",
                            "数据区域帮助",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

        private void BtnProcessFile_Click(object? sender, EventArgs e)
        {   
            // 验证输入
            if (!ValidateInput())
            {
                return;
            }
            
            // 准备处理数据
            btnProcess.Enabled = false;
            this.Cursor = Cursors.WaitCursor;
            progressBar1.Value = 0;
            UpdateStatus("开始处理数据...");
            
            try
            {   
                // 直接处理数据
                var sheetName = cmbSheetNames.SelectedItem?.ToString() ?? string.Empty;
                ProcessResult result = ProcessDataSynchronously(
                    _currentFilePath,
                    sheetName,
                    txtRange.Text
                );
                
                // 处理完成，更新UI
                progressBar1.Value = 100;
                
                if (result.Success)
                {   
                    // 显示处理结果
                    StringBuilder statusMessage = new StringBuilder();
                    statusMessage.AppendLine("数据处理完成！");
                    statusMessage.AppendLine($"成功恢复了 {result.ProcessedCount} 个单元格的数据");
                    statusMessage.AppendLine($"跳过了 {result.SkippedCount} 个单元格");
                    statusMessage.AppendLine($"处理后文件已保存至：");
                    statusMessage.AppendLine(result.OutputPath);
                    statusMessage.AppendLine();
                    statusMessage.AppendLine("错误摘要：");
                    statusMessage.AppendLine(result.ErrorSummary);
                    
                    UpdateStatus(statusMessage.ToString());
                    
                    // 提示用户
                    DialogResult userChoice = MessageBox.Show(
                        "数据处理已完成！\n\n" +
                        $"成功恢复：{result.ProcessedCount} 个单元格\n" +
                        $"跳过：{result.SkippedCount} 个单元格\n\n" +
                        "是否打开处理后的文件？", 
                        "处理完成", 
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);
                    
                    if (userChoice == DialogResult.Yes)
                    {   
                        try
                            {   
                                // 在Windows上直接启动文件，让系统使用默认程序打开
                                Process.Start(new ProcessStartInfo
                                {
                                    FileName = result.OutputPath,
                                    UseShellExecute = true
                                });
                            }
                            catch (Exception ex)
                            {   
                                MessageBox.Show($"无法打开文件：{ex.Message}", "错误", 
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                    }
                }
                else
                {   
                    UpdateStatus(result.ErrorMessage);
                    MessageBox.Show(result.ErrorMessage, "处理失败", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {   
                UpdateStatus($"处理数据时发生错误：{ex.Message}");
                MessageBox.Show($"处理失败：{ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {   
                // 恢复界面状态
                btnProcess.Enabled = true;
                this.Cursor = Cursors.Default;
            }
        }
        
        /// <summary>
        /// 同步处理数据的方法
        /// </summary>
        private ProcessResult ProcessDataSynchronously(string filePath, string sheetName, string rangeString)
        {   
            try
            {   
                // 创建新的ExcelPackage实例进行处理
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                if (string.IsNullOrEmpty(filePath))
                {   
                    return new ProcessResult { Success = false, ErrorMessage = "错误：文件路径为空" };
                }
                
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {   
                    // 获取选定的工作表
                    if (string.IsNullOrEmpty(sheetName))
                    {   
                        return new ProcessResult { Success = false, ErrorMessage = "错误：工作表名称为空" };
                    }
                    
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
                    
                    if (worksheet == null)
                    {   
                        return new ProcessResult { Success = false, ErrorMessage = $"错误：找不到工作表 '{sheetName}'" };
                    }
                    
                    // 获取用户指定的多个单元格区域
                    List<string> rangeStrings = ProcessDataArgs.GetRangeStrings(rangeString);
                    
                    if (rangeStrings.Count == 0)
                    {   
                        return new ProcessResult { Success = false, ErrorMessage = "错误：未指定有效的数据区域" };
                    }
                    
                    // 计算总单元格数
                    int totalCells = 0;
                    List<ExcelRange> ranges = new List<ExcelRange>();
                    
                    foreach (string rangeStr in rangeStrings)
                    {   
                        try
                        {   
                            ExcelRange range = worksheet.Cells[rangeStr];
                            ranges.Add(range);
                            totalCells += (range.End.Row - range.Start.Row + 1) * (range.End.Column - range.Start.Column + 1);
                        }
                        catch (Exception ex)
                        {   
                            return new ProcessResult { Success = false, ErrorMessage = $"错误：区域 '{rangeStr}' 无效。{ex.Message}" };
                        }
                    }
                    
                    // 更新进度
                    progressBar1.Value = 10;
                    UpdateStatus($"准备处理 {rangeStrings.Count} 个数据区域\n总单元格数：{totalCells}\n开始恢复水位数据...");
                    Application.DoEvents(); // 允许UI更新
                    
                    // 创建水位恢复器实例
                    WaterLevelRestorer restorer = new WaterLevelRestorer();
                    
                    // 处理每个区域的数据
                    foreach (ExcelRange range in ranges)
                    {   
                        restorer.RestoreWaterLevels(range);
                    }
                    
                    // 更新进度
                    progressBar1.Value = 80;
                    UpdateStatus("数据恢复完成，正在保存文件...");
                    Application.DoEvents(); // 允许UI更新
                    
                    // 保存修改后的文件
                    string directory = Path.GetDirectoryName(filePath) ?? Directory.GetCurrentDirectory();
                    string fileName = Path.GetFileNameWithoutExtension(filePath) ?? "输出文件";
                    string extension = Path.GetExtension(filePath) ?? ".xlsx";
                    string outputPath = Path.Combine(directory, $"{fileName}_已恢复{extension}");
                    
                    // 检查文件是否已存在
                    int counter = 1;
                    string basePath = outputPath;
                    while (File.Exists(outputPath))
                    {   
                        outputPath = Path.Combine(directory, 
                            $"{fileName}_已恢复{counter}{extension}");
                        counter++;
                    }
                    
                    // 保存文件
                    package.SaveAs(new FileInfo(outputPath));
                    
                    // 构建结果
                    return new ProcessResult
                    {   
                        Success = true,
                        ProcessedCount = restorer.ProcessedCount,
                        SkippedCount = restorer.SkippedCount,
                        OutputPath = outputPath,
                        ErrorSummary = restorer.GetErrorSummary()
                    };
                }
            }
            catch (Exception ex)
            {   
                return new ProcessResult { Success = false, ErrorMessage = $"处理数据时发生错误：{ex.Message}" };
            }
        }
        
        // 移除了异步处理相关的事件处理方法
        
        /// <summary>
        /// 预览数据按钮点击事件
        /// </summary>
        private void BtnPreviewData_Click(object? sender, EventArgs e)
        {   
            // 验证输入
            if (string.IsNullOrEmpty(_currentFilePath))
            {
                UpdateStatus("错误：请先选择Excel文件");
                return;
            }
            
            if (cmbSheetNames.SelectedIndex == -1)
            {
                UpdateStatus("错误：请选择工作表");
                return;
            }
            
            if (string.IsNullOrEmpty(txtRange.Text))
            {
                UpdateStatus("错误：请输入数据区域范围");
                return;
            }
            
            try
            {
                // 获取选定的工作表
                var selectedItem = cmbSheetNames.SelectedItem;
                if (selectedItem == null)
                {
                    UpdateStatus("错误：请选择一个工作表");
                    return;
                }
                
                string sheetName = selectedItem.ToString();
                
                if (_excelPackage == null)
                {
                    UpdateStatus("错误：Excel包未初始化");
                    return;
                }
                
                ExcelWorksheet worksheet = _excelPackage.Workbook.Worksheets[sheetName];
                
                if (worksheet == null)
                {
                    UpdateStatus($"错误：找不到工作表 '{sheetName}'");
                    return;
                }
                
                // 获取用户指定的多个单元格区域
                List<string> rangeStrings = ProcessDataArgs.GetRangeStrings(txtRange.Text);
                
                if (rangeStrings.Count == 0)
                {
                    UpdateStatus("错误：未指定有效的数据区域");
                    return;
                }
                
                // 创建简单的预览对话框
                using (Form previewForm = new Form())
                {
                    previewForm.Text = "数据预览";
                    previewForm.Size = new Size(800, 600);
                    previewForm.StartPosition = FormStartPosition.CenterParent;
                    
                    // 创建数据网格视图
                    DataGridView dataGridView = new DataGridView
                    {
                        Dock = DockStyle.Fill,
                        AutoGenerateColumns = true,
                        ReadOnly = true,
                        AllowUserToAddRows = false,
                        AllowUserToDeleteRows = false,
                        RowHeadersVisible = true
                    };
                    
                    // 创建数据表
                    DataTable dataTable = new DataTable();
                    
                    // 添加区域标识列
                    dataTable.Columns.Add("区域", typeof(string));
                    
                    // 处理每个数据区域
                    int totalRows = 0;
                    foreach (string rangeStr in rangeStrings)
                    {
                        try
                        {
                            ExcelRange range = worksheet.Cells[rangeStr];
                            
                            // 确保数据表有足够的列
                            int colCount = range.End.Column - range.Start.Column + 1;
                            while (dataTable.Columns.Count < colCount + 1) // +1 因为有区域标识列
                            {
                                dataTable.Columns.Add(GetExcelColumnName(range.Start.Column + dataTable.Columns.Count - 1));
                            }
                            
                            // 添加数据行
                            for (int row = range.Start.Row; row <= range.End.Row; row++)
                            {
                                DataRow dataRow = dataTable.NewRow();
                                dataRow["区域"] = rangeStr; // 设置区域标识
                                
                                for (int col = range.Start.Column; col <= range.End.Column; col++)
                                {
                                    int dataColIndex = col - range.Start.Column + 1; // +1 因为第一列是区域标识
                                    dataRow[dataColIndex] = worksheet.Cells[row, col].Text ?? "";
                                }
                                dataTable.Rows.Add(dataRow);
                                totalRows++;
                            }
                        }
                        catch (Exception ex)
                        {
                            UpdateStatus($"错误：区域 '{rangeStr}' 无效。{ex.Message}");
                            continue; // 继续处理下一个区域
                        }
                    }
                    
                    // 设置数据源
                    dataGridView.DataSource = dataTable;
                    
                    // 更新窗口标题显示总记录数
                    previewForm.Text = $"数据预览 - 共 {totalRows} 条记录";
                    
                    // 添加关闭按钮
                    Button btnClose = new Button
                    {
                        Text = "关闭",
                        Dock = DockStyle.Bottom,
                        Height = 30
                    };
                    btnClose.Click += (s, ev) => previewForm.Close();
                    
                    // 添加控件到预览窗口
                    previewForm.Controls.Add(dataGridView);
                    previewForm.Controls.Add(btnClose);
                    
                    // 显示预览窗口
                    previewForm.ShowDialog(this);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"预览数据时发生错误：{ex.Message}");
                MessageBox.Show($"预览失败：{ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// 获取Excel列名（如A, B, C...AA, AB等）
        /// </summary>
        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        
        /// <summary>
        /// 错误日志按钮点击事件
        /// </summary>
        private void BtnErrorLog_Click(object? sender, EventArgs e)
        {   
            // 显示提示信息
            MessageBox.Show("错误日志功能将在后续版本中提供", "提示", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        /// <summary>
        /// 验证用户输入是否有效
        /// </summary>
        private bool ValidateInput()
        {
            if (string.IsNullOrEmpty(_currentFilePath))
            {
                UpdateStatus("错误：请先选择Excel文件");
                return false;
            }
            
            if (cmbSheetNames.SelectedIndex == -1)
            {
                UpdateStatus("错误：请选择工作表");
                return false;
            }
            
            if (string.IsNullOrEmpty(txtRange.Text))
            {
                UpdateStatus("错误：请输入数据区域范围");
                return false;
            }
            
            return true;
        }
        
        /// <summary>
        /// 获取用户选择的数据区域
        /// </summary>
        private ExcelRange? GetUserSelectedRange(ExcelWorksheet worksheet)
        {
            try
            {
                if (worksheet == null || string.IsNullOrEmpty(txtRange.Text))
                {
                    UpdateStatus("错误：工作表或数据区域无效");
                    return null;
                }
                return worksheet.Cells[txtRange.Text];
            }
            catch (Exception ex)
            {
                UpdateStatus($"错误：无效的数据区域范围。{ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// 保存处理后的文件
        /// </summary>
        private string SaveProcessedFile()
        {
            if (string.IsNullOrEmpty(_currentFilePath))
                throw new InvalidOperationException("文件路径不能为空");
                
            string directory = Path.GetDirectoryName(_currentFilePath) ?? Directory.GetCurrentDirectory();
            string fileName = Path.GetFileNameWithoutExtension(_currentFilePath) ?? "输出文件";
            string extension = Path.GetExtension(_currentFilePath) ?? ".xlsx";
            
            if (string.IsNullOrEmpty(directory) || string.IsNullOrEmpty(fileName))
                throw new InvalidOperationException("无法获取有效的文件路径信息");
                
            string outputPath = Path.Combine(directory, $"{fileName}_已恢复{extension}");
            
            // 检查文件是否已存在
            int counter = 1;
            string basePath = outputPath;
            while (File.Exists(outputPath))
            {
                outputPath = Path.Combine(directory, 
                    $"{fileName}_已恢复{counter}{extension}");
                counter++;
            }
            
            // 保存文件
            _excelPackage.SaveAs(new FileInfo(outputPath));
            return outputPath;
        }

        private void UpdateStatus(string message)
        {    
            message = message ?? string.Empty;
            
            // 确保在UI线程中更新
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(UpdateStatus), message);
                return;
            }
            
            txtStatus.Text = message;
            toolStripStatusLabel1.Text = message.Split('\n').Last();
        }
        
        #region 数据模型类
        
        /// <summary>
        /// 处理数据的参数辅助类
        /// </summary>
        private class ProcessDataArgs
        {   
            private string _filePath = string.Empty;
            private string _sheetName = string.Empty;
            private string _rangeString = string.Empty;
            
            public string FilePath 
            { 
                get => _filePath; 
                set => _filePath = value ?? string.Empty; 
            }
            
            public string SheetName 
            { 
                get => _sheetName; 
                set => _sheetName = value ?? string.Empty; 
            }
            
            public string RangeString 
            { 
                get => _rangeString; 
                set => _rangeString = value ?? string.Empty; 
            }
            
            /// <summary>
            /// 获取多个数据区域（用逗号分隔）
            /// </summary>
            public List<string> GetRangeStrings()
            {   
                return RangeString.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(r => r.Trim())
                    .Where(r => !string.IsNullOrWhiteSpace(r))
                    .ToList();
            }
            
            /// <summary>
            /// 静态方法：获取多个数据区域（用逗号分隔）
            /// </summary>
            public static List<string> GetRangeStrings(string rangeString)
            {   
                return rangeString.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(r => r.Trim())
                    .Where(r => !string.IsNullOrWhiteSpace(r))
                    .ToList();
            }
        }
        
        /// <summary>
        /// 处理结果类
        /// </summary>
        private class ProcessResult
        {
            public bool Success { get; set; } = false;
            public string ErrorMessage { get; set; } = string.Empty;
            public int ProcessedCount { get; set; } = 0;
            public int SkippedCount { get; set; } = 0;
            public string OutputPath { get; set; } = string.Empty;
            public string ErrorSummary { get; set; } = string.Empty;
        }
        
        #endregion

        private void ClearSheetSelection()
        {
            cmbSheetNames.Items.Clear();
            btnProcess.Enabled = false;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 释放Excel包资源
            if (_excelPackage != null)
            {
                _excelPackage.Dispose();
            }
        }

        // UI组件字段
        private TextBox txtFilePath = new TextBox();
        private ComboBox cmbSheetNames = new ComboBox();
        private TextBox txtRange = new TextBox();
        private Button btnProcess = new Button();
        private Button btnPreviewData = new Button();
        private TextBox txtStatus = new TextBox();
        private ProgressBar progressBar1 = new ProgressBar();
        private StatusStrip statusStrip1 = new StatusStrip();
        private ToolStripStatusLabel toolStripStatusLabel1 = new ToolStripStatusLabel();
        private ToolTip toolTip1 = new ToolTip();
        private ToolStripButton btnErrorLog = new ToolStripButton();
    }
}