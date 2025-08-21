using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Button_interface
{
    public partial class Form1 : Form
    {
        [DllImport("kernel32.dll")]
        static extern bool AllocConsole();

        private readonly Dictionary<int, bool> columnSortDirections = new Dictionary<int, bool>();
        private static int rainbowColorIndex = 0;
        private readonly Color[] rainbowColors = new[]
        {
            Color.Red, Color.Orange, Color.Yellow, Color.Green, Color.Cyan, Color.Blue, Color.Purple
        };

        private readonly string[] headers = new[]
        {
            "Group", "Category", "Product/Project Name", "Product Innovation Type",
            "Phase", "Condition", "Status", "OBS Date", "Next Gate", "CPM",
            "PE", "CDM", "MPE", "DE", "Awarded Vendor", "Potential Vendor", "Material Number"
        };

        public Form1()
        {
            InitializeComponent();
            SetupLayout();
            AllocConsole();
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            for (int i = 0; i < 17; i++)
            {
                columnSortDirections[i] = true;
            }
        }

        private readonly string shortcutPath = @"C:\Users\jacky.he\OneDrive - Kids II\Desktop\RealTime 2.2 V09.10.023 - 3.lnk";
        private readonly string processName = "RealTime";

        private void SetupLayout()
        {
            this.Text = "Fake RealTime";
            this.Size = new System.Drawing.Size(1800, 600);
            this.BackColor = System.Drawing.Color.LightGray;
            this.AutoScaleMode = AutoScaleMode.None;

            TableLayoutPanel tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 17,
                RowCount = 15,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.None, // ← 去掉边框
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            for (int i = 0; i < 17; i++)
            {
                tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / 17));
            }

            for (int i = 0; i < 15; i++)
            {
                if (i == 1)
                {
                    tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 150f / 15));
                }
                else
                {
                    tableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f / 15));
                }
            }

            for (int col = 0; col < 17; col++)
            {
                AddSortButtonToTable(tableLayout, headers[col], col, 0, Color.Red, Color.White);
            }

            for (int col = 0; col < 17; col++)
            {
                AddComboBoxToTable(tableLayout, "", col, 1, Color.White, Color.Black);
            }

            string[] items = { "11", "AA", "33", "DD", "55", "FF", "77", "H6", "1L1L" };
            for (int row = 2; row <= 10; row++)
            {
                for (int col = 0; col < 17; col++)
                {
                    AddButtonToTable(tableLayout, items[row - 2], col, row, Color.LightGreen, Color.Black);
                }
            }

            AddButtonToTable(tableLayout, "Select All", 0, 11, Color.Blue, Color.White);
            AddButtonToTable(tableLayout, "UnSelect All", 1, 11, Color.LightGray, Color.Black);
            AddButtonToTable(tableLayout, "Total: 504, 1 Cells Selected", 2, 11, Color.Pink, Color.Black);
            AddButtonToTable(tableLayout, "Loan Title", 3, 11, Color.Pink, Color.Black);
            AddButtonToTable(tableLayout, "Kanban All", 4, 11, Color.LightSalmon, Color.Black);
            AddButtonToTable(tableLayout, "Prod Sync 1080", 5, 11, Color.LightYellow, Color.Black);
            AddButtonToTable(tableLayout, "Name", 6, 11, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "Late Tasks", 7, 11, Color.Red, Color.White);
            AddButtonToTable(tableLayout, "Key Milestones", 8, 11, Color.Yellow, Color.Black);
            AddButtonToTable(tableLayout, "Gate Submit", 9, 11, Color.LightYellow, Color.Black);
            AddButtonToTable(tableLayout, "Update Data", 10, 11, Color.Pink, Color.Black);
            AddButtonToTable(tableLayout, "All Actions", 11, 11, Color.Magenta, Color.White);
            AddButtonToTable(tableLayout, "WF", 12, 11, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "E2E", 13, 11, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "", 14, 11, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "", 15, 11, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "", 16, 11, Color.White, Color.Black);

            AddButtonToTable(tableLayout, "82", 0, 12, Color.Pink, Color.Black);
            AddButtonToTable(tableLayout, "Kan Forecast", 1, 12, Color.Red, Color.White);
            AddButtonToTable(tableLayout, "ProdSync 2K", 2, 12, Color.Red, Color.White);
            AddButtonToTable(tableLayout, "Assignee", 3, 12, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "Late & Upcoming", 4, 12, Color.Yellow, Color.Black);
            AddButtonToTable(tableLayout, "Timelines", 5, 12, Color.Blue, Color.White);
            AddButtonToTable(tableLayout, "Material Status", 6, 12, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "Get Material No", 7, 12, Color.Red, Color.White);
            AddButtonToTable(tableLayout, "Action List", 8, 12, Color.Red, Color.White);
            AddButtonToTable(tableLayout, "Open", 9, 12, Color.LightGray, Color.Black);
            AddButtonToTable(tableLayout, "Tool Release Check", 10, 12, Color.Blue, Color.White);
            AddButtonToTable(tableLayout, "PEM Timeline", 11, 12, Color.Blue, Color.White);
            AddButtonToTable(tableLayout, "", 12, 12, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "", 13, 12, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "", 14, 12, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "", 15, 12, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "", 16, 12, Color.White, Color.Black);

            AddButtonToTable(tableLayout, "Refresh the List", 0, 13, Color.LightGreen, Color.Black);
            AddButtonToTable(tableLayout, "Summary", 1, 13, Color.Khaki, Color.Black);
            AddButtonToTable(tableLayout, "Kanban 4 - Cost", 2, 13, Color.LightCoral, Color.Black);
            AddButtonToTable(tableLayout, "Core Team Deck", 3, 13, Color.LightPink, Color.Black);
            AddButtonToTable(tableLayout, "WF Project", 4, 13, Color.LightCyan, Color.Black);
            AddButtonToTable(tableLayout, "Find Projects", 5, 13, Color.LightBlue, Color.Black);
            AddButtonToTable(tableLayout, "All Open Tasks", 6, 13, Color.LightCyan, Color.Black);
            AddButtonToTable(tableLayout, "Time & Owner", 7, 13, Color.Blue, Color.White);
            AddButtonToTable(tableLayout, "MSRP FOB", 8, 13, Color.Black, Color.White);
            AddButtonToTable(tableLayout, "Business Brief", 9, 13, Color.Orange, Color.Black);
            AddButtonToTable(tableLayout, "All Action List", 10, 13, Color.LightCyan, Color.Black);
            AddButtonToTable(tableLayout, "Sample Tracker", 11, 13, Color.Blue, Color.White);
            AddButtonToTable(tableLayout, "Gate Deliverables", 12, 13, Color.Pink, Color.Black);
            AddButtonToTable(tableLayout, "Get Values", 13, 13, Color.Pink, Color.Black);
            AddButtonToTable(tableLayout, "", 14, 13, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "", 15, 13, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "", 16, 13, Color.White, Color.Black);

            AddButtonToTable(tableLayout, "Choose Views", 0, 14, Color.Pink, Color.Black);
            AddButtonToTable(tableLayout, "Load All Projects", 1, 14, Color.Yellow, Color.Black);
            AddButtonToTable(tableLayout, "Start RT", 2, 14, Color.Red, Color.White);
            AddButtonToTable(tableLayout, "Kanban 2", 3, 14, Color.LightSalmon, Color.Black);
            AddButtonToTable(tableLayout, "Budget - All", 4, 14, Color.LightBlue, Color.Black);
            AddButtonToTable(tableLayout, "Vendor Deck", 5, 14, Color.Yellow, Color.Black);
            AddButtonToTable(tableLayout, "Folder List", 6, 14, Color.LightCyan, Color.Black);
            AddButtonToTable(tableLayout, "Assignment", 7, 14, Color.Red, Color.White);
            AddButtonToTable(tableLayout, "Smart Sheet (TBD)", 8, 14, Color.Purple, Color.White);
            AddButtonToTable(tableLayout, "PKG Timeline", 9, 14, Color.Blue, Color.White);
            AddButtonToTable(tableLayout, "Brief Summary", 10, 14, Color.Orange, Color.Black);
            AddButtonToTable(tableLayout, "PAE Test Lab", 11, 14, Color.Black, Color.White);
            AddButtonToTable(tableLayout, "LRB Timeline", 12, 14, Color.Blue, Color.White);
            AddButtonToTable(tableLayout, "Plan VS Actual", 13, 14, Color.Red, Color.White);
            AddButtonToTable(tableLayout, "Kids2 EE Site", 14, 14, Color.Red, Color.White);
            AddButtonToTable(tableLayout, "", 15, 14, Color.White, Color.Black);
            AddButtonToTable(tableLayout, "", 16, 14, Color.White, Color.Black);

            this.Controls.Add(tableLayout);
        }

        private void AddButtonToTable(TableLayoutPanel table, string text, int col, int row, Color backColor, Color foreColor)
        {
            // 如果背景色是白色，使用彩虹色
            Color buttonBackColor = backColor == Color.White
                ? rainbowColors[rainbowColorIndex++ % rainbowColors.Length]
                : backColor;

            Button button = new Button
            {
                Text = text,
                BackColor = buttonBackColor,
                ForeColor = foreColor,
                FlatStyle = FlatStyle.Flat,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(0),
                Font = new Font("Microsoft Sans Serif", 13.5f)
            };
            button.Click += ClickButton;
            table.Controls.Add(button, col, row);
        }
        private void AddSortButtonToTable(TableLayoutPanel table, string text,
                                   int col, int row,
                                   System.Drawing.Color backColor,
                                   System.Drawing.Color foreColor)
        {
            Button btnHeader = new Button
            {
                Text = text + " (ASC)",
                BackColor = backColor,
                ForeColor = foreColor,
                FlatStyle = FlatStyle.Flat,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(0),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 13.5f),
                Tag = col
            };

            btnHeader.Click += (s, e) =>
            {
                if (!(s is Button hdr && hdr.Tag is int columnIndex)) return;

                // 1. 取得当前“可见”按钮的文本
                var visibleItems = new System.Collections.Generic.List<string>();
                foreach (Control c in table.Controls)
                {
                    if (c is SearchableComboBox sc && table.GetColumn(c) == columnIndex)
                    {
                        visibleItems.AddRange(sc.GetVisibleItems());
                        break;
                    }
                }
                if (visibleItems.Count == 0) return;   // 没有任何可见按钮

                // 2. 排序（自然排序）
                bool isAscending = columnSortDirections[columnIndex];
                columnSortDirections[columnIndex] = !isAscending;
                string dir = isAscending ? "DESC" : "ASC";
                hdr.Text = text + $" ({dir})";
                Console.WriteLine($"排序列 {columnIndex} ({text}): {dir}");

                visibleItems.Sort(new NaturalStringComparer());
                if (!isAscending) visibleItems.Reverse();

                // 3. 只把结果写回“可见”按钮
                int idx = 0;
                for (int r = 2; r <= 10; r++)
                {
                    var cell = table.GetControlFromPosition(columnIndex, r);
                    if (cell is Button b && b.Visible)
                    {
                        b.Text = visibleItems[idx++];
                    }
                }

                // 4. 同步 SearchableComboBox 的数据
                foreach (Control c in table.Controls)
                {
                    if (c is SearchableComboBox sc && table.GetColumn(c) == columnIndex)
                    {
                        sc.SetSource(visibleItems.ToArray());
                        break;
                    }
                }
            };

            table.Controls.Add(btnHeader, col, row);
        }

        //private void AddSortButtonToTable(TableLayoutPanel table, string text, int col, int row, Color backColor, Color foreColor)
        //{
        //    Button button = new Button
        //    {
        //        Text = text + " (ASC)",
        //        BackColor = backColor,
        //        ForeColor = foreColor,
        //        FlatStyle = FlatStyle.Flat,
        //        Dock = DockStyle.Fill,
        //        Margin = new Padding(0),
        //        Padding = new Padding(0),
        //        Font = new Font("Microsoft Sans Serif", 13.5f),
        //        Tag = col
        //    };
        //    button.Click += (sender, e) =>
        //    {
        //        if (sender is Button btn && btn.Tag is int columnIndex)
        //        {
        //            bool isAscending = columnSortDirections[columnIndex];
        //            columnSortDirections[columnIndex] = !isAscending;
        //            string sortDirection = isAscending ? "DESC" : "ASC";
        //            btn.Text = text + $" ({sortDirection})";
        //            Console.WriteLine($"排序列 {columnIndex} ({text}): {sortDirection}");
        //            btn.ForeColor = Color.Aqua;
        //            btn.BackColor = Color.White;

        //            foreach (Control control in table.Controls)
        //            {
        //                if (control is Button otherBtn && otherBtn.Tag is int otherCol && otherBtn != btn && table.GetRow(otherBtn) == 0)
        //                {
        //                    columnSortDirections[otherCol] = true;
        //                    otherBtn.Text = headers[otherCol] + " (ASC)";
        //                    otherBtn.ForeColor = Color.White;
        //                    otherBtn.BackColor = Color.Red;
        //                }
        //            }

        //            // 收集行2-10的按钮文本
        //            var items = new List<string>();
        //            for (int r = 2; r <= 10; r++)
        //            {
        //                var control = table.GetControlFromPosition(columnIndex, r);
        //                if (control is Button btnInColumn)
        //                {
        //                    items.Add(btnInColumn.Text);
        //                }
        //            }

        //            // 使用自定义比较器排序
        //            items.Sort(new NaturalStringComparer());
        //            if (!isAscending)
        //            {
        //                items.Reverse();
        //            }

        //            // 更新行2-10的按钮文本
        //            for (int r = 2; r <= 10; r++)
        //            {
        //                var control = table.GetControlFromPosition(columnIndex, r);
        //                if (control is Button btnInColumn)
        //                {
        //                    btnInColumn.Text = items[r - 2];
        //                }
        //            }

        //            // 更新行1的ComboBox项目
        //            foreach (Control control in table.Controls)
        //            {
        //                if (control is ComboBox cb && table.GetRow(cb) == 1 && table.GetColumn(cb) == columnIndex)
        //                {
        //                    cb.Items.Clear();
        //                    cb.Items.AddRange(items.Cast<object>().ToArray());
        //                    break;
        //                }
        //            }
        //        }
        //    };
        //    table.Controls.Add(button, col, row);
        //}

        private class NaturalStringComparer : IComparer<string>
        {
            public int Compare(string? x, string? y)
            {
                if (x == null && y == null) return 0;
                if (x == null) return -1;
                if (y == null) return 1;

                // 使用正则表达式提取数字和非数字部分
                var regex = new Regex(@"^(\d+)(.*)$");
                var matchX = regex.Match(x);
                var matchY = regex.Match(y);

                if (matchX.Success && matchY.Success)
                {
                    // 提取数字和剩余部分
                    if (int.TryParse(matchX.Groups[1].Value, out int numX) && int.TryParse(matchY.Groups[1].Value, out int numY))
                    {
                        if (numX != numY)
                        {
                            return numX.CompareTo(numY);
                        }
                        // 如果数字相同，比较剩余部分
                        return StringComparer.OrdinalIgnoreCase.Compare(matchX.Groups[2].Value, matchY.Groups[2].Value);
                    }
                }

                // 如果不是数字开头，按字典序比较
                return StringComparer.OrdinalIgnoreCase.Compare(x, y);
            }
        }
        private void AddComboBoxToTable(TableLayoutPanel table, string text,
                                int col, int row,
                                System.Drawing.Color backColor,
                                System.Drawing.Color foreColor)
        {
            var cb = new SearchableComboBox(table, col)
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),      // ← 关键
                Padding = new Padding(0),     // ← 再加一层保险
                Font = new Font("Microsoft Sans Serif", 18f),
                BackColor = backColor,
                ForeColor = foreColor
            };

            string[] items = { "11", "AA", "33", "DD", "55", "FF", "77", "H6", "1L1L" };
            cb.SetSource(items);

            table.Controls.Add(cb, col, row);
        }

        //private void AddComboBoxToTable(TableLayoutPanel table, string text, int col, int row, Color backColor, Color foreColor)
        //{
        //    ComboBox comboBox = new ComboBox
        //    {
        //        Text = text,
        //        BackColor = backColor,
        //        ForeColor = foreColor,
        //        FlatStyle = FlatStyle.Flat,
        //        Dock = DockStyle.Fill,
        //        Margin = new Padding(0),
        //        Padding = new Padding(0),
        //        Font = new Font("Microsoft Sans Serif", 18f),
        //        DropDownStyle = ComboBoxStyle.DropDown
        //    };
        //    string[] items = { "1A", "B", "3C", "D", "5E", "F", "7G", "H", "9I" };
        //    comboBox.Items.AddRange(items.Cast<object>().ToArray());
        //    comboBox.TextChanged += (sender, e) =>
        //    {
        //        if (sender is ComboBox cb)
        //        {
        //            Console.WriteLine($"输入: {cb.Text}");
        //            cb.ForeColor = Color.Aqua;
        //            cb.BackColor = Color.White;
        //        }
        //    };
        //    comboBox.SelectedIndexChanged += (sender, e) =>
        //    {
        //        if (sender is ComboBox cb)
        //        {
        //            Console.WriteLine($"选择: {cb.Text}");
        //            cb.ForeColor = Color.Aqua;
        //            cb.BackColor = Color.White;
        //        }
        //    };
        //    table.Controls.Add(comboBox, col, row);
        //}

        private void ClickButton(object? sender, EventArgs e)
        {
            Button? clickedButton = sender as Button;
            if (clickedButton != null)
            {
                Console.WriteLine($"点击: {clickedButton.Text}");
                clickedButton.ForeColor = Color.Aqua;
                clickedButton.BackColor = Color.White;

                if (clickedButton.Text == "WF")
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = "https://kidsii.my.workfront.com/my-updates",
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"打开 WF 网页时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (clickedButton.Text == "E2E")
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = "https://kidsii.sharepoint.com/sites/Dept_IT_BusinessTechnology/Shared%20Documents/Forms/AllItems.aspx",
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"打开 E2E 网页时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (clickedButton.Text == "Start RT")
                {
                    Process[] isOpen = Process.GetProcessesByName(processName);
                    if (isOpen.Length > 0)
                    {
                        MessageBox.Show("请先关闭 RealTime 再重新启动。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    else
                    {
                        MessageBox.Show("RealTime 未运行，将继续启动。", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = shortcutPath,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"打开快捷方式时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}