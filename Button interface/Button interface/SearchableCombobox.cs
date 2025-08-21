using System;
using System.Linq;
using System.Windows.Forms;

public class SearchableComboBox : ComboBox
{
    private string[] _originalSource = Array.Empty<string>();
    private readonly TableLayoutPanel _parentTable;
    private readonly int _columnIndex;

    public SearchableComboBox(TableLayoutPanel parentTable, int columnIndex)
    {
        _parentTable = parentTable;
        _columnIndex = columnIndex;

        DropDownStyle = ComboBoxStyle.DropDown;
        AutoCompleteMode = AutoCompleteMode.None;
    }

    public void SetSource(string[] source)   // 外部调用
    {
        _originalSource = source;
        Items.Clear();
        Items.AddRange(source);
    }

    /* ↓↓↓ 新增：把当前可见按钮的文本返回给排序按钮 ↓↓↓ */
    public string[] GetVisibleItems()
    {
        return Enumerable.Range(2, 9)                        // 行 2-10
               .Select(r => _parentTable.GetControlFromPosition(_columnIndex, r))
               .OfType<Button>()
               .Where(b => b.Visible)
               .Select(b => b.Text)
               .ToArray();
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            e.Handled = true;
            DoFilter();
            return;
        }
        base.OnKeyDown(e);
    }

    private void DoFilter()
    {
        string keyword = Text.Trim();
        bool showAll = string.IsNullOrEmpty(keyword);

        for (int r = 2; r <= 10; r++)
        {
            var ctl = _parentTable.GetControlFromPosition(_columnIndex, r);
            if (ctl is Button btn)
            {
                bool visible = showAll ||
                               btn.Text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;
                btn.Visible = visible;
            }
        }
    }
}