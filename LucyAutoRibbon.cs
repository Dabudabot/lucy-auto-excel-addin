using System;
#if DEBUG
using System.Diagnostics;
#endif
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace LucyAutoExAddIn
{
  public partial class LucyAutoRibbon
  {
    private void LucyAutoRibbon_Load(object sender, RibbonUIEventArgs e)
    {

    }

    private bool IsValidCell(string cell_str)
    {
      bool charMode = true;
      foreach (var c in cell_str)
      {
        if (charMode)
        {
          if (char.IsLetter(c)) continue;
          else
          {
            charMode = false;
            if (!char.IsDigit(c)) return false;
          }
        }
        else
        {
          if (!char.IsDigit(c)) return false;
        }
      }

#if DEBUG
      Debug.WriteLine(cell_str + " is valid cell");
#endif
      return true;
    }

    private bool IsValidRange(string range_str)
    {
      var substr = range_str.Split(':');

      if (2 != substr.Length) return false;
      if (!IsValidCell(substr[0]) || !IsValidCell(substr[1])) return false;

#if DEBUG
      Debug.WriteLine(range_str + " is valid range");
#endif
      return true;
    }

    private Worksheet GetSheetByName(string name)
    {
      foreach (Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
      {
        if (name.ToLower().Trim() == sheet.Name.ToLower().Trim())
        {
          return sheet;
        }
      }

      return null;
    }

    private Chart GetChartByName(string name, Worksheet sheet)
    {
      string fix;
      if (name.Length > 32)
      {
        fix = name.Substring(0, 32);
      }
      else
      {
        fix = name;
      }
      fix = fix.ToLower().Trim();

      foreach (ChartObject co in sheet.ChartObjects())
      {
        if (fix == co.Name.ToLower().Trim())
        {
          return co.Chart;
        }
      }

      return null;
    }

    private Range FindCell(string text, Worksheet sheet)
    {
      Range searchRange = sheet.get_Range("A1", "AAA999");
      Range foundRange = searchRange.Find(text, Type.Missing,
            XlFindLookIn.xlValues, XlLookAt.xlWhole,
            XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false,
            Type.Missing, Type.Missing);
      return foundRange;
    }

    private Range FindLastCell(Range cell)
    {
      var result = cell.Offset[1, 0];

      while (null != result.Value2 && 0 != ((string)result.Value2).Trim().Length)
      {
        result = result.Offset[1, 0];
      }

      result = result.Offset[-1, 0];

      return result;
    }

    private bool UpdateDate(Range fromRange, Range toRange, string suffix)
    {
      string fromStr = fromRange.Value2;
      int sector = fromStr[0] - '0';
      string yearStr = fromStr.Substring(6);

      if (!int.TryParse(yearStr, out int year))
      {
        return false;
      }

      sector++;

      if (sector > 4)
      {
        sector = 1;
        year++;
      }

      toRange.Value2 = sector.ToString() + " кв. " + year.ToString() + suffix;

      return true;
    }

    private string UpdateChartFormula(string source)
    {
      source = source.Trim();
      string result = "";
      int begin = 0;
      bool skip = false;
      Regex orderRegex = new Regex(@"\,\d+\)");
      Regex rangeRegex = new Regex(@"\$[A-Za-z]\$\d+\:\$[A-Za-z]\$\d+");
      Regex numberRegex = new Regex(@"\d+");

      Match orderMatch = orderRegex.Match(source);

      if (!orderMatch.Success) return null;

      if (!int.TryParse(orderMatch.Value.Substring(1, orderMatch.Value.Length - 2), out int order))
      {
        return null;
      }

      if (1 != order) skip = true;

      MatchCollection rangeMatches = rangeRegex.Matches(source);

      foreach (Match rangeMatch in rangeMatches)
      {
        result += source.Substring(begin, rangeMatch.Index - begin);
        begin = (rangeMatch.Index + rangeMatch.Length);
        MatchCollection numberMatches = numberRegex.Matches(rangeMatch.Value);
        if (numberMatches.Count != 2)
        {
          return null;
        }

        result += rangeMatch.Value.Substring(0, numberMatches[0].Index);

        if (!int.TryParse(numberMatches[0].Value, out int number1))
        {
          return null;
        }

        if (!int.TryParse(numberMatches[1].Value, out int number2))
        {
          return null;
        }

        if (!skip)
        {
          number1++;
          number2++;

          if (number2 - number1 < 4)
          {
            number1--;
          }
        }
        else
        {
          skip = false;
        }

        result += number1.ToString();
        result += rangeMatch.Value.Substring(
          numberMatches[0].Index + numberMatches[0].Length,
          numberMatches[1].Index - (numberMatches[0].Index + numberMatches[0].Length));
        result += number2.ToString();
      }

      result += source.Substring(begin);

      return result;
    }

    private void UpdateCellValue(Range sourceRowsRange, Range sourceColsRange, int rowNumber, int colNumber, Range targetCell, int targetRightOffset)
    {
      Range sourceIntersectionCell = Globals.ThisAddIn.Application.Cells[((Range)sourceRowsRange.Item[rowNumber]).Row, sourceColsRange.Item[colNumber].Column];
      dynamic sourceValue = sourceIntersectionCell.Value2;
      var targetValueCell = targetCell.Offset[0, targetRightOffset]; // move right
      targetValueCell.Value2 = sourceValue;
      targetValueCell.NumberFormat = sourceIntersectionCell.NumberFormat;

#if DEBUG
      Debug.WriteLine("Set new value: " + (object)targetValueCell.Value2 + " to cell: " + targetValueCell.Row + " - " + targetValueCell.Column);
#endif
    }

    private void RunProcess(Range sheetNamesRange, Range itemNamesRange, int jumpAmount, bool append, string suffix)
    {
      int progress = 0;

      foreach (Range sheetName in sheetNamesRange)
      {
        ProgressBar.Label = "Progress: " + progress.ToString() + " / " + sheetNamesRange.Count;

        string sheetNameStr = sheetName.Value2;

#if DEBUG
        Debug.WriteLine("Target selected: " + sheetNameStr);
#endif
        Worksheet targetSheet = GetSheetByName(sheetNameStr);

        if (null == targetSheet)
        {
          MessageBox.Show("Failed to find sheet with name = " + sheetNameStr, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
          continue;
        }

        for (var i = 1; i <= itemNamesRange.Count; i += (jumpAmount+1))
        {
          string lookingText = ((Range)itemNamesRange.Item[i]).Value2;

#if DEBUG
          Debug.WriteLine("Target text: " + lookingText);
#endif

          // find cell with this text on targetSheet
          var lookingTextCell = FindCell(lookingText, targetSheet);

          if (null == lookingTextCell)
          {
            MessageBox.Show("Failed to find text: " + lookingText + " on sheet "  + sheetNameStr, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            continue;
          }

#if DEBUG
          Debug.WriteLine("Found cell: " + lookingTextCell.Row + " - " + lookingTextCell.Column + " with text: " + lookingText);
#endif

          var lastCell = FindLastCell(lookingTextCell.Offset[1, 0]);

          if (null == lastCell)
          {
            MessageBox.Show("Failed to find last avaliable cell for " + lookingText + " on sheet " + sheetNameStr, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            continue;
          }

#if DEBUG
          Debug.WriteLine("Found last cell: " + lastCell.Row + " - " + lastCell.Column + " with text: " + (string)lastCell.Value2);
#endif

          lastCell.Value2 = ((string)lastCell.Value2).Trim(new char[] { 'П' });

#if DEBUG
          Debug.WriteLine("Now value may be updated due to remove char: " + (string)lastCell.Value2);
#endif

          Range targetCell;

          if (append)
          {
            var downCell = lastCell.Offset[1, 0]; // move down
            downCell.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown);
            targetCell = downCell.Offset[-1, 0]; // move up

            if (!UpdateDate(lastCell, targetCell, suffix))
            {
              MessageBox.Show("Failed to update date = " + (string)lastCell.Value2, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
              continue;
            }
          }
          else
          {
            targetCell = lastCell;
          }

#if DEBUG
          Debug.WriteLine("Target cell: " + targetCell.Row + " - " + targetCell.Column + " with text: " + (string)targetCell.Value2);
#endif
          if (0 == jumpAmount)
          {
            UpdateCellValue(itemNamesRange, sheetName, i, 1, targetCell, 1);
          }
          else
          {
            for (var j = 1; j <= jumpAmount; j++)
            {
              UpdateCellValue(itemNamesRange, sheetName, i + j, 1, targetCell, j);
            }
          }

          if (append)
          {
            // get graph
            var chart = GetChartByName(lookingText, targetSheet);

            if (null == chart)
            {
              MessageBox.Show("Failed to find chart " + lookingText + " on sheet " + sheetNameStr, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
              continue;
            }

            foreach (Series sc in chart.SeriesCollection())
            {
#if DEBUG
              Debug.WriteLine("Old formula: " + sc.Formula);
#endif
              sc.Formula = UpdateChartFormula(sc.Formula);

#if DEBUG
              Debug.WriteLine("New formula: " + sc.Formula);
#endif
            }
          }
        }

        progress++;
      }
    }

    private void Run_Click(object sender, RibbonControlEventArgs e)
    {
#if DEBUG
      Debug.WriteLine("ListsRangeBox value: " + SheetsRangeBox.Text);
      Debug.WriteLine("CellsRangeBox value: " + ItemsRangeBox.Text);
      Debug.WriteLine("JumpAmountBox value: " + JumpAmountBox.Text);
      Debug.WriteLine("DateSuffixBox value: " + DateSuffixBox.Text);
      Debug.WriteLine("AppendBox value: " + AppendBox.Checked);
#endif

      if (!IsValidRange(SheetsRangeBox.Text))
      {
        MessageBox.Show("Lists range are wrong!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
      }

      if (!IsValidRange(ItemsRangeBox.Text))
      {
        MessageBox.Show("Cells range are wrong!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
      }

      if (!int.TryParse(JumpAmountBox.Text, out int jumpAmountValue))
      {
        MessageBox.Show("Jump amount are wrong!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
      }

#if DEBUG
      Debug.WriteLine("Input parameters are OK!");
#endif

      RunBtn.Visible = false;
      ProgressBar.Visible = true;

      RunProcess(Globals.ThisAddIn.Application.get_Range(SheetsRangeBox.Text), 
        Globals.ThisAddIn.Application.get_Range(ItemsRangeBox.Text), 
        jumpAmountValue, 
        AppendBox.Checked,
        DateSuffixBox.Text
      );

      RunBtn.Visible = true;
      ProgressBar.Visible = false;
    }
  }
}

