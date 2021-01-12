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
  // all ribbon magic happens here

  public partial class LucyAutoRibbon
  {
    private void LucyAutoRibbon_Load(object sender, RibbonUIEventArgs e)
    {
      // this function called as ribbon loaded
    }

    private string FixUpType(string type, string value)
    {
      if (type == "Shift")
      {
        int.TryParse(value, out int valueInt);
        if (0 == valueInt)
        {
          return "Append";
        }
      }
      return type;
    }

    private bool IsValidLookup(string type, string value)
    {
      // sorry for such code, must use enums
      if (type != "Append" && type != "Exact" && type != "Shift")
      {
        return false;
      }

      if (type == "Shift")
      {
        return int.TryParse(value, out int valueInt);
      }

      // may be exact value also has to be cheched

      return true;
    }

    private Range FindNextCellByType(string type, string value, Range source)
    {
      if (type == "Shift")
      {
        int.TryParse(value, out int valueInt);

        if (valueInt > 0)
        {
          // search from the top
          var result = source;

          for (int i = 0; i < valueInt; i++)
          {
            result = result.Offset[1, 0];
          }

          return result;
        }
        else
        {
          // search from bottom

          var result = FindLastCell(source);

          for (int i = valueInt+1; i < 0; i++)
          {
            result = result.Offset[-1, 0];
          }

          return result;
        }
      }

      if (type == "Exact")
      {
        var result = source.Offset[1, 0];

        while (!((string)result.Value2).Contains(value))
        {
          result = result.Offset[1, 0];
          if (result.Value2 == null)
          {
            // value was not found
            // it will return null, send error message
            
            break;
          }
        }

        return result;
      }

      // append by default
      return FindLastCell(source);
    }

    private bool IsValidCell(string cell_str)
    {
      // this checks if string represents valid cell
      // for example A1 or B42 is valid, but A2A not

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
      // checks if range represented by string is valid
      // valid range example A3:B6

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
      // iterate over all sheets in document and return sheet with given name

      foreach (Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
      {
        // name ignores case and delete spaces
        if (name.ToLower().Trim() == sheet.Name.ToLower().Trim())
        {
          return sheet;
        }
      }

      // callee have to check return value of this function
      return null;
    }

    private Chart GetChartByName(string name, Worksheet sheet)
    {
      // iterate over all charts on given sheet and return chart with name
      // chart`s name in excel is limited by 32 symbols so first we cut the
      // name if it is longer

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

      // callee have to check return value of this function
      return null;
    }

    private Range FindCell(string text, Worksheet sheet)
    {
      // find cell with specific text in it
      // text have to match exactly

      Range searchRange = sheet.get_Range("A1", "AAA999");
      Range foundRange = searchRange.Find(text, Type.Missing,
            XlFindLookIn.xlValues, XlLookAt.xlWhole,
            XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false,
            Type.Missing, Type.Missing);
      return foundRange;
    }

    private Range FindLastCell(Range cell)
    {
      // iterating down until empty cell

      var result = cell.Offset[1, 0];

      // cell value can be null or it can contain empty text
      while (null != result.Value2 && 0 != ((string)result.Value2).Trim().Length)
      {
        result = result.Offset[1, 0];
      }

      // get back to last fulfilled cell
      result = result.Offset[-1, 0];

      return result;
    }

    private bool UpdateDate(Range fromRange, Range toRange, string suffix)
    {
      // update cell with date of format 1 кв. 20
      // to move it to next date

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
    private string UpdateChartFormulaPart(string source)
    {
      // input looks like 'Брянская область'!$A$28 or Москва!$A$131:$A$143

      string second;
      string[] prolog_value = source.Split('!');

      if (2 != prolog_value.Length) 
      {
        MessageBox.Show("Failed with part of this chart: " + source, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return source; 
      }

      string[] values = prolog_value[1].Split(':');
      string prolog = prolog_value[0];
      string first = values[0];

      if (2 == values.Length)
      {
        second = values[1];
      }
      else
      {
        second = first;
      }

      string value = first + ':' + second;

      Regex numberRegex = new Regex(@"\d+");
      MatchCollection numberMatches = numberRegex.Matches(value);
      if(2 != numberMatches.Count)
      {
        MessageBox.Show("Failed with part of this chart: " + source, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return source;
      }

      if (!int.TryParse(numberMatches[0].Value, out int number1))
      {
        MessageBox.Show("Failed with part of this chart: " + source, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return source;
      }

      if (!int.TryParse(numberMatches[1].Value, out int number2))
      {
        MessageBox.Show("Failed with part of this chart: " + source, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return source;
      }

      number1++;
      number2++;

      if (number2 - number1 < 4)
      {
        number1--;
      }

      string result = prolog + '!' + value.Substring(0, numberMatches[0].Index);
      result += number1.ToString();
      result += value.Substring(
        numberMatches[0].Index + numberMatches[0].Length,
        numberMatches[1].Index - (numberMatches[0].Index + numberMatches[0].Length));
      result += number2.ToString();

      return result;
    }
    
    private string UpdateChartFormula(string source)
    {
      // update chart to move it down or prolongate it on one cell if it is shorter

      source = source.Trim();
      string[] source_parts = source.Split(',');

      if (4 != source_parts.Length)
      {
        MessageBox.Show("Strange chart: " + source, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return source;
      }

      Regex orderRegex = new Regex(@"\,\d+\)");
      Match orderMatch = orderRegex.Match(source);
      if (!orderMatch.Success) return source;
      if (!int.TryParse(orderMatch.Value.Substring(1, orderMatch.Value.Length - 2), out int order))
      {
        MessageBox.Show("Failed with this chart: " + source, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return source;
      }

      string result = source_parts[0] + ',';

      if (1 != order)
      {
        result += source_parts[1];
      }
      else
      {
        result += UpdateChartFormulaPart(source_parts[1]);
      }

      result += ',' + UpdateChartFormulaPart(source_parts[2]) + ',' + source_parts[3];

      return result;
    }

    private void UpdateCellValue(Range sourceRowsRange, Range sourceColsRange, int rowNumber, int colNumber, Range targetCell, int targetRightOffset)
    {
      // take value from source cell and put to target

      Range sourceIntersectionCell = Globals.ThisAddIn.Application.Cells[((Range)sourceRowsRange.Item[rowNumber]).Row, sourceColsRange.Item[colNumber].Column];
      dynamic sourceValue = sourceIntersectionCell.Value2;
      var targetValueCell = targetCell.Offset[0, targetRightOffset]; // move right
      targetValueCell.Value2 = sourceValue;
      targetValueCell.NumberFormat = sourceIntersectionCell.NumberFormat;

#if DEBUG
      Debug.WriteLine("Set new value: " + (object)targetValueCell.Value2 + " to cell: " + targetValueCell.Row + " - " + targetValueCell.Column);
#endif
    }

    private void RunProcess(Range sheetNamesRange, Range itemNamesRange, int jumpAmount, string lookupType, string lookup, string suffix)
    {
      // main logic goes here

      int progress = 0;

      foreach (Range sheetName in sheetNamesRange)
      {
        // update progress
        ProgressBar.Label = "Progress: " + progress.ToString() + " / " + sheetNamesRange.Count;

        // current sheet we are looking for
        string sheetNameStr = sheetName.Value2;

#if DEBUG
        Debug.WriteLine("Target selected: " + sheetNameStr);
#endif
        // now we found the sheet we are looking for
        Worksheet targetSheet = GetSheetByName(sheetNameStr);

        if (null == targetSheet)
        {
          MessageBox.Show("Failed to find sheet with name = " + sheetNameStr, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
          continue;
        }

        // iterating over the values names col
        for (var i = 1; i <= itemNamesRange.Count; i += (jumpAmount+1))
        {
          // this is the name that we are looking for on the sheet
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

          // the value that we are inserting is under the cell with value name
          // so we are looking for the last filled cell, this is col with dates
          var lastCell = FindNextCellByType(lookupType, lookup, lookingTextCell.Offset[1, 0]);

          if (null == lastCell)
          {
            MessageBox.Show("Failed to find last avaliable cell for " + lookingText + " on sheet " + sheetNameStr, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            continue;
          }
          else if (lookupType == "Exact" && null == lastCell.Value2)
          {
            MessageBox.Show("This value: " + lookup + " was not found, on sheet " + sheetNameStr, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            continue;
          }

#if DEBUG
          Debug.WriteLine("Found last cell: " + lastCell.Row + " - " + lastCell.Column + " with text: " + (string)lastCell.Value2);
#endif

          // remove the value 'П' from this last date
          lastCell.Value2 = ((string)lastCell.Value2).Trim(new char[] { 'П' });

#if DEBUG
          Debug.WriteLine("Now value may be updated due to remove char: " + (string)lastCell.Value2);
#endif

          Range targetCell;

          if (lookupType == "Append")
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

          // chart shift
          if (lookupType == "Append")
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
      Debug.WriteLine("LookupBox value: " + LookupBox.Text);
      Debug.WriteLine("LookupTypeBox value: " + LookupTypeBox.Text);
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

      if (!IsValidLookup(LookupTypeBox.Text, LookupBox.Text))
      {
        MessageBox.Show("Lookup value are wrong!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
      }

#if DEBUG
      Debug.WriteLine("Input parameters are OK!");
#endif

      string type = FixUpType(LookupTypeBox.Text, LookupBox.Text);

      RunBtn.Visible = false;
      ProgressBar.Visible = true;

      RunProcess(Globals.ThisAddIn.Application.get_Range(SheetsRangeBox.Text), 
        Globals.ThisAddIn.Application.get_Range(ItemsRangeBox.Text), 
        jumpAmountValue,
        type,
        LookupBox.Text,
        DateSuffixBox.Text
      );

      RunBtn.Visible = true;
      ProgressBar.Visible = false;
    }
  }
}

