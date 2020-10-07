﻿namespace LucyAutoExAddIn
{
  partial class LucyAutoRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
  {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    public LucyAutoRibbon()
        : base(Globals.Factory.GetRibbonFactory())
    {
      InitializeComponent();
    }

    /// <summary> 
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
      if (disposing && (components != null))
      {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Component Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
      this.AutomatizationRibbon = this.Factory.CreateRibbonTab();
      this.Automatization = this.Factory.CreateRibbonGroup();
      this.box1 = this.Factory.CreateRibbonBox();
      this.SheetsRangeBox = this.Factory.CreateRibbonEditBox();
      this.ItemsRangeBox = this.Factory.CreateRibbonEditBox();
      this.JumpAmountBox = this.Factory.CreateRibbonEditBox();
      this.separator1 = this.Factory.CreateRibbonSeparator();
      this.box2 = this.Factory.CreateRibbonBox();
      this.AppendBox = this.Factory.CreateRibbonCheckBox();
      this.DateSuffixBox = this.Factory.CreateRibbonEditBox();
      this.RunBtn = this.Factory.CreateRibbonButton();
      this.ProgressBar = this.Factory.CreateRibbonLabel();
      this.AutomatizationRibbon.SuspendLayout();
      this.Automatization.SuspendLayout();
      this.box1.SuspendLayout();
      this.box2.SuspendLayout();
      this.SuspendLayout();
      // 
      // AutomatizationRibbon
      // 
      this.AutomatizationRibbon.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
      this.AutomatizationRibbon.ControlId.OfficeId = "TabData";
      this.AutomatizationRibbon.Groups.Add(this.Automatization);
      this.AutomatizationRibbon.Label = "TabData";
      this.AutomatizationRibbon.Name = "AutomatizationRibbon";
      // 
      // Automatization
      // 
      this.Automatization.Items.Add(this.box1);
      this.Automatization.Items.Add(this.separator1);
      this.Automatization.Items.Add(this.box2);
      this.Automatization.Label = "Lucy`s automatization";
      this.Automatization.Name = "Automatization";
      // 
      // box1
      // 
      this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
      this.box1.Items.Add(this.SheetsRangeBox);
      this.box1.Items.Add(this.ItemsRangeBox);
      this.box1.Items.Add(this.JumpAmountBox);
      this.box1.Name = "box1";
      // 
      // SheetsRangeBox
      // 
      this.SheetsRangeBox.Label = "Sheets range:";
      this.SheetsRangeBox.Name = "SheetsRangeBox";
      this.SheetsRangeBox.ScreenTip = "Choose range of values to search for as sheets in this document";
      this.SheetsRangeBox.Text = null;
      // 
      // ItemsRangeBox
      // 
      this.ItemsRangeBox.Label = "Cells range:";
      this.ItemsRangeBox.Name = "ItemsRangeBox";
      this.ItemsRangeBox.ScreenTip = "Choose colomn range to search on found sheets";
      this.ItemsRangeBox.Text = null;
      // 
      // JumpAmountBox
      // 
      this.JumpAmountBox.Label = "Jump amount:";
      this.JumpAmountBox.Name = "JumpAmountBox";
      this.JumpAmountBox.ScreenTip = "While iterating on cells range we have to skip several cells due to answers";
      this.JumpAmountBox.Text = null;
      // 
      // separator1
      // 
      this.separator1.Name = "separator1";
      // 
      // box2
      // 
      this.box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
      this.box2.Items.Add(this.AppendBox);
      this.box2.Items.Add(this.DateSuffixBox);
      this.box2.Items.Add(this.RunBtn);
      this.box2.Items.Add(this.ProgressBar);
      this.box2.Name = "box2";
      // 
      // AppendBox
      // 
      this.AppendBox.Label = "Append";
      this.AppendBox.Name = "AppendBox";
      this.AppendBox.ScreenTip = "If selected then create new cell on found";
      // 
      // DateSuffixBox
      // 
      this.DateSuffixBox.Label = "Date suffix:";
      this.DateSuffixBox.Name = "DateSuffixBox";
      this.DateSuffixBox.Text = null;
      // 
      // RunBtn
      // 
      this.RunBtn.Label = "Do this magic";
      this.RunBtn.Name = "RunBtn";
      this.RunBtn.ScreenTip = "Vzzhhuuuuh...";
      this.RunBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Run_Click);
      // 
      // ProgressBar
      // 
      this.ProgressBar.Label = "Progress: 0/0";
      this.ProgressBar.Name = "ProgressBar";
      this.ProgressBar.Visible = false;
      // 
      // LucyAutoRibbon
      // 
      this.Name = "LucyAutoRibbon";
      this.RibbonType = "Microsoft.Excel.Workbook";
      this.Tabs.Add(this.AutomatizationRibbon);
      this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.LucyAutoRibbon_Load);
      this.AutomatizationRibbon.ResumeLayout(false);
      this.AutomatizationRibbon.PerformLayout();
      this.Automatization.ResumeLayout(false);
      this.Automatization.PerformLayout();
      this.box1.ResumeLayout(false);
      this.box1.PerformLayout();
      this.box2.ResumeLayout(false);
      this.box2.PerformLayout();
      this.ResumeLayout(false);

    }

    #endregion

    internal Microsoft.Office.Tools.Ribbon.RibbonTab AutomatizationRibbon;
    internal Microsoft.Office.Tools.Ribbon.RibbonGroup Automatization;
    internal Microsoft.Office.Tools.Ribbon.RibbonButton RunBtn;
    internal Microsoft.Office.Tools.Ribbon.RibbonEditBox SheetsRangeBox;
    internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ItemsRangeBox;
    internal Microsoft.Office.Tools.Ribbon.RibbonEditBox JumpAmountBox;
    internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox AppendBox;
    internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    internal Microsoft.Office.Tools.Ribbon.RibbonEditBox DateSuffixBox;
    internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
    internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
    internal Microsoft.Office.Tools.Ribbon.RibbonLabel ProgressBar;
  }

  partial class ThisRibbonCollection
  {
    internal LucyAutoRibbon LucyAutoRibbon
    {
      get { return this.GetRibbon<LucyAutoRibbon>(); }
    }
  }
}
