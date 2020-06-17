using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Patterns;
using FlaUI.UIA3;
using FlaUI.UIA3.Patterns;
using Gma.System.MouseKeyHook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Interop.UIAutomationClient;
using Interop.UIAutomationCore;

namespace CheckOtherApps
{

    public partial class Form1 : Form
    {
        private IKeyboardMouseEvents m_GlobalHook;

        private Rectangle rec;

        private Window wd;
        public Form1()
        {
            InitializeComponent();
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            InitRec();
            Subscribe();
            richTextBox1.Text += $"\n Binding done. ";
        }

        public void Subscribe()
        {
            // Note: for the application hook, use the Hook.AppEvents() instead
            //m_GlobalHook = Hook.AppEvents(); // GlobalEvents();
            m_GlobalHook = Hook.GlobalEvents();

            m_GlobalHook.MouseDownExt += GlobalHookMouseDownExt;
            //m_GlobalHook.KeyPress += GlobalHookKeyPress;
        }

        private void GlobalHookKeyPress(object sender, KeyPressEventArgs e)
        {
            richTextBox1.Text += $"\nSender: \t{sender}";
            richTextBox1.Text += $"\nKeyPress: \t{e.KeyChar}";
        }

        private void GlobalHookMouseDownExt(object sender, MouseEventExtArgs e)
        {
            //richTextBox1.Text += $"\nSender: \t{sender}";
            //richTextBox1.Text += $"\nMouseDown: \t{e.Button}; \t System Timestamp: \t{e.Timestamp}";
            //richTextBox1.Text += $"\n x: \t{e.X}; \t y: \t{e.Y}";

            if (rec.Contains(new Point(e.X, e.Y)))
            {
                var seltext = GetSelectedText().Result;

                richTextBox1.Text += $"\n You clicked the \"do it\" button!!!!";
                richTextBox1.Text += $"\n Value selected: {seltext} ";

                if(seltext.Contains("john"))
                {
                    richTextBox1.Text += $"\n john selected";
                }
            }

            // uncommenting the following line will suppress the middle mouse button click
            // if (e.Buttons == MouseButtons.Middle) { e.Handled = true; }
        }

        public void Unsubscribe()
        {
            m_GlobalHook.MouseDownExt -= GlobalHookMouseDownExt;
            m_GlobalHook.KeyPress -= GlobalHookKeyPress;

            //It is recommened to dispose it
            m_GlobalHook.Dispose();
        }

        public void InitRec()
        {
            Process[] processes = Process.GetProcessesByName("FormsToControl");

            foreach (Process p in processes)
            {
                var app = FlaUI.Core.Application.Attach(p);

                using (var automation = new UIA3Automation())
                {
                    wd = app.GetMainWindow(automation);
                    var children = wd.FindAllChildren();
                    ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());

                    foreach (var child in children)
                    {
                        if (child.Name == "do it")
                        {
                            //richTextBox1.Text += $"\nctl loc: {child.BoundingRectangle.Location}";
                            //richTextBox1.Text += $"\nctl right: {child.BoundingRectangle.Right}";
                            //richTextBox1.Text += $"\nctl top: {child.BoundingRectangle.Top}";
                            //richTextBox1.Text += $"\nctl bottom: {child.BoundingRectangle.Bottom}";

                            rec.Location = child.BoundingRectangle.Location;
                            rec.Width = child.BoundingRectangle.Right - child.BoundingRectangle.Location.X;
                            rec.Height = child.BoundingRectangle.Bottom - child.BoundingRectangle.Location.Y;
                        }
                    }
                }
            }
        }
        public async Task<string> GetSelectedText()
        {
            var children = wd.FindAllChildren();
            ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());

            foreach (var child in children)
            {
                if (child.Name == "DataGridView")
                {
                    var pattern = child.Patterns.Grid;
                    var rowCount = pattern.Pattern.RowCount;
                    var colCount = pattern.Pattern.ColumnCount;

                    var rows = child.FindAllChildren();

                    foreach (var r in rows)
                    {
                        //richTextBox1.Text += $"\nr: {r.Name?.ToString()}";

                        var rsel = r.Patterns.LegacyIAccessible.PatternOrDefault.State;
                        if (rsel.ValueOrDefault.ToString().IndexOf("STATE_SYSTEM_SELECTED") > -1)
                        {
                            //richTextBox1.Text += $"\nSTATE_SYSTEM_SELECTED: TRUE";
                            return r.Patterns.LegacyIAccessible.PatternOrDefault.Value;
                        }

                        var rcell = r.AsGridCell().FindAllChildren();

                        var cells = rcell.Select(x => new DataGridViewCell(x.FrameworkAutomationElement)).ToArray();
                        foreach (var c in cells)
                        {
                            if (c.Value != null)
                            {
                                //richTextBox1.Text += $"\nrcell-txt: {c.Value}";
                            }

                            var patds = c.IsPatternSupported(LegacyIAccessiblePattern.Pattern);

                            if (patds)
                            {
                                var lp = c.Patterns.LegacyIAccessible.PatternOrDefault.State;
                                if (lp.ValueOrDefault.ToString().IndexOf("STATE_SYSTEM_SELECTED") > -1)
                                {
                                    //richTextBox1.Text += $"\nSTATE_SYSTEM_SELECTED: TRUE";
                                    return c.Value;
                                }
                            }
                        }
                    }
                }
            }

            return string.Empty;
        }

        public void ForInfoOnly ()
        {
            Process[] processes = Process.GetProcessesByName("FormsToControl");

            //var app = FlaUI.Core.Application.Launch("calc.exe");
            //using (var automation = new UIA3Automation())
            //{
            //    var window = app.GetMainWindow(automation);
            //    Console.WriteLine(window.Title);

            //}               


            foreach (Process p in processes)
            {
                var app = FlaUI.Core.Application.Attach(p);

                using (var automation = new UIA3Automation())
                {
                    var window = app.GetMainWindow(automation);
                    //richTextBox1.Text += $"\nwindow title: {window.Title}";
                    var children = window.FindAllChildren();

                    ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());


                    //var bt = window.FindFirstDescendant(cf.ByName("button2")).AsButton();
                    //bt.RegisterPropertyChangedEvent(FlaUI.Core.Definitions.TreeScope.Element, (a, b, c) =>
                    //{
                    //    richTextBox1.Text += $"\na: {a.ToString()}";
                    //    richTextBox1.Text += $"\nb: {b.ToString()}";
                    //    richTextBox1.Text += $"\nc: {c.ToString()}";
                    //});

                    foreach (var child in children)
                    {
                        richTextBox1.Text += $"\nName: {child.Name?.ToString()}";
                        richTextBox1.Text += $"\nControlType: {child.ControlType.ToString()}";

                        if (child.Name == "DataGridView")
                        {
                            var pattern = child.Patterns.Grid;
                            var rowCount = pattern.Pattern.RowCount;
                            var colCount = pattern.Pattern.ColumnCount;

                            //var values = new List<object>();
                            richTextBox1.Text += $"\nrowCount: {rowCount}";
                            richTextBox1.Text += $"\ncolCount: {colCount}";

                            //for (int i = 1; i < rowCount; i++)
                            //{
                            //    //var item = pattern.Pattern.GetItem(i, 1);
                            //    //var item1 = pattern.Pattern.GetItem(i, 1);
                            //    //var item2 = pattern.Pattern.GetItem(i, 2);

                            //    //richTextBox1.Text += $"\rcell-txt: {item1.Patterns.Value.Pattern.Value}";
                            //}

                            var rows = child.FindAllChildren();

                            foreach (var r in rows)
                            {
                                richTextBox1.Text += $"\nr: {r.Name?.ToString()}";

                                var rcell = r.AsGridCell().FindAllChildren();

                                var cells = rcell.Select(x => new DataGridViewCell(x.FrameworkAutomationElement)).ToArray();
                                foreach (var c in cells)
                                {
                                    richTextBox1.Text += $"\nrcell-c: {c.Name?.ToString()}";

                                    if (c.Value != null)
                                    {
                                        //TODO: get select value from grid
                                        richTextBox1.Text += $"\nrcell-txt: {c.Value}";
                                    }

                                    //var ptt = ((LegacyIAccessiblePattern)child.get(LegacyIAccessiblePattern.Pattern));
                                    //var state = pattern.GetIAccessible().accState;


                                    richTextBox1.Text += $"\nrcell-ItemStatus: {c.Properties.ItemStatus}";

                                    ////if (child.AsGridCell().Patterns.SelectionItem != null)
                                    ////{
                                    ////    if (child.AsGridCell().Patterns.SelectionItem.PatternOrDefault != null)
                                    ////    {
                                    ////        richTextBox1.Text += $"\rcell-IsSelected: {child.AsGridCell().Patterns.SelectionItem.PatternOrDefault.IsSelected}";
                                    ////    }
                                    ////}

                                    var patds = c.IsPatternSupported(LegacyIAccessiblePattern.Pattern);

                                    richTextBox1.Text += $"\nLegacyIAccessiblePattern: {patds}";

                                    if (patds)
                                    {
                                        var lp = c.Patterns.LegacyIAccessible.PatternOrDefault.State;
                                        if (lp.ValueOrDefault.ToString().IndexOf("STATE_SYSTEM_SELECTED") > -1)
                                        {
                                            richTextBox1.Text += $"\nSTATE_SYSTEM_SELECTED: TRUE";
                                        }
                                        //richTextBox1.Text += $"\nstate: {lp.ValueOrDefault}";
                                    }
                                    //var pats = c.GetSupportedPatterns();
                                    //foreach (var pt in pats)
                                    //{
                                    //    var st = pt.
                                    //    //if (pt is ILegacyIAccessiblePattern)
                                    //    {
                                    //        richTextBox1.Text += $"\rcell-ppp: {pt.AvailabilityProperty}";
                                    //    }

                                    //}
                                }
                                //richTextBox1.Text += $"\rtype: {r.ControlType.ToString()}";
                            }
                        }


                        if (child.Name == "do it")
                        {
                            //richTextBox1.Text += $"\nName: {child.Name?.ToString()}";
                            //richTextBox1.Text += $"\nControlType: {child.ControlType.ToString()}";
                            //richTextBox1.Text += $"\nctl hndl: {child.Properties.NativeWindowHandle}";
                            //richTextBox1.Text += $"\nctl x: {child.GetClickablePoint().X}";
                            //richTextBox1.Text += $"\nctl y: {child.GetClickablePoint().Y}";
                            richTextBox1.Text += $"\nctl loc: {child.BoundingRectangle.Location}";
                            richTextBox1.Text += $"\nctl right: {child.BoundingRectangle.Right}";
                            richTextBox1.Text += $"\nctl top: {child.BoundingRectangle.Top}";
                            richTextBox1.Text += $"\nctl bottom: {child.BoundingRectangle.Bottom}";

                            rec.Location = child.BoundingRectangle.Location;
                            rec.Width = child.BoundingRectangle.Right - child.BoundingRectangle.Location.X;
                            rec.Height = child.BoundingRectangle.Bottom - child.BoundingRectangle.Location.Y;
                        }
                    }

                }

                //this.CurWindow = new WindowInfo(p.MainWindowHandle);
            }
        }
    }
    
    public class DataGridViewCell : AutomationElement
    {
        public DataGridViewCell(FrameworkAutomationElementBase basicAutomationElement) : base(basicAutomationElement)
        {
        }

        protected IValuePattern ValuePattern
        {
            get
            {
                try
                {
                    return Patterns?.Value?.Pattern;
                }
                catch { }
                return null;
            }
        }

        public string Value
        {
            get
            {
                try
                {
                    return ValuePattern.Value.Value.ToString();
                }
                catch { }
                return null;
            }
            set => ValuePattern.SetValue(value);
        }
    }
}
