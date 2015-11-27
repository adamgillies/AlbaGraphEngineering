using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AlbaLibrary.ToolBox;
using AlbaLibrary.IO;
using AlbaLibrary.Chart;
using System.IO;
using AlbaLibrary.Extension;
using AlbaLibrary.Interpolation;
using AlbaLibrary.XL;
using System.Threading.Tasks;
using System.Threading;
using AlbaLibrary.Clone;

namespace AlbaGraphEngineering
{
    public partial class GraphForm : Form
    {
        private enum SaveVariable { GB = 0, ZTheta = 1};
        private bool Busy = false;
        private int lineWidth = 2;

        private Object[] AutoGraphs;

        private void EnableSelection(RadioButton sender, ComboBox comboBox)
        {
            if (((RadioButton)sender).Checked)
            {
                comboBox.Visible = true;
                comboBox.Enabled = true;
                comboBox.TabStop = true;
                if (comboBox.Items.Count > 0)
                    comboBox.SelectedIndex = 0;
            }
            else
            {
                comboBox.TabStop = false;
                comboBox.Visible = false;
                comboBox.Enabled = false;
            }
        }

        public GraphForm()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Load the initial files from the input directory string
        /// </summary>
        /// <param name="directoryString">the string of the directory</param>
        public void LoadFiles(string directoryString, bool updateIndex)
        {
            try
            {
                // load the directory files and sort
                var files = new SortedDirectoryFiles(directoryString, "*.xls");
                // assign the files to the correct allocation in the drop down
                myIO.AssignFolders(files, updateIndex, cmb_folder, lstbx_data);

                // assign number of files
                txt_filecount.Text = lstbx_data.Items.Count.ToString("###0");
                nup_end.Value = lstbx_data.Items.Count;
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Load Files Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
       
       


       
        /// <summary>
        /// plot the tranraw or summary data obtained from the directory
        /// </summary>
        /// <param name="directoryString">top level directory</param>
        /// <param name="fileIndex">file to plot</param>
        private void PlotAppropriateFile(string directoryString, ListBox myListBox, bool summaryRead, bool nameSort)
        {
            try
            {
                var fileIndex = myListBox.SelectedIndex;

                if (fileIndex > -1)
                {
                    if (!chbx_hold.Checked || AutoGraphs == null)
                    {
                        zgc_graph.GraphPane.CurveList.Clear();
                        AutoGraphs = new ZedGraph.GraphPane[7] { zgc_graph.GraphPane.Clone(), 
                    zgc_graph.GraphPane.Clone(), zgc_graph.GraphPane.Clone(), zgc_graph.GraphPane.Clone(), 
                    zgc_graph.GraphPane.Clone(), zgc_graph.GraphPane.Clone(), zgc_graph.GraphPane.Clone() };
                    }

                    // load the files
                    var selectedFile = myIO.CreateSelectedFiles(new SortedDirectoryFiles(directoryString, "*.xls"), cmb_folder, summaryRead, nameSort)
                        .ToArray()[fileIndex];

                    if (rb_transraw.Checked)
                        PlotTranRawFile(selectedFile, fileIndex, cmbx_plotSelect.SelectedIndex, AutoGraphs.Cast<ZedGraph.GraphPane>().ToArray());
                    else
                        PlotSummaryFile(selectedFile);
                }
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Plot Appropriate Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// plot the tran raw file data
        /// </summary>
        /// <param name="selectedFile">file to plot</param>
        private void PlotTranRawFile(FileInfo selectedFile, int fileIndex, int variableIndex, ZedGraph.GraphPane[] Graphs)
        {
            try
            {
                var series = Path.GetFileNameWithoutExtension(selectedFile.Name);
                var curveList = Graphs[variableIndex].CurveList;

                if (chbx_hold.Checked == true && curveList.Select(i => i).Where(j => j.Label.Text == series).ToList().Count > 0)
                {
                    var curveItem = curveList.First(i => i.Label.Text == series);
                    var tag = curveItem.Tag.ToString();
                    var tagList = curveList.Select(i => i.Tag.ToString()).ToList();
                    curveList.RemoveAt(tagList.IndexOf(tag));
                }
                else
                {
                    // create the data from the file
                    var impeData = new AlbaLibrary.Data.Impedance(selectedFile, '\t');

                    if (rb_on.Checked)
                    {
                        var InterpolationType = InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex);

                        if (rb_pnts.Checked)
                            impeData.Frequency.Data = impeData.Frequency.Data.ReturnInterpolationXPnts(Convert.ToInt16(txt_nopnts.Text)).ToArray();
                        else
                            impeData.Frequency.Data = impeData.Frequency.Data.ReturnInterpolationXStep(Convert.ToDouble(txt_freqstep.Text)).ToArray();

                        impeData = impeData.InterpolateData(impeData.Frequency.Data, InterpolationType);
                    }

                    var frequency = impeData.Frequency.Data;
                    var options = new ParallelOptions();

                    options.MaxDegreeOfParallelism = AlbaLibrary.Extension.Extension.DetectCores() - 1;

                    #region
                    
                    Parallel.ForEach<ZedGraph.GraphPane>(Graphs, options, pane =>
                    {
                        int paneIndex = Graphs.ToList().IndexOf(pane);
                        var outputData = PlotSelecter(pane, paneIndex, impeData).ToArray();

                        pane.AddSeries(frequency, outputData, series, ZedGraph.SymbolType.None, lineWidth, fileIndex);

                        if (variableIndex == paneIndex && pane.CurveList.Count > 0)
                        {
                            CreateStats(outputData);
                            SetAutoGraph(variableIndex);
                        }
                    });

                    #endregion
                }
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Plot Transraw Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private IEnumerable<double> PlotSelecter(ZedGraph.GraphPane pane, int paneIndex, AlbaLibrary.Data.Impedance impeData)
        {          
            var unit = impeData.Frequency.Unit;

            pane.YAxis.Type = ZedGraph.AxisType.Linear;
            pane.XAxis.Title.Text = "Frequency (" + unit + ")";

            switch (paneIndex)
            {
                default:
                    pane.YAxis.Title.Text = "Conductance (" + impeData.Conductance.Unit + ")";
                   return impeData.Conductance.Data;
                    
                case 1:
                    pane.YAxis.Title.Text = "Susceptance (" + impeData.Susceptance.Unit + ")";
                    return impeData.Susceptance.Data;
                    
                case 2:
                    pane.YAxis.Title.Text = "Capacitance (" + impeData.Capacitance.Unit + ")";
                    return impeData.Capacitance.Data;
                  
                case 3:
                    pane.YAxis.Title.Text = "Resistance (" + impeData.Resistance.Unit + ")";
                    return impeData.Resistance.Data;

                case 4:
                    pane.YAxis.Title.Text = "Reactance (" + impeData.Reactance.Unit + ")";
                    return impeData.Reactance.Data;

                case 5:
                    pane.YAxis.Title.Text = "Impedance Magnitude (" + impeData.ImpedanceMagnitude.Unit + ")";                  
                    pane.YAxis.Type = ZedGraph.AxisType.Log;
                    return impeData.ImpedanceMagnitude.Data;
                case 6:
                    pane.YAxis.Title.Text = "Impedance Phase (" + impeData.ImpedancePhase.Unit + ")";
                    return impeData.ImpedancePhase.Data;                  
            }     
        }

        private void PlotAutoTranRawFile(FileInfo selectedFile, ZedGraph.GraphPane[] Graphs, ParallelOptions options, string series, int tag)
        {
            try
            {
                // create the data from the file
                var impeData = new AlbaLibrary.Data.Impedance(selectedFile, '\t');

                if (rb_on.Checked)
                {
                    var InterpolationType = InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex);

                        if (rb_pnts.Checked)
                            impeData.Frequency.Data = impeData.Frequency.Data.ReturnInterpolationXPnts(Convert.ToInt16(txt_nopnts.Text)).ToArray();
                        else
                            impeData.Frequency.Data = impeData.Frequency.Data.ReturnInterpolationXStep(Convert.ToDouble(txt_freqstep.Text)).ToArray();

                    impeData = impeData.InterpolateData(impeData.Frequency.Data, InterpolationType);
                }

                // create the data from the file
                #region
                double[] frequency = impeData.Frequency.Data;
                List<double[]> outputData = new List<double[]>()
            {
                impeData.Conductance.Data,
                impeData.Susceptance.Data,
                impeData.Capacitance.Data,
                impeData.Resistance.Data,
                impeData.Reactance.Data,
                impeData.ImpedanceMagnitude.Data,
                impeData.ImpedancePhase.Data,
            };
                #endregion

                // check for interpolation and then plot the relevant data
                #region

                Parallel.ForEach(Graphs, options, Graph =>
                {
                    Graph.AddAutoSeries(frequency, outputData[Array.IndexOf(Graphs, Graph)], series, ZedGraph.SymbolType.None, lineWidth, tag);
                }
                );
                #endregion
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Auto Transraw Plot", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public int ComboBoxIndex(ComboBox comboBox)
        {
            int count = 0;
            comboBox.Invoke(new MethodInvoker(delegate
            {
                count = comboBox.SelectedIndex;
            }));
            return count;
        }

       

        /// <summary>
        /// plot the summary file data 
        /// </summary>
        /// <param name="selectedFile">file to plot</param>
        private void PlotSummaryFile(FileInfo selectedFile)
        {
            try
            {
                // create the data from the file
                var series = Path.GetFileNameWithoutExtension(selectedFile.Name) + " " + cmbx_summarySelect.SelectedItem.ToString();
                var summData = new AlbaLibrary.Data.SummaryData(selectedFile, '\t', '_');
                IEnumerable<double> outputData;

                if (!chbx_hold.Checked)
                    zgc_graph.GraphPane.CurveList.Clear();

                if (chbx_hold.Checked == true && zgc_graph.GraphPane.CurveList.Select(i => i).Where(j => j.Label.Text == series).ToList().Count > 0)
                {
                    var curveItem = zgc_graph.GraphPane.CurveList.First(i => i.Label.Text == series);
                    var tag = curveItem.Tag.ToString();
                    var tagList = zgc_graph.GraphPane.CurveList.Select(i => i.Tag.ToString()).ToList();
                    zgc_graph.GraphPane.CurveList.RemoveAt(tagList.IndexOf(tag));
                }
                else
                {
                    #region
                    switch (cmbx_summarySelect.SelectedIndex)
                    {
                        case 0:

                            outputData = summData.Fr.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "Fr (" + summData.Fr.Unit + ")";
                            break;
                        case 1:

                            outputData = summData.Fa.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "Fa (" + summData.Fa.Unit + ")";
                            break;
                        case 2:

                            outputData = summData.Fe.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "Fe (" + summData.Fe.Unit + ")";
                            break;
                        case 3:

                            outputData = summData.Fm.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "Fm (" + summData.Fm.Unit + ")";
                            break;
                        case 4:

                            outputData = summData.GFr.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "GFr (" + summData.GFr.Unit + ")";
                            break;
                        case 5:

                            outputData = summData.BFr.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "BFr (" + summData.BFr.Unit + ")";
                            break;
                        case 6:

                            outputData = summData.RFr.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "RFr (" + summData.RFr.Unit + ")";
                            break;
                        case 7:

                            outputData = summData.FHalfGFrLow.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "FHalfGFr Low (" + summData.FHalfGFrLow.Unit + ")";
                            break;
                        case 8:

                            outputData = summData.FHalfGFrHigh.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "FHalfGFr High (" + summData.FHalfGFrHigh.Unit + ")";
                            break;
                        case 9:

                            outputData = summData.Cp1000.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "Cp at 1 kHz (" + summData.Cp1000.Unit + ")";
                            break;
                        case 10:

                            outputData = summData.CpHF.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "Cp HF (" + summData.CpHF.Unit + ")";
                            break;
                        case 11:

                            outputData = summData.BW3dB.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "BW 3dB (" + summData.BW3dB.Unit + ")";
                            break;
                        case 12:

                            outputData = summData.FBW.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "FBW (" + summData.FBW.Unit + ")";
                            break;
                        case 13:

                            outputData = summData.Q.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "Q";
                            break;
                        case 14:

                            outputData = summData.Kt.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "Kt (" + summData.Kt.Unit + ")";
                            break;
                        case 15:

                            outputData = summData.Keff.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "Keff (" + summData.Keff.Unit + ")";
                            break;
                        case 16:

                            outputData = summData.Temperature.Data;
                            zgc_graph.GraphPane.YAxis.Title.Text = "Temperature (" + summData.Temperature.Unit + ")";
                            break;
                        default:
                            outputData = summData.ElementNumber.Data;
                            break;
                    }
                    #endregion

                    zgc_graph.GraphPane.AddSeries(summData.ElementNumber.Data, outputData.ToArray(),
                              series, ZedGraph.SymbolType.None, lineWidth,
                              cmb_folder.SelectedIndex * 100 + cmbx_summarySelect.SelectedIndex);
                    zgc_graph.GraphPane.XAxis.Title.Text = "Element Number";
                    CreateStats(outputData.ToArray());
                }

                zgc_graph.SetupZedGraph();
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Plot Summary Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// plot at a particular frequency data value for each file 
        /// </summary>
        /// <param name="directoryString">directory of files</param>
        private void PlotAt(string directoryString, int variableIndex)
        {
            try
            {
                if (!chbx_hold.Checked)
                    zgc_graph.GraphPane.CurveList.Clear();

                // load the files
                var files = new SortedDirectoryFiles(directoryString, "*.xls");
                // 
                var selectedFiles = myIO.CreateSelectedFiles(files, cmb_folder, rb_summary.Checked, rb_nam.Checked).ToArray();
                // assign the file to plot on the graph

                ExtractPlotAtData(selectedFiles, variableIndex);

                zgc_graph.SetupZedGraph();
                var xMax = zgc_graph.GraphPane.XAxis.Scale.Max + 1;

                if (xMax % 2 == 0)
                    zgc_graph.GraphPane.XAxis.Scale.Max = xMax;
                else
                    zgc_graph.GraphPane.XAxis.Scale.Max = xMax + 1;

                zgc_graph.GraphPane.AxisChange();
                zgc_graph.Refresh();
                zgc_graph.Update();
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Plot At Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

       

        /// <summary>
        /// extract specific data from the data set for each file in 
        /// the file array
        /// </summary>
        /// <param name="files">array of files to extract data from</param>
        private void ExtractPlotAtData(FileInfo[] files, int variableIndex)
        {
            try
            {
                // variables
                #region
                var elementNumber = new List<double>();
                var outputData = new List<double>();

                var yAxisLabel = new string[9]{
                "Frequency (Hz)",
                "Conductance (S)",
                "Susceptance (S)",
                "Capacitance (F)",
                "Resistance (Ω)",
                "Reactance (Ω)",
                "ImpedanceMagnitude (Ω)",
                "ImpedancePhase (°)",
                "Q-Factor",
            };

                var filteredFiles = myIO.FilterFiles(files, (int)nup_start.Value, (int)nup_end.Value, (int)nup_step.Value);
                #endregion

                // loop through each file extracting relevant data
                #region
                foreach (var file in filteredFiles)
                {
                    // create an index tracker
                    var index = Array.IndexOf(files, file);
                    // assign the element
                    elementNumber.Add(index + 1d);
                    // create the data from the file
                    var impeData = new AlbaLibrary.Data.Impedance(file, '\t');
                    var frequency = impeData.Frequency.Data;
                    var conductance = impeData.Conductance.Data;

                    if (rb_on.Checked)
                    {
                        var InterpolationType = InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex);

                        if (rb_pnts.Checked)
                            frequency = impeData.Frequency.Data.ReturnInterpolationXPnts(Convert.ToInt16(txt_nopnts.Text)).ToArray();
                        else
                            frequency = impeData.Frequency.Data.ReturnInterpolationXStep(Convert.ToDouble(txt_freqstep.Text)).ToArray();

                        conductance = InterpolationAlgorithms.Interpolate(impeData.Frequency.Data, conductance, frequency, InterpolationType).ToArray();
                    }

                    var searchValue = frequency[conductance.ToList().GetClosestIndex<double>(conductance.Max())];

                    if (rb_value.Checked)
                        searchValue = txt_value.Text.GetDouble() * 1000d;

                    var dataIndex = frequency.ToList().GetClosestIndex<double>(searchValue);

                    // switch on the appropriate variables to select the data
                    #region

                    switch (variableIndex)
                    {
                        case 0:
                            if (rb_on.Checked)
                                outputData.Add(frequency[dataIndex]);
                            else
                                outputData.Add(impeData.Frequency.Data[dataIndex]);
                            break;
                        case 1:
                            if (rb_on.Checked)
                                outputData.Add(conductance[dataIndex]);
                            else
                                outputData.Add(impeData.Conductance.Data[dataIndex]);
                            break;
                        case 2:
                            if (rb_on.Checked)
                                outputData.Add(InterpolationAlgorithms.Interpolate(impeData.Frequency.Data, impeData.Susceptance.Data, frequency, InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex)).ToArray()[dataIndex]);
                            else
                                outputData.Add(impeData.Susceptance.Data[dataIndex]);
                            break;
                        case 3:
                            if (rb_on.Checked)
                                outputData.Add(InterpolationAlgorithms.Interpolate(impeData.Frequency.Data, impeData.Capacitance.Data, frequency, InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex)).ToArray()[dataIndex]);
                            else
                                outputData.Add(impeData.Capacitance.Data[dataIndex]);
                            break;
                        case 4:
                            if (rb_on.Checked)
                                outputData.Add(InterpolationAlgorithms.Interpolate(impeData.Frequency.Data, impeData.Resistance.Data, frequency, InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex)).ToArray()[dataIndex]);
                            else
                                outputData.Add(impeData.Resistance.Data[dataIndex]);
                            break;
                        case 5:
                            if (rb_on.Checked)
                                outputData.Add(InterpolationAlgorithms.Interpolate(impeData.Frequency.Data, impeData.Reactance.Data, frequency, InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex)).ToArray()[dataIndex]);
                            else
                                outputData.Add(impeData.Reactance.Data[dataIndex]);
                            break;
                        case 6:
                            if (rb_on.Checked)
                                outputData.Add(InterpolationAlgorithms.Interpolate(impeData.Frequency.Data, impeData.ImpedanceMagnitude.Data, frequency, InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex)).ToArray()[dataIndex]);
                            else
                                outputData.Add(impeData.ImpedanceMagnitude.Data[dataIndex]);
                            break;
                        case 7:
                            if (rb_on.Checked)
                                outputData.Add(InterpolationAlgorithms.Interpolate(impeData.Frequency.Data, impeData.ImpedancePhase.Data, frequency, InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex)).ToArray()[dataIndex]);
                            else
                                outputData.Add(impeData.ImpedancePhase.Data[dataIndex]);
                            break;
                        case 8:
                            outputData.Add(impeData.Q);
                            break;
                    }
                }

                zgc_graph.GraphPane.AddSeries(elementNumber.ToArray(), outputData.ToArray(), cmbx_plotSelect.SelectedItem.ToString(), ZedGraph.SymbolType.Circle, lineWidth, lstbx_data.SelectedIndex + 100);
                zgc_graph.GraphPane.XAxis.Title.Text = "Element Number";
                zgc_graph.GraphPane.YAxis.Title.Text = yAxisLabel[variableIndex];
                zgc_graph.GraphPane.Title.Text = "";
                CreateStats(outputData.ToArray());

                    #endregion
                #endregion
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Extract Plot Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Overlay all of the data from the selected directory
        /// </summary>
        /// <param name="directoryString">top level directory string</param>
        private void AutoPlot(string directoryString, int variableIndex, ZedGraph.GraphPane[] graphs)
        {
            try
            {
                zgc_graph.GraphPane.CurveList.Clear();

                // reset graph storage
                graphs = new ZedGraph.GraphPane[7] { zgc_graph.GraphPane.Clone(), 
                zgc_graph.GraphPane.Clone(), zgc_graph.GraphPane.Clone(), zgc_graph.GraphPane.Clone(), 
                zgc_graph.GraphPane.Clone(), zgc_graph.GraphPane.Clone(), zgc_graph.GraphPane.Clone() };

                // load the files
                var selectedFiles = myIO.FilterFiles(myIO.CreateSelectedFiles(new SortedDirectoryFiles(directoryString, "*.xls"), cmb_folder, rb_summary.Checked, rb_nam.Checked), 
                    (int)nup_start.Value, (int)nup_end.Value, (int)nup_step.Value).ToArray();

                this.Cursor = Cursors.WaitCursor;

                // set axis
                #region
                // axis labels
                string unit = string.Empty;
                string xAxis = "Frequency (Hz)";
                string[] yAxisLabel = new string[7]{
                "Conductance (S)",
                "Susceptance (S)",
                "Capacitance (F)",
                "Resistance (Ω)",
                "Reactance (Ω)",
                "ImpedanceMagnitude (Ω)",
                "ImpedancePhase (°)",
            };

                List<ZedGraph.YAxis> yAxis = new List<ZedGraph.YAxis>()
            {
                new ZedGraph.YAxis { Type = ZedGraph.AxisType.Linear},  
                new ZedGraph.YAxis { Type = ZedGraph.AxisType.Linear},  
                new ZedGraph.YAxis { Type = ZedGraph.AxisType.Linear},  
                new ZedGraph.YAxis { Type = ZedGraph.AxisType.Log},  
                new ZedGraph.YAxis { Type = ZedGraph.AxisType.Log},  
                new ZedGraph.YAxis { Type = ZedGraph.AxisType.Log},  
                new ZedGraph.YAxis { Type = ZedGraph.AxisType.Linear},  
            };

                #endregion

                var options = new ParallelOptions();
                options.MaxDegreeOfParallelism = AlbaLibrary.Extension.Extension.DetectCores() - 1;

                Parallel.ForEach<FileInfo>(selectedFiles, options, file =>
                {
                    var index = Array.IndexOf(selectedFiles.ToArray(), file);
                    PlotAutoTranRawFile(file, graphs, options, Path.GetFileNameWithoutExtension(file.Name), index);
                });

                foreach (var Graph in graphs)
                {
                    var index = Array.IndexOf(graphs, Graph);
                    Graph.XAxis.Title.Text = xAxis;
                    Graph.YAxis.Title.Text = yAxisLabel[index];
                    Graph.YAxis.Type = yAxis[index].Type;
                    Graph.Title.Text = "";

                    // sort the series on the tag value
                    Graph.CurveList.Sort(new AlbaLibrary.Chart.Utility.CurveItemTagComparer());
                    Graph.CurveList.Sort(new AlbaLibrary.Chart.Utility.CurveItemTagComparer());

                    AlbaLibrary.Chart.SeriesColor.AssignColor(Graph);
                }

                AutoGraphs = graphs;
                this.Cursor = Cursors.Default;
                SetAutoGraph(variableIndex);
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Auto Plot", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Find the min max average and standard dev for last added series
        /// </summary>
        /// <param name="data">data set of last added series</param>
        private void CreateStats(IEnumerable<double> data)
        {
            try
            {
                var format = "0.0#E0";
                InvokeTextBox(txt_aver, data.Average().ToString(format));
                InvokeTextBox(txt_min, data.Min().ToString(format));
                InvokeTextBox(txt_max, data.Max().ToString(format));
                // calc standard dev
                InvokeTextBox(txt_std, Math.Sqrt(data.Select(i => Math.Pow(i - data.Average(), 2)).ToList().Sum() / data.Count<double>()).ToString(format));
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Create Stats", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Save the data on the graph
        /// </summary>
        /// <param name="zgc">graph</param>
        private void SaveOnScreenData(ZedGraph.ZedGraphControl zgc)
        {
            try
            {
                var save = new SaveFileDialog();
                save.Filter = "Microsoft Excel Workbook (*.xls) | *.xls";
                save.InitialDirectory = txt_directory.Text;

                if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var Graph = zgc.GraphPane;

                    var dataList = new List<string[]>();
                    var headerList = new List<string>();
                    var titleList = new List<string>();

                    foreach (var lineItem in Graph.CurveList)
                    {
                        titleList.AddRange(new string[3] { lineItem.Label.Text, "", "" });
                        headerList.AddRange(new string[3] { Graph.XAxis.Title.Text, Graph.YAxis.Title.Text, "" });

                        var xAxis = Enumerable.Range(0, lineItem.NPts).Select(i => lineItem.Points[i].X.ToString()).ToArray();
                        var yAxis = Enumerable.Range(0, lineItem.NPts).Select(i => lineItem.Points[i].Y.ToString()).ToArray();
                        var space = Enumerable.Range(0, lineItem.NPts).Select(i => "").ToArray();

                        dataList.AddRange(new string[3][] { xAxis, yAxis, space });
                    }

                    var WriteXL = new WriteXL(titleList.ToArray(), headerList.ToArray(), dataList.ToArray(), save.FileName, "On Screen Data", AlbaLibrary.XL.WriteXL.WriteXLMode.Xls);
                }
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Save On Screen Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Save the raw data from the currently selected directory
        /// </summary>
        /// <param name="directoryString">main directory</param>
        /// <param name="saveVariable">what type of data to save GB or ZA</param>
        private void SaveAll(string directoryString, SaveVariable saveVariable, bool interp)
        {
            try
            {
                var save = new SaveFileDialog();
                save.Filter = "Microsoft Excel Workbook (*.xls) | *.xls";
                save.InitialDirectory = txt_directory.Text;

                this.Cursor = Cursors.WaitCursor;

                if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // load the files
                    var files = new SortedDirectoryFiles(directoryString, "*.xls");
                    // 
                    var selectedFiles = myIO.CreateSelectedFiles(files, cmb_folder, rb_summary.Checked, rb_nam.Checked)
                        .ToArray();

                    var dataList = new List<string[]>();
                    var headerList = new List<string>();
                    var titleList = new List<string>();

                    foreach (FileInfo file in selectedFiles)
                    {
                        titleList = new List<string>() { Path.GetFileNameWithoutExtension(file.Name), "", "" };

                        switch (saveVariable)
                        {
                            case SaveVariable.GB:
                                // create the data from the file
                                var admData = new AlbaLibrary.Data.Admittance(file, '\t');

                                headerList.AddRange( new string[3] {"Frequency ("+admData.Frequency.Unit + ")",
                                                                    "Conductance (" + admData.Conductance.Unit + ")" ,
                                                                    "Susceptance (" + admData.Susceptance.Unit + ")" ,
                                });

                                if (interp)
                                {
                                    var InterpolationType = InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex);
                                    var interpF = admData.Frequency.Data;

                                    if (rb_pnts.Checked)
                                        interpF = interpF.ReturnInterpolationXPnts(Convert.ToInt16(txt_nopnts.Text)).ToArray();
                                    else
                                        interpF = interpF.ReturnInterpolationXStep(Convert.ToDouble(txt_freqstep.Text)).ToArray();
       
                                    var interpG = InterpolateClass.InterpolateYData(admData.Frequency.Data, admData.Conductance.Data, interpF, InterpolationType);
                                    var interpB = InterpolateClass.InterpolateYData(admData.Frequency.Data, admData.Susceptance.Data, interpF, InterpolationType);

                                    dataList.Add(Array.ConvertAll<double, string>(interpG.ToArray()[0].ToArray(), Convert.ToString));
                                    dataList.Add(Array.ConvertAll<double, string>(interpG.ToArray()[1].ToArray(), Convert.ToString));
                                    dataList.Add(Array.ConvertAll<double, string>(interpB.ToArray()[1].ToArray(), Convert.ToString));
                                }
                                else
                                {
                                    dataList.Add(Array.ConvertAll<double, string>(admData.Frequency.Data, Convert.ToString));
                                    dataList.Add(Array.ConvertAll<double, string>(admData.Conductance.Data, Convert.ToString));
                                    dataList.Add(Array.ConvertAll<double, string>(admData.Susceptance.Data, Convert.ToString));
                                }
                                break;

                            case SaveVariable.ZTheta:
                                // create the data from the file
                                AlbaLibrary.Data.Impedance impeData = new AlbaLibrary.Data.Impedance(file, '\t');

                                headerList.AddRange( new string[3] { "Frequency (" +impeData.Frequency.Unit + ")",
                                                                     "Impedance Magnitude (" + impeData.ImpedanceMagnitude.Unit + ")" ,
                                                                     "Impedance Phase (" + impeData.ImpedancePhase.Unit + ")" ,
                                });
                        
                                if (interp)
                                {
                                    var InterpolationType = InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex);
                                    var interpF = impeData.Frequency.Data;

                                    if (rb_pnts.Checked)
                                        interpF = interpF.ReturnInterpolationXPnts(Convert.ToInt16(txt_nopnts.Text)).ToArray();
                                    else
                                        interpF = interpF.ReturnInterpolationXStep(Convert.ToDouble(txt_freqstep.Text)).ToArray();

                                    var interpZ = InterpolateClass.InterpolateYData(impeData.Frequency.Data, impeData.ImpedanceMagnitude.Data, interpF, InterpolationType);
                                    var interpZT = InterpolateClass.InterpolateYData(impeData.Frequency.Data, impeData.ImpedancePhase.Data, interpF, InterpolationType);

                                    dataList.Add(Array.ConvertAll<double, string>(interpZ.ToArray()[0].ToArray().ToArray(), Convert.ToString));
                                    dataList.Add(Array.ConvertAll<double, string>(interpZ.ToArray()[1].ToArray(), Convert.ToString));
                                    dataList.Add(Array.ConvertAll<double, string>(interpZT.ToArray()[1].ToArray(), Convert.ToString));
                                }
                                else
                                {
                                    dataList.Add(Array.ConvertAll<double, string>(impeData.Frequency.Data, Convert.ToString));
                                    dataList.Add(Array.ConvertAll<double, string>(impeData.ImpedanceMagnitude.Data, Convert.ToString));
                                    dataList.Add(Array.ConvertAll<double, string>(impeData.ImpedancePhase.Data, Convert.ToString));
                                }
                                break;
                        }

                    }

                    var WriteXL = new WriteXL(titleList.ToArray(), headerList.ToArray(), dataList.ToArray(), save.FileName, "Data", AlbaLibrary.XL.WriteXL.WriteXLMode.Xls);
                    this.Cursor = Cursors.Default;
                }
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Sort All Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveTransRawData(string directoryString, SaveVariable saveVariable, bool interp)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                // load the files
                var files = new SortedDirectoryFiles(directoryString, "*.xls");
                // 
                var selectedFiles = myIO.CreateSelectedFiles(files, cmb_folder, rb_summary.Checked, rb_nam.Checked)
                    .ToArray();

                foreach (var file in selectedFiles)
                {
                    var dataList = new List<string[]>();
                    var headerList = new List<string>();
                    var titleList = new List<string>();

                    titleList.AddRange(new string[3] { Path.GetFileNameWithoutExtension(file.Name), "", "" });

                    var newPath = new DirectoryInfo(Path.Combine(file.DirectoryName, "Interpolated Data"));

                    if (!newPath.Exists)
                        Directory.CreateDirectory(newPath.FullName);

                    #region
                    switch (saveVariable)
                    {
                        case SaveVariable.GB:
                            // create the data from the file
                            var admData = new AlbaLibrary.Data.Admittance(file, '\t');

                            headerList.AddRange(
                                new string[3]
                        {"Frequency ("+admData.Frequency.Unit + ")",
                         "Conductance (" + admData.Conductance.Unit + ")" ,
                         "Susceptance (" + admData.Susceptance.Unit + ")" ,
                        });

                            if (interp)
                            {
                                var InterpolationType = InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex);
                                var interpF = admData.Frequency.Data;

                                if (rb_pnts.Checked)
                                    interpF = interpF.ReturnInterpolationXPnts(Convert.ToInt16(txt_nopnts.Text)).ToArray();
                                else
                                    interpF = interpF.ReturnInterpolationXStep(Convert.ToDouble(txt_freqstep.Text)).ToArray();

                                var interpG = InterpolateClass.InterpolateYData(admData.Frequency.Data, admData.Conductance.Data, interpF, InterpolationType);
                                var interpB = InterpolateClass.InterpolateYData(admData.Frequency.Data, admData.Susceptance.Data, interpF, InterpolationType);


                                dataList.Add(Array.ConvertAll<double, string>(interpG.ToArray()[0].ToArray(), Convert.ToString));
                                dataList.Add(Array.ConvertAll<double, string>(interpG.ToArray()[1].ToArray(), Convert.ToString));
                                dataList.Add(Array.ConvertAll<double, string>(interpB.ToArray()[1].ToArray(), Convert.ToString));
                            }
                            else
                            {
                                dataList.Add(Array.ConvertAll<double, string>(admData.Frequency.Data, Convert.ToString));
                                dataList.Add(Array.ConvertAll<double, string>(admData.Conductance.Data, Convert.ToString));
                                dataList.Add(Array.ConvertAll<double, string>(admData.Susceptance.Data, Convert.ToString));
                            }
                            break;

                        case SaveVariable.ZTheta:
                            // create the data from the file
                            var impeData = new AlbaLibrary.Data.Impedance(file, '\t');

                            headerList.AddRange(
                                new string[3]
                        {"Frequency ("+impeData.Frequency.Unit + ")",
                         "Impedance Magnitude (" + impeData.ImpedanceMagnitude.Unit + ")" ,
                         "Impedance Phase (" + impeData.ImpedancePhase.Unit + ")" ,
                        });
                            if (interp)
                            {
                                var InterpolationType = InterpolateClass.ReturnInterpolationType(cmbx_intertype.SelectedIndex);
                                var interpF = impeData.Frequency.Data;

                                if (rb_pnts.Checked)
                                    interpF = interpF.ReturnInterpolationXPnts(Convert.ToInt16(txt_nopnts.Text)).ToArray();
                                else
                                    interpF = interpF.ReturnInterpolationXStep(Convert.ToDouble(txt_freqstep.Text)).ToArray();

                                var interpZ = InterpolateClass.InterpolateYData(impeData.Frequency.Data, impeData.ImpedanceMagnitude.Data, interpF, InterpolationType);
                                var interpZT = InterpolateClass.InterpolateYData(impeData.Frequency.Data, impeData.ImpedancePhase.Data, interpF, InterpolationType);

                                dataList.Add(Array.ConvertAll<double, string>(interpZ.ToArray()[0].ToArray(), Convert.ToString));
                                dataList.Add(Array.ConvertAll<double, string>(interpZ.ToArray()[1].ToArray(), Convert.ToString));
                                dataList.Add(Array.ConvertAll<double, string>(interpZT.ToArray()[1].ToArray(), Convert.ToString));
                            }
                            else
                            {
                                dataList.Add(Array.ConvertAll<double, string>(impeData.Frequency.Data, Convert.ToString));
                                dataList.Add(Array.ConvertAll<double, string>(impeData.ImpedanceMagnitude.Data, Convert.ToString));
                                dataList.Add(Array.ConvertAll<double, string>(impeData.ImpedancePhase.Data, Convert.ToString));
                            }
                            break;
                    }
                    #endregion

                    var WriteXL = new WriteXL(titleList.ToArray(), headerList.ToArray(), dataList.ToArray(), Path.Combine(newPath.FullName, file.Name), "Data", AlbaLibrary.XL.WriteXL.WriteXLMode.Xls);
                    this.Cursor = Cursors.Default;
                }
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Save Tranraw Data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// Assign the default values for drop down lists and graph
        /// </summary>
        private void AssignDefault()
        {    
            zgc_graph.GraphPane.Title.Text = "";
            zgc_graph.GraphPane.XAxis.Title.Text = "Frequency (Hz)";
            zgc_graph.GraphPane.YAxis.Title.Text = "Conductance (S)";

            zgc_graph.SetupZedGraph();
            zgc_graph.GraphPane.Legend.IsVisible = false;
            zgc_graph.GraphPane.XAxis.Scale.Max = 800;
            zgc_graph.GraphPane.YAxis.Scale.Max = 10;
            zgc_graph.AxisChange();
            zgc_graph.Refresh();

            cmbx_plotSelect.SelectedIndex = 0;
            cmbx_intertype.SelectedIndex = 0;
            cmbx_autoplot.SelectedIndex = 0;
            cmbx_atplot.SelectedIndex = 0;

            AutoGraphs = new ZedGraph.GraphPane[7];
        }
        /// <summary>
        /// Setting the main graph to an auto generated graph
        /// </summary>
        /// <param name="variableIndex">index of auto generated graph</param>
        private void SetAutoGraph(int variableIndex)
        {
            try
            {
                if (AutoGraphs != null)
                {
                    var Graphs = AutoGraphs.Cast<ZedGraph.GraphPane>().ToArray();
                    zgc_graph.GraphPane.CurveList = Graphs[variableIndex].CurveList;

                    zgc_graph.GraphPane.Title.Text = Graphs[cmbx_autoplot.SelectedIndex].Title.Text;
                    zgc_graph.GraphPane.XAxis.Title.Text = Graphs[variableIndex].XAxis.Title.Text;
                    zgc_graph.GraphPane.YAxis.Title.Text = Graphs[variableIndex].YAxis.Title.Text;

                    zgc_graph.GraphPane.YAxis.Type = Graphs[variableIndex].YAxis.Type;

                    Graphs[cmbx_autoplot.SelectedIndex].XAxis.Scale.Clone(zgc_graph.GraphPane.XAxis);
                    Graphs[cmbx_autoplot.SelectedIndex].YAxis.Scale.Clone(zgc_graph.GraphPane.YAxis);

                    zgc_graph.SetupZedGraph();
                }
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Set Auto Graph", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// delegate for invoking a textbox from cross-thread
        /// </summary>
        /// <param name="textBox">textbox to invoke</param>
        /// <param name="newValue">value to assign</param>
        private void InvokeTextBox(TextBox textBox, string newValue)
        {
            if (textBox.InvokeRequired)
                textBox.Invoke((MethodInvoker)delegate { textBox.Text = newValue; });
            else
                textBox.Text = newValue;
        }
        
        /// <summary>
        /// Event handlers for the entire form below this summary
        /// </summary>
        /// <param name="sender">the object on the form making the request</param>
        /// <param name="e">the arguement sent by the object</param>

        private void txt_directory_TextChanged(object sender, EventArgs e)
        {
            LoadFiles(txt_directory.Text, true);
        }

        private void cmb_folder_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_folder.SelectedIndex > -1)
                myIO.SwitchFiles(txt_directory.Text, ((ComboBox)sender).SelectedIndex, cmb_folder, lstbx_data, rb_summary.Checked, rb_nam.Checked);
            
        }

        private void lstbx_data_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstbx_data.Items.Count > 0 && Busy == false)
            {
                Busy = true;
                var myListBox = (ListBox)sender;
                PlotAppropriateFile(txt_directory.Text, myListBox, rb_summary.Checked, rb_nam.Checked);
                Busy = false;
            }
        }

        private void GraphForm_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            AssignDefault();
        }

        private void cmbx_plotSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstbx_data.Items.Count > 0)
                SetAutoGraph(cmbx_plotSelect.SelectedIndex);
        }

        private void rb_transraw_CheckedChanged(object sender, EventArgs e)
        {
            EnableSelection((RadioButton)sender, cmbx_plotSelect);
        }

        private void rb_summary_CheckedChanged(object sender, EventArgs e)
        {
            lstbx_data.SelectedIndex = -1;
            EnableSelection((RadioButton)sender, cmbx_summarySelect);

            if (cmb_folder.SelectedIndex > -1)
                myIO.SwitchFiles(txt_directory.Text, cmb_folder.SelectedIndex, cmb_folder, lstbx_data, rb_summary.Checked, rb_nam.Checked);          
        }

        private void cmbx_summarySelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstbx_data.Items.Count > 0)
                PlotAppropriateFile(txt_directory.Text, lstbx_data, rb_summary.Checked, rb_nam.Checked);
            
        }

        private void btn_plot_Click(object sender, EventArgs e)
        {
            if (!Busy)
            {
                Busy = true;
                PlotAt(txt_directory.Text, cmbx_atplot.SelectedIndex);
                Busy = false;
            }
        }

        private void btn_auto_Click(object sender, EventArgs e)
        {
            if (!Busy)
            {
                Busy = true;
                AutoPlot(txt_directory.Text, cmbx_autoplot.SelectedIndex, AutoGraphs.Cast<ZedGraph.GraphPane>().ToArray());
                Busy = false;
            }
        }

        private void graphOptionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var toolBox = new ToolBox(ref zgc_graph);
        }

        private void graphToolStripMenuItem_Click(object sender, EventArgs e)
        {
            zgc_graph.SavePng(txt_directory.Text);
        }

        private void dataToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void onScreenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveOnScreenData(zgc_graph);
        }

        private void gBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveAll(txt_directory.Text, SaveVariable.GB, false);
        }

        private void zToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveAll(txt_directory.Text, SaveVariable.ZTheta, false);
        }

        private void cmbx_autoplot_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbx_autoplot.SelectedIndex < 7 && lstbx_data.Items.Count > 0)
            {
                SetAutoGraph(cmbx_autoplot.SelectedIndex);
            }
        }

        private void rb_num_CheckedChanged(object sender, EventArgs e)
        {
            LoadFiles(txt_directory.Text, false);
        }

        private void txt_linewidth_TextChanged(object sender, EventArgs e)
        {
            foreach (ZedGraph.CurveItem item in zgc_graph.GraphPane.CurveList)
                ((ZedGraph.LineItem)item).Line.Width = (float)txt_linewidth.Text.GetDecimal();
            
            zgc_graph.SetupZedGraph();
        }

        private void cmbx_atplot_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstbx_data.Items.Count > 0 && Busy == false)
            {
                Busy = true;
                PlotAt(txt_directory.Text, cmbx_atplot.SelectedIndex);
                Busy = false;
            }
        }

        private void interpolatedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveAll(txt_directory.Text, SaveVariable.GB, true);
        }

        private void interpolatedToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            SaveAll(txt_directory.Text, SaveVariable.ZTheta, true);
        }

        private void individualTranRawToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveTransRawData(txt_directory.Text, SaveVariable.GB, true);
        }

        private void individualTranRawToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            SaveTransRawData(txt_directory.Text, SaveVariable.ZTheta, true);
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

    }
}
