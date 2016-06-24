using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using LumenWorks.Framework.IO.Csv;
using Excel = Microsoft.Office.Interop.Excel; 

namespace MyIVI
{
    public partial class MyIVI : Form
    {
        List<Records> object_Records = new List<Records>();
        List<OutputRecords> object_OutputRecords = new List<OutputRecords>();
        List<Cumulative_Frequency> cumulative_Frequency = new List<Cumulative_Frequency>();
      
        public MyIVI()
        {
            InitializeComponent();

        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Do_checked();
        }
        private void Do_checked()
        {
            btnSubmit.Enabled = checkBox1.Checked;
        }

        private void MyIVI_Load(object sender, EventArgs e)
        {
        }
     
        private String text2, text3, text4, text5, text6;
        private double text7,text8,text9;
        private int text1;
        private int text10 = 0;
        public static int different_Plot_Numbers = 0;
        public static int number_Of_OutputRecords = 0;
        public static int total_Genera = 0;
        public static int total_Species = 0;
        public static int total_Family = 0;
        public static int temp_Cumulative_Count = 0;
        public static double total_Density = 0;
        public static double total_Basal_Area = 0;
        public static double total_Frequency = 0;
        public static double highest_Ivi_Value = 0;
        public static double total_Shanon_Deviation = 0;
        public static double total_Shanon_Deviation2 = 0;
        public static string dominant_Species;

        private static int count = 0; // static variable that stores the record that is read at present.

        private void button1_Click(object sender, EventArgs e)
        {
            string temp_Show_Caution = "The following Headers must be specified with exact spelling in the input file for correct results : 'PLOT' , 'Botanical Name' , 'DBH' , 'FAMILY' ";
            MessageBox.Show(temp_Show_Caution);
            openFileDialog1.Filter = "CSV files (*.csv)|*.csv";  // Show only .csv files among all the different files        
            DialogResult result = openFileDialog1.ShowDialog();
            checkBox1.Enabled = true;
            if (result == DialogResult.OK) // Test result.
            {
                String file = openFileDialog1.FileName;
                try
                {
                    textBoxFilePath.Text = file;
                    using (CsvReader csv = new CsvReader(new StreamReader(file), true))
                    {
                        int fieldCount = csv.FieldCount;
                        string[] headers = csv.GetFieldHeaders();
                        // Code for checking the presence of the header Botanical Name in the Input File.
                        bool botanical_Name_Flag = false;
                        for (int a = 0; a < headers.Count(); a++)
                        {
                            if(headers[a].ToLower() == "botanical name" )
                            {
                                botanical_Name_Flag = true;
                            }
                        }
                        if(botanical_Name_Flag == false)
                        {
                            MessageBox.Show("The Botanical Name Header is missing in the Input File.", "Important Message");
                            Application.Exit();
                        }
                        // Code for checking the presence of header Plot in the Input File.
                        bool plot_Flag = false;
                        for (int a = 0; a < headers.Count(); a++)
                        {
                            if (headers[a].ToLower() == "plot")
                            {
                                plot_Flag = true;
                            }
                        }
                        if (plot_Flag == false)
                        {
                            MessageBox.Show("The Plot Header is missing in the Input File.", "Important Message");
                            Application.Exit();
                        }
                        //Code for checking the presence of header DBH in the Input File.
                        bool dbh_Flag = false;
                        for (int a = 0; a < headers.Count(); a++)
                        {
                            if (headers[a].ToLower() == "dbh")
                            {
                                dbh_Flag = true;
                            }
                        }
                        if (dbh_Flag == false)
                        {
                            MessageBox.Show("The DBH Header is missing in the Input File.", "Important Message");
                            Application.Exit();
                        }
                        //Code for checking the presence of header Family in the Input File.
                        bool family_Flag = false;
                        for (int a = 0; a < headers.Count(); a++)
                        {
                            if (headers[a].ToLower() == "family")
                            {
                                family_Flag = true;
                            }
                        }
                        if (family_Flag == false)
                        {
                            MessageBox.Show("The Family Header is missing in the Input File.", "Important Message");
                            Application.Exit();
                        }

                            while (csv.ReadNextRecord())
                            {
                                count += 1;
                                for (int i = 0; i < fieldCount; i++)
                                {
                                    try
                                    {
                                        switch (headers[i].ToLower())
                                        {
                                            case "plot":
                                                string pass_Message_Box_Plot;
                                                bool res = int.TryParse(csv[i], out text1);
                                                pass_Message_Box_Plot = "The PLOT number of the records cannot be empty. Please check the record: " + count.ToString() + " Start the application after verifying that record!";
                                                if (res == false)
                                                {
                                                    MessageBox.Show(pass_Message_Box_Plot, "Important Message");
                                                    Application.Exit();
                                                }
                                                break;
                                            case "local name":
                                                text2 = csv[i];
                                                break;
                                            case "botanical name":
                                                string pass_Message_Box_Botanical_Name;
                                                text3 = csv[i];
                                                pass_Message_Box_Botanical_Name = "The Botanical Name of the records cannot be empty. Please check the record: " + (count+1).ToString() + " Start the application after verifying that record!";
                                                if (string.IsNullOrEmpty(text3))
                                                {
                                                    MessageBox.Show(pass_Message_Box_Botanical_Name, "Important Message");
                                                    Application.Exit();
                                                }

                                                break;
                                            case "genera":
                                                text4 = csv[i];
                                                break;
                                            case "species":
                                                text5 = csv[i];
                                                break;
                                            case "family":
                                                text6 = csv[i];
                                                string pass_Message_Box4;
                                                int count_Plus = count + 1;
                                                if (string.IsNullOrEmpty(text6))
                                                {
                                                    pass_Message_Box4 = "Do you want to continue with 'EMPTY' fields in Family column of the Input file with Record Number: " + count_Plus.ToString() + " ?  This may produce wrong results!";
                                                    DialogResult result1 = MessageBox.Show(pass_Message_Box4, "Important Message!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                                                    if (result1 == DialogResult.No)
                                                    {
                                                        Utilities.ResetAllControls(this);
                                                        Application.Exit();
                                                    }
                                                }
                                                break;
                                            case "dbh":
                                                // Code for verifying Empty Fields and Handling them!
                                                string pass_Message_Box1;
                                                Double.TryParse(csv[i], out text7);
                                                if (text7 == 0)
                                                {
                                                    pass_Message_Box1 = "Do you want to continue with 'EMPTY' fields in DBH column of the Input file with Record Number: " + (count+1).ToString() + " ?  This may produce wrong results!";
                                                    DialogResult result1 = MessageBox.Show(pass_Message_Box1, "Important Message!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                                                    if (result1 == DialogResult.No)
                                                    {
                                                        Utilities.ResetAllControls(this);
                                                        Application.Exit();
                                                    }
                                                }
                                                break;
                                            /*             case "ba(sqm)":
                                                             Double.TryParse(csv[i], out text8);
                                                             if (text8 == 0)
                                                             {
                                                                string pass_Message_Box2 = "Do you want to continue with 'EMPTY' fields in BA column of the Input file with Record Number: " + count.ToString() + " ?  This may produce wrong results!";
                                                                DialogResult result1 = MessageBox.Show(pass_Message_Box2, "Important Message!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);                                  
                                                                 if (result1 == DialogResult.No)
                                                                 {
                                                                     Utilities.ResetAllControls(this);
                                                                     Application.Exit();
                                                                 }
                                                             }
                                                             break;
                                                         case "ba":
                                                             Double.TryParse(csv[i], out text8);
                                                             if (text8 == 0)
                                                             {
                                                                 string pass_Message_Box3 = "Do you want to continue with EMPTY fields in BA column of the input file with row number:  " + count.ToString() + " ? This may produce wrong results!";
                                                                 DialogResult result1 = MessageBox.Show(pass_Message_Box3,"Important Message!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                                                                 if (result1 == DialogResult.No)
                                                                 {
                                                                     Utilities.ResetAllControls(this);
                                                                     Application.Exit();
                                                                 }
                                                             }
                                                             break;*/
                                            case "height":
                                                Double.TryParse(csv[i], out text9);
                                                /*            if (text8 == 0)
                                                            {
                                                                DialogResult result1 = MessageBox.Show("Do you want to continue with EMPTY fields in Height column of your file? This may produce wrong results!",
                                                                    "Important Message!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                                                                if (result1 == DialogResult.No)
                                                                {
                                                                    Utilities.ResetAllControls(this);
                                                                    Application.Exit();
                                                                }
                                                            }*/
                                                break;
                                            case "":
                                                MessageBox.Show(" Null Header!");
                                                break;
                                            default:
                                                //                       MessageBox.Show("Please verify the column headers of the file. They must match with the headers mentioned to obtain desired output!", " Important Message!");
                                                //                       Application.Exit();
                                                break;
                                        }
                                    }
                                    catch (System.ArgumentException)
                                    {
                                        MessageBox.Show("You are trying to input a file with similar headers for columns. Please correct them!");
                                        Application.Exit();
                                    }
                                }
                                object_Records.Add(new Records(text1, text2, text3, text4, text5, text6, text7, text8, text9, text10));
                            }
                        object_Records.Sort(delegate(Records x, Records y) // Arranges the records stored in ascending order.
                        {
                            return x.plot.CompareTo(y.plot);
                        });
                        Find_Total_Plots(ref object_Records);        // Call to fuction that finds the number of plots in the file.                       
                        Find_Total_Family(ref object_Records);
                        Find_Density_BasalArea_Frequency(ref object_Records, ref object_OutputRecords); // call to function that groups the species and sets in object_Outputrecords.
                        Find_Relative_Density_Basal_Area_Frequency(ref object_OutputRecords); // call to function that finds the relative density,relative basal area and relative frequency.
                        Find_Ivi_Value(ref object_OutputRecords);
                        Find_Shanon_Deviation(ref object_OutputRecords);
                        Find_Total_Genera_Species(ref object_OutputRecords);
                        Find_Cumulative_Frequency(ref object_Records,ref cumulative_Frequency);
                    }
                    textBox1.Text = count.ToString();
                }
                catch (IOException)
                {
                    MessageBox.Show("Please correct the Input File and retry!");
                }
                catch(System.ArgumentException)
                {
                    MessageBox.Show("The Input file is having a problem with the headers. Please correct them and retry!!");
                    Application.Exit();
                }
                catch(NullReferenceException)
                {
                    MessageBox.Show("The Input file is having a problem with the headers. Please correct them and retry!!");
                    Application.Exit();
                }
                catch(IndexOutOfRangeException)
                {
                    Application.Exit();
                }
            }
            else
            {
                checkBox1.Enabled = false;
            }
        }
        private void btnQuit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            Utilities.ResetAllControls(this);
            Application.Restart();
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
 /*           DataGridViewTextBoxColumn bnColumn = new DataGridViewTextBoxColumn();
            bnColumn.DataPropertyName = "Botanical_Name";
            bnColumn.Width = 130;
            bnColumn.HeaderText = "Botanical Name";

            DataGridViewTextBoxColumn generaColumn = new DataGridViewTextBoxColumn();
            generaColumn.DataPropertyName = "Genera";
            generaColumn.HeaderText = "Genera";

            DataGridViewTextBoxColumn speciesColumn = new DataGridViewTextBoxColumn();
            speciesColumn.DataPropertyName = "Species";
            speciesColumn.HeaderText = "Species";

            DataGridViewTextBoxColumn familyColumn = new DataGridViewTextBoxColumn();
            familyColumn.DataPropertyName = "Family";
            familyColumn.HeaderText = "Family";

            DataGridViewTextBoxColumn individualsColumn = new DataGridViewTextBoxColumn();
            individualsColumn.DataPropertyName = "Individuals";
            individualsColumn.HeaderText = "Individuals";
            DataGridViewTextBoxColumn countColumn = new DataGridViewTextBoxColumn();
            countColumn.DataPropertyName = "Count";
            countColumn.HeaderText = "Count";
            DataGridViewTextBoxColumn frequencyColumn = new DataGridViewTextBoxColumn();
            frequencyColumn.DataPropertyName = "Frequency";
            frequencyColumn.HeaderText = "Frequency";
            DataGridViewTextBoxColumn densityColumn = new DataGridViewTextBoxColumn();
            densityColumn.DataPropertyName = "Density";
            densityColumn.HeaderText = "Density";
            DataGridViewTextBoxColumn densityColumn = new DataGridViewTextBoxColumn();
            densityColumn.DataPropertyName = "Density";
            densityColumn.HeaderText = "Density";
            DataGridViewTextBoxColumn densityColumn = new DataGridViewTextBoxColumn();
            densityColumn.DataPropertyName = "Density";
            densityColumn.HeaderText = "Density";
            DataGridViewTextBoxColumn densityColumn = new DataGridViewTextBoxColumn();
            densityColumn.DataPropertyName = "Density";
            densityColumn.HeaderText = "Density";

            dataGridView1.Columns.Add(bnColumn);
            dataGridView1.Columns.Add(generaColumn);
            dataGridView1.Columns.Add(speciesColumn);
            dataGridView1.Columns.Add(familyColumn);
            dataGridView1.Columns.Add(individualsColumn);
            dataGridView1.Columns.Add(countColumn);
            dataGridView1.Columns.Add(frequencyColumn);*/

            this.dataGridView1.AutoGenerateColumns = true;            
            dataGridView1.Visible = true;
            dataGridView1.DataSource = object_OutputRecords;

            textBox2.Text = highest_Ivi_Value.ToString();
            textBox3.Text = dominant_Species;
            textBoxGenera.Text = total_Genera.ToString();
            textBoxSpecies.Text = total_Species.ToString();
            textBoxFamiles.Text = total_Family.ToString();
            textBoxShanonIndex.Text = total_Shanon_Deviation.ToString();
            textBoxShanonIndex2.Text = total_Shanon_Deviation2.ToString();
            
            checkBox1.Enabled = false;
            btnSubmit.Visible = false;
            btnChooseFile.Visible = false;
            textBoxFilePath.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            label7.Visible = false;
            label11.Visible = false;
            label10.Visible = false;
            label8.Visible = false;
            checkBox1.Visible = false;

            buttonLoadChart.Visible = true;
            
            btnGenerate.Visible = true;           
            
            btnShowOutputData.Visible = true;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
       
        // Code to find the total number of plots in the given file.
        public static void Find_Total_Plots(ref List<Records> object_Records)
        {
            for (int i = 0; i < count; i++)                 // Code to find the total number of plots in the given file.
            {
                if (object_Records[i].plot > different_Plot_Numbers)
                {
                    different_Plot_Numbers++;
                }
            }
        }

        public static int count_Value;

        // Code for finding the Cumulative Frequency of the Plots.
        public static void Find_Cumulative_Frequency(ref List<Records> object_Records, ref List<Cumulative_Frequency> cumulative_Frequency)
        {
            List<int> plot_Removed_Duplicates = new List<int>();
            List<int> total_Plots = new List<int>();
            
            for (int i = 0; i < count; i++)
            {
                total_Plots.Add(object_Records[i].plot);
                if (object_Records[i].flag3 == false)
                {
                    for (int j = i+1 ; j < count; j++)
                    {
                        if ((object_Records[i].botanical_Name.Equals(object_Records[j].botanical_Name)) && (object_Records[j].flag3== false))
                        {
                            object_Records[j].flag3 = true;
                        }
                    }
                }
            }
            for(int k=0; k<count ;k++)
            {
                if(object_Records[k].flag3 == false)
                {
                    plot_Removed_Duplicates.Add(object_Records[k].plot);
                }
            }
            
            List<int> distinct_Total_Plots = total_Plots.Distinct().ToList();
            int total_Distinct_Plots = distinct_Total_Plots.Count();
            plot_Removed_Duplicates.Sort();
            plot_Removed_Duplicates.Add(0);
            int total_Cumulative_Plots = plot_Removed_Duplicates.Count() ;
            int temp = plot_Removed_Duplicates[0];
            int temp_Add_Plot;
            int temp_Count = 0;
            int h;
            int l = total_Cumulative_Plots;
            int p = l-1;
            
            for(int i=0 ; i< total_Cumulative_Plots; i++)
            {
                                            
                if(plot_Removed_Duplicates[i] ==  temp)
                {
                    temp_Count++;
                }
                else
                {
                    h = (i - 1);
                    temp_Add_Plot = plot_Removed_Duplicates[h]; 
                    temp = plot_Removed_Duplicates[i];                 
                    cumulative_Frequency.Add(new Cumulative_Frequency(temp_Add_Plot, temp_Count));
                    
                    temp_Count = 1;
                }
             }
            int cumulative_Count = cumulative_Frequency.Count();
            for(int a=0; a<cumulative_Count; a++)
            {
                for(int b=0; b < total_Distinct_Plots; b++)
                {
                    if(cumulative_Frequency[a].plot_Number == distinct_Total_Plots[b])
                    {
                        distinct_Total_Plots.RemoveAt(b);
                        break;
                    }
                }
            }
            int timepass_Count = distinct_Total_Plots.Count();
            for(int c=0; c<timepass_Count; c++)
            {
                int temp_Add_Plot_Number;
                temp_Add_Plot_Number = distinct_Total_Plots[c];
                cumulative_Frequency.Add(new Cumulative_Frequency(temp_Add_Plot_Number, 0));
            }
            cumulative_Frequency.Sort(delegate(Cumulative_Frequency x, Cumulative_Frequency y) // Arranges the records stored in ascending order.
            {
                return x.plot_Number.CompareTo(y.plot_Number);
            });
            
            for(int d=0; d< total_Distinct_Plots; d++)
            {    
                cumulative_Frequency[d].Count = cumulative_Frequency[d].Count + temp_Cumulative_Count ;
                temp_Cumulative_Count = cumulative_Frequency[d].Count;
            }
          }

        
        // Code to find total number of families in the given file.
        public static void Find_Total_Family(ref List<Records> object_Records)
        {
            for (int i = 0; i < count; i++)
            {
                if (object_Records[i].flag2 == false)
                {
                    for (int j = i + 1; j < count; j++)
                    {
                        if ((object_Records[i].family == object_Records[j].family) && (object_Records[j].flag2 == false))
                        {
                            object_Records[j].flag2 = true;
                        }
                    }
                }
            }
            for (int i = 0; i < count; i++)
            {
                if (object_Records[i].flag2 == false)
                {
                    total_Family++;
                }
            }
        }
        

        // Code that computes the density, basal area and frequency of each species in the given file.
        public static void Find_Density_BasalArea_Frequency( ref List<Records> object_Records, ref List<OutputRecords> object_OutputRecords)
        {
            int temp_Count_Botanical_Number = 0;
            for (int i = 0; i < count; i++)
            {
                temp_Count_Botanical_Number++;
                List<int> temp_Count_Duplicate_Plots = new List<int>();
                if (object_Records[i].flag == false)
                {
                    int temp_Count = 0;
                    int temp_Individuals = 0;
                    double temp_Basal_Area = 0;
                    string temp_Botanical_Name;
                    double temp_Density = 0;
                    string temp_Genera;
                    string temp_Species;
                    string temp_Family;
                    string temp_Extra_Phrases;
                    temp_Family = object_Records[i].family;
                    temp_Botanical_Name = object_Records[i].botanical_Name;
                    string[] words = temp_Botanical_Name.Split(' ');
                    if(words.Count() != 2)
                    {
                        temp_Extra_Phrases = "The botanical name in record number: " + temp_Count_Botanical_Number.ToString() +" is not correctly entered. Do you still want to continue? ";
                        DialogResult result1 = MessageBox.Show(temp_Extra_Phrases, "Important Message!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                        if (result1 == DialogResult.No)
                        {
                            Application.Exit();
                        }
                    }
                   
                    temp_Genera = words[0];
                    if (words.Count() == 1)
                    {
                        temp_Species = null;
                    }
                    else
                    {
                        temp_Species = words[1];
                    }
                    double temp_Frequency = 0;
                    
                    for (int j = i; j < count; j++)
                    {

                        if (object_Records[i].botanical_Name.Equals(object_Records[j].botanical_Name))
                        {
                            object_Records[j].flag = true;
                            temp_Individuals++;
                            temp_Basal_Area = temp_Basal_Area + object_Records[j].basal_Area;
                            temp_Count_Duplicate_Plots.Add(object_Records[j].plot);
                        }
                    }
                    List<int> distinct = temp_Count_Duplicate_Plots.Distinct().ToList();
                    temp_Count = distinct.Count;
                    temp_Density = ((double)temp_Individuals / (double)different_Plot_Numbers) * 100;
                    temp_Frequency = ((double)temp_Count / (double)different_Plot_Numbers) * 100;
                    total_Density = total_Density + temp_Density;           // Code for computing total density of all the species.
                    total_Basal_Area = total_Basal_Area + temp_Basal_Area;  // Code for computing total basal area of all species.
                    total_Frequency = total_Frequency + temp_Frequency;     // Code for computing total frequency of all the species.
                    object_OutputRecords.Add(new OutputRecords(temp_Individuals, temp_Count, temp_Botanical_Name, temp_Genera, temp_Species, temp_Family, temp_Basal_Area, temp_Density, temp_Frequency));
                    number_Of_OutputRecords++;
                }
            }
            object_OutputRecords.Sort(delegate(OutputRecords x, OutputRecords y) // Arranges the records stored in ascending order.
            {
                return x.botanical_Name.CompareTo(y.botanical_Name);
            });
        }

        
        // Code that computes the relative density, relative basal area, relative frequency of each species.
        public static void Find_Relative_Density_Basal_Area_Frequency(ref List<OutputRecords> object_OutputRecords)
        {
            
            for(int i=0; i<number_Of_OutputRecords; i++)
            {
                object_OutputRecords[i].relative_Density = (((double)object_OutputRecords[i].density) / ((double)total_Density)) * 100;
                object_OutputRecords[i].relative_Basal_Area = (((double)object_OutputRecords[i].species_Basal_Area) / ((double)total_Basal_Area)) * 100;
                object_OutputRecords[i].relative_Frequency = (((double)object_OutputRecords[i].frequency) / ((double)total_Frequency)) * 100;
            }
        }
       

        // Code that computes the IVI (Importance value Index) value.
        public static void Find_Ivi_Value(ref List<OutputRecords> object_OutputRecords)
        {
            
            for(int i=0; i<number_Of_OutputRecords; i++)
            {
                object_OutputRecords[i].ivi_Value = (object_OutputRecords[i].relative_Frequency) + (object_OutputRecords[i].relative_Density) + (object_OutputRecords[i].relative_Basal_Area);
                if (object_OutputRecords[i].ivi_Value >= highest_Ivi_Value)
                {
                    highest_Ivi_Value = object_OutputRecords[i].ivi_Value;
                    dominant_Species = object_OutputRecords[i].botanical_Name;
                }
            }
        }

       
        // Code that computes the Shanon Deviation.
        public static void Find_Shanon_Deviation(ref List<OutputRecords> object_OutputRecords)
        {
            // Code for computing shanon index value to the base 2.
            double temp_value;
            for(int i=0; i<number_Of_OutputRecords; i++)
            {
                temp_value = (double)object_OutputRecords[i].individuals / (double)(count);
                object_OutputRecords[i].shanon_Deviation = temp_value * (Math.Log(temp_value, 2));
                total_Shanon_Deviation = total_Shanon_Deviation + object_OutputRecords[i].shanon_Deviation;
            }
            total_Shanon_Deviation = -(total_Shanon_Deviation);
            // Code for computing shanon index value to the base 10.
            double temp_value2;
            for (int i = 0; i < number_Of_OutputRecords; i++)
            {
                temp_value2 = (double)object_OutputRecords[i].individuals / (double)(count);
                object_OutputRecords[i].shanon_Deviation2 = temp_value2 * (Math.Log(temp_value2, 10));
                total_Shanon_Deviation2 = total_Shanon_Deviation2 + object_OutputRecords[i].shanon_Deviation2;
            }
            total_Shanon_Deviation2 = -(total_Shanon_Deviation2);
        }

       
        // Code that computes total genera and total species.
        public static void Find_Total_Genera_Species(ref List<OutputRecords> object_OutputRecords)
        {
            for(int i=0; i < number_Of_OutputRecords; i++)
            {
                if (object_OutputRecords[i].flag1 == false)
                {
                    for (int j = i+1; j < number_Of_OutputRecords; j++)
                    {
                        if ((object_OutputRecords[i].genera == object_OutputRecords[j].genera) && (object_OutputRecords[j].flag1 == false))
                        {
                            object_OutputRecords[j].flag1 = true;
                        }
                    }
                }
                if(object_OutputRecords[i].flag2 == false)
                {
                    for(int k=i+1; k<number_Of_OutputRecords; k++)
                    {
                        if((object_OutputRecords[i].species == object_OutputRecords[k].species) && (object_OutputRecords[k].flag2 == false))
                        {
                            object_OutputRecords[k].flag2 = true;
                        }
                    }
                }
            }
            for (int i = 0; i < number_Of_OutputRecords; i++)
            {
                if(object_OutputRecords[i].flag1==false)
                {
                    total_Genera++;
                }
                if(object_OutputRecords[i].flag2==false)
                {
                    total_Species++;
                }
            }
        }
     
        
        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxFilePath_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label13_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxGenera_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {
                    }

        private void label8_Click_1(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            
            
        }

        private void chart1_Click(object sender, EventArgs e)
        {
                    }

        private void chart1_Click_1(object sender, EventArgs e)
        {

        }


        // Code for Load Chart Button
        private void buttonLoadChart_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Visible = false;
                chart1.Visible = true;
                int temp_Count = cumulative_Frequency.Count();
                this.chart1.ChartAreas[0].AxisX.IsStartedFromZero = true;
                this.chart1.ChartAreas[0].AxisY.IsStartedFromZero = true;
                this.chart1.ChartAreas[0].AxisX.Minimum = 0;
                this.chart1.ChartAreas[0].AxisY.Minimum = 0;
                this.chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                this.chart1.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
                this.chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                this.chart1.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
                //         this.chart1.ChartAreas[0].AxisX.MajorGrid.LineWidth = 0;
                //         this.chart1.ChartAreas[0].AxisY.MajorGrid.LineWidth = 0;
                this.chart1.Titles.Add("Species Accumulation Curve");
                this.chart1.ChartAreas[0].AxisX.Title = " Number of Plots ";
                this.chart1.ChartAreas[0].AxisY.Title = " Cumulative Number of Species ";
                this.chart1.Series["Graph"].Points.AddXY(0, 0);
                for (int i = 0; i < temp_Count; i++)
                {
                    this.chart1.Series["Graph"].Points.AddXY(cumulative_Frequency[i].plot_Number, cumulative_Frequency[i].Count);
                }                          
            }
            catch (DirectoryNotFoundException)
            {
                MessageBox.Show("Path wrong dude!");
            }
            
            btnSaveGraph.Visible = true;
            btnShowOutputData.Visible = true;
            buttonLoadChart.Visible = false;
            btnShowGraph.Visible = true;
        }


        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp ;
            Excel.Workbook xlWorkBook ;
            Excel.Worksheet xlWorkSheet ;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int i = 0;
            int j = 0; 

            for (i = 0; i <= dataGridView1.RowCount  - 1; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount  - 1; j++)
                {
                    DataGridViewCell cell = dataGridView1[j, i];
                    xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
                }
            }
            SaveFileDialog dlg2 = new SaveFileDialog();
            dlg2.Title = "Specify Destination Filename";
            dlg2.FilterIndex = 1;
            dlg2.OverwritePrompt = true;
            dlg2.DefaultExt = "xls";
            dlg2.Filter = "eXcel files (*.xls)|*.xls";
            if (dlg2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                xlWorkBook.SaveAs(dlg2.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
                btnGenerate.Enabled = false;
            }
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        

        private void btnSaveGraph_Click(object sender, EventArgs e)
        {           
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Title = "Specify Destination Filename";
            dlg.FilterIndex = 1;
            dlg.OverwritePrompt = true;            
            dlg.DefaultExt = "Png";         
            dlg.Filter = "Jpeg files (*.Jpeg)|*.Jpeg";
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.chart1.SaveImage(dlg.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Jpeg);
                btnSaveGraph.Enabled = false;
                
            }
            else
            {
                btnSaveGraph.Visible = true;
            }          
        }

        private void btnShowOutputData_Click(object sender, EventArgs e)
        {
            chart1.Visible = false;
            dataGridView1.Visible = true;
          
            
        }

        private void btnShowGraph_Click(object sender, EventArgs e)
        {
            dataGridView1.Visible = false;
            chart1.Visible = true;
        }

    }
    public class Cumulative_Frequency
    {
        public int plot_Number;
        public int count_Value_Of_Plot;
        public Cumulative_Frequency(int plot_Number1, int count_Value_Of_Plot1)
        {
            this.plot_Number = plot_Number1;
            this.count_Value_Of_Plot = count_Value_Of_Plot1;
        }
        public int Plot_Number
        {
            get { return plot_Number; }
            set { plot_Number = value; }
        }
        public int Count
        {
            get { return count_Value_Of_Plot; }
            set { count_Value_Of_Plot = value; }
        }
    }
    // class to store the output.
    public class OutputRecords
    {
        public int individuals;
        public int count;
        public string botanical_Name;
        public string genera;
        public string species;
        public string family;
        public double density;
        public double species_Basal_Area;
        public double frequency;
        public double relative_Frequency;
        public double relative_Density;
        public double relative_Basal_Area;
        public double ivi_Value;
        public double shanon_Deviation;
        public double shanon_Deviation2;
        public bool flag1;
        public bool flag2;
        public OutputRecords(int individuals1, int count1, string botanical_Name1, string genera1, string species1, string family1,double species_Basal_Area1,double density1,double frequency1)
        {
            this.individuals = individuals1;
            this.count = count1;
            this.botanical_Name = botanical_Name1;
            this.species_Basal_Area = species_Basal_Area1;
            this.genera = genera1;
            this.species = species1;
            this.family = family1;
            this.density = density1;
            this.frequency = frequency1;
            this.relative_Frequency = 0;
            this.relative_Density = 0;
            this.relative_Basal_Area = 0;
            this.ivi_Value = 0;
            this.shanon_Deviation = 0;
            this.shanon_Deviation2 = 0;
            this.flag1 = false;
            this.flag2 = false;
        }
        public string Botanical_Name
        {
            get { return botanical_Name; }
            set { botanical_Name = value; }
        }
        public string Genera
        {
            get { return genera; }
            set { genera = value; }
        }
        public string Species
        {
            get { return species; }
            set { species = value; }
        }
        public string Family
        {
            get { return family; }
            set { family = value; }
        }
        public int Individuals
        {
            get { return individuals; }
            set { individuals = value; }
        }
        public int Count
        {
            get { return count; }
            set { count = value; }
        }
        public double Frequency
        {
            get { return frequency; }
            set { frequency = value; }
        }
        public double Density
        {
            get { return density; }
            set { density = value; }
        }
        public double Species_Basal_Area
        {
            get { return species_Basal_Area; }
            set { species_Basal_Area = value; }
        }
        public double Relative_Frequency
        {
            get { return relative_Frequency; }
            set { relative_Frequency = value; }
        }
        public double Relative_Density
        {
            get { return relative_Density; }
            set { relative_Density = value; }
        }
        public double Relative_Basal_Area
        {
            get { return relative_Basal_Area; }
            set { relative_Basal_Area = value; }
        }
        public double Ivi_Value
        {
            get { return ivi_Value; }
            set { ivi_Value = value; }
        }
        public double Shanon_Index_Log2
        {
            get { return shanon_Deviation; }
            set { shanon_Deviation = value; }
        }
        public double Shanon_Index_Log10
        {
            get { return shanon_Deviation2; }
            set { shanon_Deviation2 = value; }
        }
    }
   
    // class to store the input csv file records.
    public class Records
    {
            public int plot;
            public string local_Name;
            public string botanical_Name;
            public string genera;
            public string species;
            public string family;
            public double dbh;
            public double basal_Area;
            public double height;
            public int count_Botanical_Name;
            public bool flag;
            public bool flag2;
            public bool flag3;
            public Records(int plot1, string local_Name1, string botanical_Name1, string genera1, string species1, string family1, double dbh1, double basal_Area1, double height1,int count_Botanical_Name1)
            {
                this.plot = plot1;
                this.local_Name = local_Name1;
                this.botanical_Name = botanical_Name1;
                this.genera = genera1;
                this.species = species1;
                this.family = family1;
                this.dbh = dbh1;
                this.basal_Area = (double)(((double)dbh * (double)dbh) / 12.57);
                this.height = height1;
                this.count_Botanical_Name = count_Botanical_Name1;
                this.flag = false;
                this.flag2 = false;
                this.flag3 = false;
            }
            public int Plot
            {
                get { return plot; }
                set { plot = value; }
            }

            public string Local_Name
            {
                get { return local_Name; }
                set { local_Name = value; }
            }
            public string Botanical_Name
            {
                get { return botanical_Name; }
                set { botanical_Name = value; }
            }
            public string Genera
            {
                get { return genera; }
                set { genera = value; }
            }
            public string Species
            {
                get { return species; }
                set { species = value; }
            }
            public string Family
            {
                get { return family; }
                set { family = value; }
            }

            public double Dbh
            {
                get { return dbh; }
                set { dbh = value; }
            }
            public double Basal_Area
            {
                get { return basal_Area; }
                set { basal_Area = value; }
            }
            public double Height
            {
                get { return height; }
                set { height = value; }
            }
            public int Count
            {
                get { return count_Botanical_Name; }
                set { count_Botanical_Name = value; }
            }
           
     }
   
    public class Utilities
        {
            public static void ResetAllControls(Control form)
            {
                foreach (Control control in form.Controls)
                {
                    if (control is TextBox)
                    {
                        TextBox textBox = (TextBox)control;
                        textBox.Text = null;
                    }

                    if (control is ComboBox)
                    {
                        ComboBox comboBox = (ComboBox)control;
                        if (comboBox.Items.Count > 0)
                            comboBox.SelectedIndex = 0;
                    }

                    if (control is CheckBox)
                    {
                        CheckBox checkBox = (CheckBox)control;
                        checkBox.Checked = false;
                    }

                    if (control is ListBox)
                    {
                        ListBox listBox = (ListBox)control;
                        listBox.ClearSelected();
                    }
                    
                }
            }
        }
}

