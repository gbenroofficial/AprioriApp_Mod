using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;


namespace Daniel_s_UI_prototype
{

    

    public partial class Form1 : Form
    {


        int support = 0;
        int transCount = 0;

        string path;

        List<List<string>> DB = new List<List<string>>();
        List<ItemSet> unique_item_list = new List<ItemSet>();
        List<List<ItemSet>> all_freqlists = new List<List<ItemSet>>();
        List<Rule> rules_list = new List<Rule>();

        List<string> checkLST1 = new List<string>();
        List<string> checkLST2 = new List<string>();
        List<string> checkLST = new List<string>();
        

        //List<ItemSet> list1_copy = new List<ItemSet>();
        public Form1()
        {
            InitializeComponent();
            
        }


        
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            checkedListBox1.Enabled = true;
            button5.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            checkedListBox2.Enabled = true;
            button4.Enabled = true;
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

            }
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel File|*.xlsx;*.xls;*.csv";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)

            {
                


                //load excel file (workbook and sheet) through openfiledialog
                MessageBox.Show(openFileDialog1.FileName);
                path = openFileDialog1.FileName;

                //Avoid contamination from previous dataset
                checkedListBox1.Items.Clear();
                checkedListBox2.Items.Clear();
                DB.Clear();
                unique_item_list.Clear();
                all_freqlists.Clear();
                rules_list.Clear();
                listView1.Items.Clear();
                listView1.Refresh();
                dataGridView2.Rows.Clear();
                dataGridView2.Refresh();

                var excel = new Excel.Application(); 
                Excel.Workbook wb;
                Excel.Worksheet ws;

                wb = excel.Workbooks.Open(path);
                ws = wb.Worksheets[1];

                //get max length of rows and columns
                Excel.Range range = ws.UsedRange;
                int rowCount = range.Rows.Count; //max rows
                int colCount = range.Columns.Count; //max columns

                                

            //transfer all data from excel into list<list> database structure
                for (int row = 1; row <= rowCount; row++) //row traversal
                {
                    List<string> Trans = new List<string>(); //represents a new transaction
                    transCount += 1; //keepstrack of total number of transactions to convert support to percentage and back
                                  
                    for (int col = 1; col <= colCount; col++) //column traversal
                    {
                        if (ws.Cells[row, col].Value2 != null) //ensure excel cell is not empty
                        {
                            Trans.Add(ws.Cells[row, col].Value2);
                        }
                        else
                            break;
                    }

                    Trans.Sort();
                    Trans.Insert(0, row.ToString()); //add transaction ID to start of each transaction
                    
                    
                    DB.Add(Trans); //add to dataset structure
                    
                }

                //get all the unique items from list<list> structure
                unique_item_list = unique_items_gen(DB, support);

                foreach(ItemSet itm in unique_item_list)
                {
                    checkedListBox1.Items.Add(itm.Items[0]);
                    checkedListBox2.Items.Add(itm.Items[0]);
                }
                
                
                vIEWDATABASEToolStripMenuItem.Enabled = true;
                numericUpDown1.Enabled = true;
                numericUpDown2.Enabled = true;
                button1.Enabled = true;
                button2.Enabled = true;
                button5.Enabled = false;

                //function that puts header in listview
                // Set the view to show details.
                listView1.View = View.Details;      

                //button3.Enabled = true;
                generate.Enabled = true; //generate button.

            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e) //GENERATE BUTTON
        {           
                     

            List<ItemSet> checkItmSets1 = new List<ItemSet>();
            List<ItemSet> checkItmSets2 = new List<ItemSet>();

            List<string> checkTmp1 = new List<string>();
            List<string> checkTmp2 = new List<string>();

            if (checkLST1.Count != 0) //if checklist box 1 is not empty
            {
                foreach (string itm in checkLST1)
                {
                    ItemSet tmpSet = new ItemSet();
                    List<string> tmpList = new List<string>();


                    tmpList.Add(itm);
                    tmpSet.Items = tmpList;
                    checkItmSets1.Add(tmpSet);
                }

            }

            if (checkLST2.Count != 0) //if checklist box 2 is not empty
            {
                foreach (string itm in checkLST2)
                {
                    ItemSet tmpSet = new ItemSet();
                    List<string> tmpList = new List<string>();


                    tmpList.Add(itm);
                    tmpSet.Items = tmpList;
                    checkItmSets2.Add(tmpSet);
                }

            }

            Stopwatch stopWatch = new Stopwatch();

            listView1.Items.Clear();
            listView1.Refresh();

            dataGridView2.Rows.Clear();
            dataGridView2.Refresh();


            int i = Convert.ToInt32((numericUpDown1.Value / 100) * transCount); //support value conversion
            float c = ((float)numericUpDown2.Value) / 100; //confidence value conversion
            
            

            if (checkLST1.Count() > 0 || checkLST2.Count() > 0) //if any checkboxes are ticked
            {
                if(checkLST1.Count() > 0 && checkLST2.Count > 0) //both are ticked
                {
                    checkTmp1 = checkLST1;
                    checkTmp2 = checkLST2;
                }
                if (checkLST1.Count() > 0 && checkLST2.Count() == 0) //box 1 is ticked
                {
                    checkTmp2 = null;
                    checkTmp1 = checkLST1;
                }
                if (checkLST1.Count() == 0 && checkLST2.Count() > 0) //box 2 is ticked
                {
                    checkTmp1 = null;
                    checkTmp2 = checkLST2;
                }


            }

           if (checkLST1.Count() == 0 && checkLST2.Count() == 0) //both checklist boxes are not ticked
            {
                checkTmp1 = null;
                checkTmp2 = null;

            }


            stopWatch.Start();
            all_freqlists = all_freqlists_gen(DB, i, checkTmp1, checkTmp2); //generate frequent itemsets
            rules_list = rule_Gen(all_freqlists, i, checkTmp1, checkTmp2); //generate rules
            stopWatch.Stop();


            //display frequent items:
            int id = 0;

            if (all_freqlists.Count >= 1)
            { 
                List<string> tempLST = new List<string>();

                foreach (var itm in all_freqlists)
                {
                    foreach( var subitm in itm)
                    {
                        foreach(string str in subitm.Items)
                        {
                            if(tempLST.Contains(str) != true)
                            {
                                tempLST.Add(str);
                            }
                        }
                        
                    }
                }

                {
                    foreach(string str in tempLST)
                    {
                        id++;
                        DataGridViewTextBoxCell txtbox10 = new DataGridViewTextBoxCell();
                        DataGridViewTextBoxCell txtbox20 = new DataGridViewTextBoxCell();

                        txtbox10.Value = id;
                        txtbox20.Value = str;

                        DataGridViewRow row = new DataGridViewRow();
                        row.Cells.Add(txtbox10);
                        row.Cells.Add(txtbox20);
                        dataGridView2.Rows.Add(row);

                    }

                }
            }
            


            

            
            //display rules:
            if (rules_list.Count >= 1)
            {
                int j = 0;
                foreach (var rule in rules_list)
                {

                    j++;
                    if (rule.confidence >= c) //filter the rules based on confidence
                    {
                        string ante = null;
                        string conse = null;

                        float conf = (rule.confidence) * 100;
                        float supp = (rule.support / transCount) * 100;


                        foreach (var itm in rule.getAnte)
                        {
                            ante += itm + " ";
                        }

                        foreach (var itm2 in rule.getConse)
                        {
                            conse += itm2 + " ";
                        }


                        string[] itms = { ante, conse, supp.ToString(), conf.ToString() };
                        ListViewItem item = new ListViewItem(itms);
                        listView1.Items.Add(item);


                    }



                }

            }
            


            
            TimeSpan t_span = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00} h:{1:00} m:{2:00} s.{3:000} ms",
            t_span.Hours, t_span.Minutes, t_span.Seconds,
            t_span.Milliseconds);
            textBox1.Text = elapsedTime;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void splitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            button2.Enabled = true;
            

            checkLST1.Clear();
            foreach (object itmChecked in checkedListBox1.CheckedItems)
            {
                checkLST1.Add(itmChecked.ToString());

            }


            int[] indexes1 = checkedListBox1.CheckedIndices.Cast<int>().ToArray();
            foreach(int itm in indexes1)
            {
                checkedListBox2.SetItemChecked(itm, false);
            }


            button5.Enabled = false;
            checkedListBox1.Enabled = false;


        }

        private void button4_Click(object sender, EventArgs e)
        {
            checkLST2.Clear();

            foreach (object itmChecked in checkedListBox2.CheckedItems)
            {
                checkLST2.Add(itmChecked.ToString());
            }



            int[] indexes2 = checkedListBox2.CheckedIndices.Cast<int>().ToArray();
            foreach (int itm in indexes2)
            {
                checkedListBox1.SetItemChecked(itm, false);
            }

            button4.Enabled = false;
            checkedListBox2.Enabled = false;
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            List<string> checkedItems = new List<string>();
            foreach (var item in checkedListBox1.CheckedItems)
                checkedListBox2.Items.Remove(item); //if item checked in box 1 uncheck in box 2
        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        
        /// //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        

        //functions:

        //unique_items_gen()
        List<ItemSet> unique_items_gen(List<List<string>> db, int supp)
        {

            List<ItemSet> unique_items = new List<ItemSet>(); //List with unique items
            List<ItemSet> supported_items = new List<ItemSet>(); //List with unique item matching support


            foreach (List<string> trans in db) //scanning rows in database
            {

                foreach (string item in trans) //scanning columns in transaction
                {
                    if (item != trans[0])
                    {

                        //unique_items.Where(x => x.itemName == item).Select(x=> { x.itemSupport += 1; return x; }).ToList();
                        foreach (ItemSet unique_item in unique_items) //search unique item list
                        {
                            if (unique_item.Items[0] == item) //if found again in database transactions
                            {
                                unique_item.Support += 1;
                                unique_item.Ids.Add(trans[0]);
                            }
                        }


                        if (!unique_items.Any(x => x.Items[0] == item)) //if not present in list of items, create and add to list
                        {
                            List<string> tmpList = new List<string>();
                            List<string> idList = new List<string>();

                            tmpList.Add(item);
                            idList.Add(trans[0]);
                            unique_items.Add(new ItemSet { Items = tmpList, Support = 1, Ids = idList });
                        }

                    }                 

                }
            }

            //Add unique items that meet support to different list
            foreach (ItemSet unique_item in unique_items)
            {
                if (unique_item.Support >= supp)
                {
                    supported_items.Add(unique_item);
                }

            }
            supported_items = supported_items.OrderBy(x => x.Items[0]).ToList();
            return supported_items;
        }



        //check_support():
        float check_support(List<List<string>> db, ItemSet gen_candid)
        {
            float sum = 0;
            foreach (List<string> trans in db)
            {

                if (gen_candid.Items.Count() > 1 && (string.Compare(gen_candid.Items.First(), trans.First()) == -1 || string.Compare(gen_candid.Items.Last(), trans.Last()) == 1))
                    {
                    break;
                } //if the first and/or last element of itemset of itemset fall out of range of first and/or last item in transaction

                else
                {

                    if (!gen_candid.Items.Except(trans).Any())
                    {
                        sum += 1;
                        gen_candid.Support = sum;
                    }

                }
                //if (!gen_candid.Items.Except(trans).Any())
                //{
                //    sum += 1;
                //    gen_candid.Support = sum;
                //}

            }
            return gen_candid.Support;
        }



        //Apriori_gen():
        List<ItemSet> Apriori_gen(List<ItemSet> kMinus1)
        {
            List<ItemSet> K = new List<ItemSet>();

            if (kMinus1.Count != 0) //if list of itemsets is not empty
            {


                int size = kMinus1[0].Items.Count(); //number of items in each itemset


                if (size == 1 && kMinus1.Count > 1) //for item size 1 and list of itemsets of at least 2
                {
                    for (int i = 0; i < kMinus1.Count(); i++)
                    {
                        for (int j = i + 1; j < kMinus1.Count(); j++)
                        {
                            List<string> tmp = new List<string>();
                            tmp.Add(kMinus1[i].Items[0]);
                            tmp.Add(kMinus1[j].Items[0]);



                            K.Add(new ItemSet { Items = tmp });
                        }
                    }

                }

                if (size > 1 && kMinus1.Count > 1)
                {

                    for (int i = 0; i < kMinus1.Count(); i++)
                    {
                        for (int j = i + 1; j < kMinus1.Count(); j++)
                        {
                            List<string> tmp1 = new List<string>();
                            List<string> tmp2 = new List<string>();

                            for (int a = 0; a <= size - 2; a++)
                            {
                                tmp1.Add(kMinus1[i].Items[a]);
                                tmp2.Add(kMinus1[j].Items[a]);
                            }


                            if (tmp1.SequenceEqual(tmp2))
                            {
                                tmp1.Add(kMinus1[i].Items[size - 1]);
                                tmp1.Add(kMinus1[j].Items[size - 1]);

                                K.Add(new ItemSet { Items = tmp1 });
                            }

                        }
                    }
                }
            }

            return K;
        }


        //unique
        List<ItemSet> unique_ante(ItemSet set) //takes apart all items in a single itemset into individual itemsets in a list.
        {
            List<ItemSet> allItems = new List<ItemSet>();
            foreach (string itm in set.Items)
            {
                ItemSet temp = new ItemSet();
                List<string> tmpStr = new List<string>();
                tmpStr.Add(itm);
                temp.Items = tmpStr;
                temp.Support = set.Support;
                allItems.Add(temp);
            }
            return allItems;
        }



        float confidence_gen(List<List<ItemSet>> freqItems, ItemSet X, ItemSet Y)
        {
            ItemSet temp = new ItemSet();
            float a;
            float b;
            //int X_size = X.items.Count();
            //int Y_size = Y.items.Count();

            //foreach (ItemSet itmset in freqItems[X_size - 1])
            //{
            //    if (itmset.items.Except(X.items) == null)
            //    {
            //        X.support = itmset.support;
            //    }
            //}

            foreach (List<ItemSet> setList in freqItems) //setList is list of itemsets of same size
            {
                foreach (ItemSet itemset in setList) //find the antecedent support from list of frequent itemsets of varying sizes
                {
                    bool isEqual = Enumerable.SequenceEqual(X.Items, itemset.Items);
                    if (isEqual)
                    {
                        X.Support = itemset.Support;
                    }
                }
            }

            //join the items in antecedent and consequent
            List<string> tmpStr = new List<string>();
            foreach (string item in X.Items)
            {
                tmpStr.Add(item);
            }
            foreach (string itm in Y.Items)
            {
                tmpStr.Add(itm);
                tmpStr.Sort();
            }

            //foreach (string itm in tmpStr)
            //{
            //    temp.items.Add(itm);
            //}

            temp.Items = tmpStr;

            foreach (List<ItemSet> setList in freqItems)
            {
                foreach (ItemSet itemset in setList)
                {
                    bool isEqual = Enumerable.SequenceEqual(temp.Items, itemset.Items);
                    if (isEqual)
                    {
                        temp.Support = itemset.Support;
                    }
                }
            }

            
            a = temp.Support;
            b = X.Support;
          
            return (a / b);
        }


        List<List<ItemSet>> freqList_gen(List<List<ItemSet>> prev_freqs, int supp)
        {
            
            List<ItemSet> freqList = new List<ItemSet>();
            
            List<List<ItemSet>> fullList = new List<List<ItemSet>>();
            fullList.Add(prev_freqs[0]);
            
            do
            {
                List<ItemSet> extFreqList = new List<ItemSet>();
                if (freqList.Count() != 0)
                {                    
                    foreach(ItemSet itm in freqList)
                    {
                        extFreqList.Add(itm);
                    }

                    fullList.Add(extFreqList);
                }


                List<ItemSet> candList = new List<ItemSet>();
                candList = Apriori_gen(fullList[fullList.Count() - 1]);
                freqList.Clear();
                
                foreach (ItemSet candidSet in candList) // each candidate itemset
                {  
                    ItemSet itm1 = new ItemSet();
                    ItemSet itm2 = new ItemSet();

                    List<ItemSet> itmList = new List<ItemSet>();
                    itmList = unique_ante(candidSet); //split all items in itemset to individual 1-size itemsets and add to list

                    foreach(ItemSet itmset in fullList[0]) //update support and ids using single size itemset already given
                    {
                        foreach(ItemSet itmset2 in itmList)
                        {
                            if (itmset.Items.SequenceEqual(itmset2.Items))
                            {
                                itmset2.Support = itmset.Support;
                                itmset2.Ids = itmset.Ids;

                            }
                        }
                    }


                    ItemSet tmpItm = new ItemSet();
                    tmpItm = itmList.OrderBy(x => x.Support).FirstOrDefault(); //should return itemset with minimum support



                    itm1 = tmpItm;

                    List<string> tmpList = new List<string>();
                    List<string> tmp = new List<string>();

                    
                    tmpList = candidSet.Items.Except(itm1.Items).ToList(); //the rest of the items that are not minimum
                    

                    foreach (ItemSet ItmSet in fullList[tmpList.Count() - 1])
                    {
                        if (tmpList.SequenceEqual(ItmSet.Items)) //find itemset that matches the items in rest of the items above
                        {
                            itm2 = ItmSet;
                        }
                    }

                   if(itm2.Ids != null)
                    {
                        if (itm1.Ids.Intersect(itm2.Ids).Count() >= supp) //check the support value using the ids of the split groups
                        {
                            candidSet.Support = itm1.Ids.Intersect(itm2.Ids).Count();
                            candidSet.Ids = itm1.Ids.Intersect(itm2.Ids).ToList();
                            freqList.Add(candidSet);
                        }
                    }
                    
                    
                }               
                

            } while (freqList.Count() != 0);




            return fullList;
        }


        //all_freqlists_gen():
        List<List<ItemSet>> all_freqlists_gen(List<List<string>> db, int supp, List<string> ante = null, List<string> conse = null)
        {
            List<ItemSet> freqList = new List<ItemSet>();
            List<ItemSet> seedList = new List<ItemSet>();
            List<List<ItemSet>> fullList = new List<List<ItemSet>>();
            List<List<ItemSet>> tmpList = new List<List<ItemSet>>();
            seedList = unique_items_gen(db, supp); //frequent-1 itemset generation
            freqList = seedList;


            if (freqList.Count() > 0)
            {
                tmpList.Add(freqList);
                fullList = freqList_gen(tmpList, supp);

               
                List<string> tmpSet = new List<string>();
                if (ante != null && conse != null)
                {
                    tmpSet = ante.Union(conse).ToList();
                }

                if (ante != null && conse == null)
                {
                    tmpSet = ante;
                }

                if (conse != null && ante == null)
                {
                    tmpSet = conse;
                }


                List<List<ItemSet>> tmpFullList = new List<List<ItemSet>>();

                if (tmpSet.Count() != 0)
                {
                    foreach (List<ItemSet> setList in fullList)
                    {
                        List<ItemSet> tmpSetLst = new List<ItemSet>();

                        foreach (ItemSet itm in setList)
                        {
                            if (!ante.Except(itm.Items).Any() || !tmpSet.Except(itm.Items).Any()) //This step is important for having antecedent support and confidence gen
                            {
                                tmpSetLst.Add(itm);
                            }  
                        }
                        tmpFullList.Add(tmpSetLst);

                    }

                    fullList.Clear();
                    fullList = tmpFullList;
                }
            }

            return fullList;

        }



        //RuleGen 
        List<Rule> rule_Gen(List<List<ItemSet>> all_freq_list, float supp, List<string> ante=null, List<string> conse=null)
        {
            //list of rules to be returned
            List<Rule> rulesList = new List<Rule>();

            foreach (List<ItemSet> group in all_freq_list)
            {
                foreach (ItemSet seed in group)
                {
                    if (seed.Items.Count > 1)
                    {
                        List<ItemSet> anteList = new List<ItemSet>();
                        List<ItemSet> singles = new List<ItemSet>();


                        singles = unique_ante(seed); //need to update with support and ids
                       
                       
                        anteList.AddRange(singles.ToArray()); 

                        if(seed.Items.Count() > 2)
                        {

                            for (int i = 0; i < seed.Items.Count() - 2; i++) //generates antecedents of size 2 and above
                            {
                                singles = Apriori_gen(singles);
                                foreach (ItemSet single in singles)
                                {
                                    single.Support = seed.Support;
                                }
                                anteList.AddRange(singles.ToArray()); 

                            }

                        }


                        //anteList = anteList.GroupBy(x => x.Items).Select(y => y.First()).ToList();
                        List<ItemSet> anteSpec = new List<ItemSet>();
                        anteSpec.AddRange(anteList);



                        //select only antecedent matching user given antecedent  
                        if (ante != null)
                        {
                            anteSpec.Clear();
                            foreach(ItemSet set in anteList)
                            {
                                if (!ante.Except(set.Items).Any())
                                {
                                    anteSpec.Add(set);
                                }
                            }
                        }



                        //for each antecedent generated, create the consequent to form a new rule and add to list of rules
                        foreach (ItemSet itm in anteSpec)
                        {
                            List<string> tempList = new List<string>();
                            tempList = seed.Items.Except(itm.Items).ToList(); //items not in antecedent i.e. consequent

                            


                            ItemSet newAnte = new ItemSet();
                            ItemSet newCons = new ItemSet();

                            List<string> tmpStr = new List<string>();
                            tmpStr.AddRange(itm.Items.ToArray());

                            newAnte.Items = tmpStr;
                            newCons.Items = tempList;

                            
                            //create new rule object
                            Rule tempRule = new Rule();

                            tempRule.antecedent = newAnte;
                            tempRule.consequent = newCons;
                            tempRule.confidence = confidence_gen(all_freq_list, newAnte, newCons);
                            tempRule.support = itm.Support;
                            
                            if (conse != null)
                            {
                                if (tempRule.support >= supp && !conse.Except(tempRule.consequent.Items).Any()) //&& tempRule.support > 0 and the user consequents found in potential rule
                                {
                                    rulesList.Add(tempRule);
                                }
                            }
                            if (conse == null)
                            {
                                if (tempRule.support >= supp)
                                {
                                    rulesList.Add(tempRule);
                                }
                            }
                            
                            
                        }
                    }

                }
            }
            
            
            return rulesList;
        }

        
        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView4_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //this.dataGridView4.Rows[e.RowIndex].Cells["rn"].Value = (e.RowIndex + 1).ToString();
        }

        private void vIEWDATABASEToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            var excel = new Excel.Application();
            Excel.Workbook wb;
            Excel.Worksheet ws;

            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];
            excel.Visible= true;
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }
    }
}
