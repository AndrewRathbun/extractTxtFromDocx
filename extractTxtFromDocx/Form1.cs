using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Compression;
using ExtensionMethods;
using System.IO;
using System.Text.RegularExpressions;

namespace extractTxtFromDocx
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            tabPage3.HidePage();
            tabPage4.HidePage();
            tabPage5.HidePage();
            tabPage6.HidePage();
        }

        private void button1_Click(object sender, EventArgs e)
        {



        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        List<String> blocks = new List<string>();
        List<String> lineBlocks = new List<string>();
        List<String> lineBlocksWithTags = new List<string>();
        String textWithTags = "";

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string lineB = "";
            string line = "";
            int countt = 0;
            int pFrom = 0, pTo = 0, pCheck = 0;
            int lFrom = 0, lTo = 0, lCheck = 0;
            string zipPath;
            string extractPath = @"C:\extracted";
            string extractedTxtPath = "", extractedPropsCore = "", extractedPropsApp = "", extractedContents = "", extractedDocument = "", extractedFootNotes = "", extractedEndNotes = "";
            bool propsCore = false, propsApp = false, footNotes = false, endNotes = false;
            string extractedTxt = "";
            List<String> ID = new List<string>();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();



            openFileDialog1.InitialDirectory = @"C:\";

            openFileDialog1.Title = "Browse Docx Files";



            openFileDialog1.CheckFileExists = true;

            openFileDialog1.CheckPathExists = true;



            openFileDialog1.DefaultExt = "docx";

            openFileDialog1.Filter = "Docx files (*.docx)|*.docx|All files (*.*)|*.*";

            openFileDialog1.FilterIndex = 1;

            openFileDialog1.RestoreDirectory = true;



            openFileDialog1.ReadOnlyChecked = true;

            openFileDialog1.ShowReadOnly = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)

            {
                if (tabPage3.IsVisible())
                    tabPage3.HidePage();
                if (tabPage4.IsVisible())
                    tabPage4.HidePage();
                if (tabPage5.IsVisible())
                    tabPage5.HidePage();
                if (tabPage6.IsVisible())
                    tabPage6.HidePage();
                
                String result = "";
                int error = 0;
                System.IO.DirectoryInfo directory = new System.IO.DirectoryInfo(extractPath);
                Directory.CreateDirectory("C:\\Lines");
                foreach (System.IO.FileInfo file in directory.GetFiles()) file.Delete();
                foreach (System.IO.DirectoryInfo subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
                //string startPath = @"c:\example\start";
                //string zipPath = @"c:\example\result.zip";

                zipPath = openFileDialog1.FileName;
                //System.IO.Compression.ZipFile.CreateFromDirectory(startPath, zipPath);
                try
                {
                    System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, extractPath);
                }
                catch (Exception ex)
                {
                    error = 1;
                    result = ex.Message.ToString();
                }

                if (error == 1)
                {
                    label2.Text = "Please selcet correct docx file\nMore Details: " + result.ToString();
                    textBox1.Text = "Please selcet correct docx file\nMore Details: " + result.ToString();
                }
                else
                {

                    try
                    {


                        extractedContents = @"C:\extracted\[Content_Types].xml";

                        extractedTxt = System.IO.File.ReadAllText(extractedContents);
                        pTo = 0;
                        int count = 0;
                    ploop:
                        pFrom = extractedTxt.IndexOf("PartName=\"", pTo) + "PartName=\"".Length;
                        pFrom = extractedTxt.IndexOf("/", pFrom) + 1;
                        pTo = extractedTxt.IndexOf(".xml", pFrom) + ".xml".Length;
                        pCheck = extractedTxt.IndexOf("PartName=\"", pTo);

                        if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains("core.xml"))
                        {
                            extractedPropsCore = extractedTxt.Substring(pFrom, pTo - pFrom);
                            propsCore = true;
                        }
                        if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains("app.xml"))
                        {
                            extractedPropsApp = extractedTxt.Substring(pFrom, pTo - pFrom);
                            propsApp = true;
                        }

                        if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains("document.xml"))
                        {
                            extractedDocument = extractedTxt.Substring(pFrom, pTo - pFrom);
                            propsApp = true;
                        }

                        if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains("footnotes.xml"))
                        {
                            extractedFootNotes = extractedTxt.Substring(pFrom, pTo - pFrom);
                            footNotes = true;
                        }
                        if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains("endnotes.xml"))
                        {
                            extractedEndNotes = extractedTxt.Substring(pFrom, pTo - pFrom);
                            endNotes = true;
                        }
                        result += ++count + "- " + extractedTxt.Substring(pFrom, pTo - pFrom) + "\n\n";
                        if (pCheck > 0)
                            goto ploop;


                        label2.Text = result.ToString();
                    }
                    catch (Exception ex)
                    {
                        label2.Text = ex.Message.ToString();
                    }

                    //-----------------------------------------DOCUMENT.XML---------------------------------------------------

                    //try
                    {
                        int tCheck = 0;
                        result = "";
                        extractedTxtPath = @"C:\extracted\" + extractedDocument;
                        string extractedTxt1;
                        extractedTxt1 = System.IO.File.ReadAllText(extractedTxtPath);
                        int count = 0;
                        lCheck = 1;
                        lTo = 0;
                        textWithTags = extractedTxt1;
                        List<String> edits = new List<string>();
                        //String[] edits = new String[100];
                        while (lCheck > 0)
                        {
                            if (countt != 0)
                                textBox1.AppendText(Environment.NewLine);
                            countt++;
                            count = 0;
                            /*if ((extractedTxt1.IndexOf("<w:tbl", lTo)) >= (extractedTxt1.IndexOf("</w:tbl>", lTo)) && extractedTxt1.IndexOf("</w:tbl>", lTo) != -1)
                            {
                                break;
                            }*/

                            if ((extractedTxt1.IndexOf("<w:p w", lTo)) >= (extractedTxt1.IndexOf("<w:tbl", lTo)) && extractedTxt1.IndexOf("<w:tbl", lTo) != -1)
                            {

                                tCheck = 1;
                                lFrom = extractedTxt1.IndexOf("<w:tbl", lTo) + "<w:tbl".Length;
                                lFrom = extractedTxt1.IndexOf(">", lFrom) + 1;
                                lTo = extractedTxt1.IndexOf("</w:tbl>", lFrom);
                                lTo += "</w:tbl>".Length;
                                lCheck = extractedTxt1.IndexOf("<w:p w", lTo);
                                extractedTxt = extractedTxt1.Substring(lFrom, lTo - lFrom);
                            }
                            else
                            {
                                tCheck = 0;
                                lFrom = extractedTxt1.IndexOf("<w:p w", lTo) + "<w:p w".Length;
                                line = extractedTxt1.Substring(extractedTxt1.IndexOf("w:rsidR=\"", lTo) + "w:rsidR=\"".Length, extractedTxt1.IndexOf("\"", lFrom) - lFrom);
                                lFrom = extractedTxt1.IndexOf(">", lFrom) + 1;
                                lTo = extractedTxt1.IndexOf("</w:p>", lFrom);
                                if (lTo < 0)
                                    break;
                                lCheck = extractedTxt1.IndexOf("<w:p w", lTo);
                                extractedTxt = extractedTxt1.Substring(lFrom, lTo - lFrom);
                                /*textBox1.AppendText("----------------------------------------------------");
                                textBox1.AppendText(Environment.NewLine);
                                textBox1.AppendText(extractedTxt);
                                textBox1.AppendText(Environment.NewLine);
                                textBox1.AppendText("----------------------------------------------------");
                                textBox1.AppendText(Environment.NewLine);
                                textBox1.AppendText(lCheck.ToString());
                                textBox1.AppendText(Environment.NewLine);*/
                                //textBox1.AppendText(line);
                                if (!ID.Contains(line))
                                {
                                    ID.Add(line);
                                }
                                if (ID.IndexOf(line) == 0)
                                    textBox1.AppendText("ID: " + 1 + " "); //for disable first line showing ID 0
                                else
                                    textBox1.AppendText("ID: " + ID.IndexOf(line) + " ");
                                textBox1.AppendText("LINE " + countt.ToString() + ":  ");
                            }
                            edits.Clear();
                            pTo = 0;
                            lineBlocksWithTags.Add(extractedTxt);
                            lineB = "";
                        ploop1:
                            count++;
                            if (tCheck == 1)
                            {
                                textBox1.AppendText(Environment.NewLine);
                                textBox1.AppendText("**TABLE: ");
                                pFrom = extractedTxt.IndexOf("<w:t>", pTo) + "<w:t>".Length;
                                pTo = extractedTxt.IndexOf("</w:t>", pFrom);
                                if (pTo < 0)
                                {
                                    lineBlocks.Add(lineB);
                                    continue;
                                }
                                pCheck = extractedTxt.IndexOf("<w:t", pTo);

                            }
                            else
                            {
                            anotherloop:
                                if (!extractedTxt.Contains("<w:t>") && !extractedTxt.Contains("<w:t xml:space=\"preserve\">"))
                                {
                                    lineBlocks.Add(lineB);
                                    continue;
                                }
                                if (extractedTxt.Contains("<w:textAlignment"))
                                {
                                    lineBlocks.Add(lineB);
                                    continue;
                                }
                                pFrom = extractedTxt.IndexOf("<w:t", pTo) + "<w:t".Length;
                                if (extractedTxt.Substring(pFrom - 4, 10).Contains("<w:tab"))
                                {
                                    pFrom = extractedTxt.IndexOf(">", pFrom) + 1;
                                    extractedTxt = extractedTxt.Substring(pFrom, extractedTxt.Length - pFrom);
                                    pTo = 0;
                                    goto anotherloop;
                                }
                                pFrom = extractedTxt.IndexOf(">", pFrom) + 1;
                                pTo = extractedTxt.IndexOf("</w:t>", pFrom);
                                if (pTo < 0)
                                {
                                    lineBlocks.Add(lineB);
                                    continue;
                                }
                                pCheck = extractedTxt.IndexOf("<w:t", pTo);
                            }

                            edits.Add(extractedTxt.Substring(pFrom, pTo - pFrom));
                            //edits[count-1] = extractedTxt.Substring(pFrom, pTo - pFrom);
                            blocks.Add(extractedTxt.Substring(pFrom, pTo - pFrom));
                            lineB += extractedTxt.Substring(pFrom, pTo - pFrom);
                            textBox1.AppendText(extractedTxt.Substring(pFrom, pTo - pFrom));



                            if (pCheck > 0)
                                goto ploop1;
                            lineBlocks.Add(lineB);
                            if (tCheck == 1 && count > 1)
                            {
                                textBox1.AppendText(Environment.NewLine);
                                textBox1.AppendText("**The above table has " + count.ToString() + " cells");
                            }
                            else if (count > 1)
                            {
                                textBox1.AppendText(Environment.NewLine);
                                textBox1.AppendText("**Something added to the above line " + (count - 1).ToString() + " Times");
                                if (extractedTxt.Contains(line))
                                {
                                    pFrom = extractedTxt.IndexOf(line, 0) + line.Length;
                                    pFrom = extractedTxt.IndexOf("<w:t", pFrom) + "<w:t".Length;
                                    pFrom = extractedTxt.IndexOf(">", pFrom) + 1;
                                    pTo = extractedTxt.IndexOf("</w:t>", pFrom);
                                    textBox1.AppendText("   |  Original was: \"" + extractedTxt.Substring(pFrom, pTo - pFrom) + "\"");
                                    textBox1.AppendText(" Added: ");
                                    for (int i = 0; i < count; i++)
                                    {
                                        if (edits[i] == extractedTxt.Substring(pFrom, pTo - pFrom))
                                            continue;
                                        textBox1.AppendText("\"" + edits[i].ToString() + "\" ");
                                    }
                                }
                                else
                                {
                                    textBox1.AppendText("   |  Original was: \"" + edits[0].ToString() + "\"");
                                    textBox1.AppendText(" Added: ");
                                    for (int i = 1; i < count; i++)
                                    {
                                        textBox1.AppendText("\"" + edits[i].ToString() + "\" ");
                                    }
                                }
                            }


                        }


                    }
                   /* catch (Exception ex)
                    {
                        textBox1.Text = ex.Message.ToString();
                    }*/


                    //-----------------------------------------CORE.XML---------------------------------------------------

                    if (propsCore)
                    {
                        tabPage3.ShowPageInTabControl(tabControl1);
                        string title = "", created = "", modified = "", keywords = "", authors = "", modifiedDate = "", createdDate = "";
                        string modifiedTimes = "";
                        try
                        {

                            result = "";
                            extractedPropsCore = @"C:\extracted\" + extractedPropsCore;

                            extractedTxt = System.IO.File.ReadAllText(extractedPropsCore);

                            pFrom = extractedTxt.IndexOf("<dc:title>");
                            if (pFrom > 0)
                            {
                                pFrom = extractedTxt.IndexOf("<dc:title>", 0) + "<dc:title>".Length;
                                pTo = extractedTxt.IndexOf("</dc:title>", pFrom);

                                title = extractedTxt.Substring(pFrom, pTo - pFrom);
                            }

                            pFrom = extractedTxt.IndexOf("<cp:keywords>");
                            if (pFrom > 0)
                            {
                                pFrom = extractedTxt.IndexOf("<cp:keywords>", 0) + "<cp:keywords>".Length;
                                pTo = extractedTxt.IndexOf("</cp:keywords>", pFrom);

                                keywords = extractedTxt.Substring(pFrom, pTo - pFrom);
                            }

                            pFrom = extractedTxt.IndexOf("<cp:revision>");
                            if (pFrom > 0)
                            {
                                pFrom = extractedTxt.IndexOf("<cp:revision>", 0) + "<cp:revision>".Length;
                                pTo = extractedTxt.IndexOf("</cp:revision>", pFrom);

                                int revision = Convert.ToInt32(extractedTxt.Substring(pFrom, pTo - pFrom));
                                revision = revision / 2;

                                modifiedTimes = "modofied " + revision + " times";
                            }

                            pFrom = extractedTxt.IndexOf("<cp:lastModifiedBy>");
                            if (pFrom > 0)
                            {
                                pFrom = extractedTxt.IndexOf("<cp:lastModifiedBy>", 0) + "<cp:lastModifiedBy>".Length;
                                pTo = extractedTxt.IndexOf("</cp:lastModifiedBy>", pFrom);

                                modified = extractedTxt.Substring(pFrom, pTo - pFrom);
                            }

                            pFrom = extractedTxt.IndexOf("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">");
                            if (pFrom > 0)
                            {
                                pFrom = extractedTxt.IndexOf("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">", 0) + "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">".Length;
                                pTo = extractedTxt.IndexOf("</dcterms:modified>", pFrom);

                                modifiedDate = extractedTxt.Substring(pFrom, pTo - pFrom - 1);
                                modifiedDate = modifiedDate.Replace('T', ' ');
                            }

                            pFrom = extractedTxt.IndexOf("<dcterms:created xsi:type=\"dcterms:W3CDTF\">");
                            if (pFrom > 0)
                            {
                                pFrom = extractedTxt.IndexOf("<dcterms:created xsi:type=\"dcterms:W3CDTF\">", 0) + "<dcterms:created xsi:type=\"dcterms:W3CDTF\">".Length;
                                pTo = extractedTxt.IndexOf("</dcterms:created>", pFrom);

                                createdDate = extractedTxt.Substring(pFrom, pTo - pFrom - 1);
                                createdDate = createdDate.Replace('T', ' ');
                            }

                            pFrom = extractedTxt.IndexOf("<dc:creator>");
                            if (pFrom > 0)
                            {
                                pFrom = extractedTxt.IndexOf("<dc:creator>", 0) + "<dc:creator>".Length;
                                pTo = extractedTxt.IndexOf("</dc:creator>", pFrom);
                                if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains(";"))
                                {
                                    string txt = extractedTxt.Substring(pFrom, pTo - pFrom);
                                    string[] txtResult;
                                    txtResult = txt.Split(';');

                                    foreach (string s in txtResult)
                                    {
                                        authors += s + " , ";
                                    }
                                    created = txtResult[0];
                                }
                                else
                                    created = extractedTxt.Substring(pFrom, pTo - pFrom);
                            }

                            label3.Text = "Title: " + title;
                            label4.Text = "Tags: " + keywords;
                            label5.Text = "Authors: " + authors;
                            label6.Text = "Creator: created by " + created + " on " + createdDate;
                            label7.Text = "Modifier: modified by " + modified + " on " + modifiedDate;
                            label8.Text = "Modify: " + modifiedTimes;
                        }
                        catch (Exception ex)
                        {
                            label3.Text = ex.Message.ToString();
                        }

                    }

                    //-----------------------------------------APP.XML---------------------------------------------------

                    if (propsApp)
                    {
                        tabPage4.ShowPageInTabControl(tabControl1);

                        string characters = "", charactersWithSpaces = "", words = "", paragraphs = "";
                        try
                        {

                            result = "";
                            extractedPropsApp = @"C:\extracted\" + extractedPropsApp;

                            extractedTxt = System.IO.File.ReadAllText(extractedPropsApp);

                            pFrom = extractedTxt.IndexOf("</Words>");
                            if (pFrom > 0)
                            {
                                pFrom = extractedTxt.IndexOf("<Words>", 0) + "<Words>".Length;
                                pTo = extractedTxt.IndexOf("</Words>", pFrom);

                                words = extractedTxt.Substring(pFrom, pTo - pFrom);
                            }

                            pFrom = extractedTxt.IndexOf("<Characters>");
                            if (pFrom > 0)
                            {
                                pFrom = extractedTxt.IndexOf("<Characters>", 0) + "<Characters>".Length;
                                pTo = extractedTxt.IndexOf("</Characters>", pFrom);

                                characters = extractedTxt.Substring(pFrom, pTo - pFrom);
                            }

                            pFrom = extractedTxt.IndexOf("<CharactersWithSpaces>");
                            if (pFrom > 0)
                            {
                                pFrom = extractedTxt.IndexOf("<CharactersWithSpaces>", 0) + "<CharactersWithSpaces>".Length;
                                pTo = extractedTxt.IndexOf("</CharactersWithSpaces>", pFrom);

                                charactersWithSpaces = extractedTxt.Substring(pFrom, pTo - pFrom);
                            }

                            pFrom = extractedTxt.IndexOf("<Paragraphs>");
                            if (pFrom > 0)
                            {
                                pFrom = extractedTxt.IndexOf("<Paragraphs>", 0) + "<Paragraphs>".Length;
                                pTo = extractedTxt.IndexOf("</Paragraphs>", pFrom);

                                paragraphs = extractedTxt.Substring(pFrom, pTo - pFrom);
                            }

                            label13.Text = "Paragraphs: " + paragraphs;
                            label14.Text = "Words: " + words;
                            label15.Text = "Characters: " + characters;
                            label16.Text = "Characters With Spaces: " + charactersWithSpaces;
                            label17.Text = "Lines: " + countt;
                            label18.Text = "Added newline (different session): " + (ID.Count - 1).ToString() + " times";
                        }
                        catch (Exception ex)
                        {
                            label3.Text = ex.Message.ToString();
                        }
                    }

                    //-----------------------------------------FOOTNOTES.XML---------------------------------------------------

                    if (footNotes)
                    {

                        tabPage5.ShowPageInTabControl(tabControl1);
                        try
                        {
                            pTo = 0;
                            extractedFootNotes = @"C:\extracted\" + extractedFootNotes;
                            extractedTxt = System.IO.File.ReadAllText(extractedFootNotes);

                            pFrom = extractedTxt.IndexOf("<w:t", pTo) + "<w:t".Length;
                            if (extractedTxt.IndexOf("<w:t", pTo) == -1)
                                textBox3.Text = "Empty Foot Notes";
                            else
                            {
                                pFrom = extractedTxt.IndexOf(">", pFrom) + 1;
                                pTo = extractedTxt.IndexOf("</w:t>", pFrom);
                                textBox3.Text = extractedTxt.Substring(pFrom, pTo - pFrom);
                            }
                        }
                        catch (Exception ex)
                        {
                            textBox3.Text = ex.Message.ToString();
                        }
                    }

                    //-----------------------------------------ENDNOTES.XML---------------------------------------------------

                    if (endNotes)
                    {
                        tabPage6.ShowPageInTabControl(tabControl1);
                        try
                        {
                            pTo = 0;
                            extractedEndNotes = @"C:\extracted\" + extractedEndNotes;
                            extractedTxt = System.IO.File.ReadAllText(extractedEndNotes);

                            pFrom = extractedTxt.IndexOf("<w:t", pTo) + "<w:t".Length;
                            if (extractedTxt.IndexOf("<w:t", pTo) == -1)
                                textBox2.Text = "Empty End Notes";
                            else
                            {
                                pFrom = extractedTxt.IndexOf(">", pFrom) + 1;
                                pTo = extractedTxt.IndexOf("</w:t>", pFrom);
                                textBox2.Text = extractedTxt.Substring(pFrom, pTo - pFrom);
                            }
                        }
                        catch (Exception ex)
                        {
                            textBox2.Text = ex.Message.ToString();
                        }
                    }

                }
            }
            tabPage7.HidePage();
            tabPage7.ShowPageInTabControl(tabControl1);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.ScrollBars = ScrollBars.Vertical;
        }

        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("docx forensics created By Hasan 12/25/2015 for JUST");
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.SelectAll();
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Copy();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Cut();
        }

        private void selectNoneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.DeselectAll();
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Paste();
        }

        private void lastFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string lineB = "";
            string line = "";
            int countt = 0;
            int pFrom = 0, pTo = 0, pCheck = 0;
            int lFrom = 0, lTo = 0, lCheck = 0;
            string extractedTxtPath = "", extractedPropsCore = "", extractedPropsApp = "", extractedContents = "", extractedDocument = "", extractedFootNotes = "", extractedEndNotes = "";
            bool propsCore = false, propsApp = false, footNotes = false, endNotes = false;
            string extractedTxt = "";
            List<String> ID = new List<string>();





            if (tabPage3.IsVisible())
                tabPage3.HidePage();
            if (tabPage4.IsVisible())
                tabPage4.HidePage();
            if (tabPage5.IsVisible())
                tabPage5.HidePage();
            if (tabPage6.IsVisible())
                tabPage6.HidePage();
            

            String result = "";



            try
            {


                extractedContents = @"C:\extracted\[Content_Types].xml";

                extractedTxt = System.IO.File.ReadAllText(extractedContents);
                pTo = 0;
                int count = 0;
            ploop:
                pFrom = extractedTxt.IndexOf("PartName=\"", pTo) + "PartName=\"".Length;
                pFrom = extractedTxt.IndexOf("/", pFrom) + 1;
                pTo = extractedTxt.IndexOf(".xml", pFrom) + ".xml".Length;
                pCheck = extractedTxt.IndexOf("PartName=\"", pTo);

                if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains("core.xml"))
                {
                    extractedPropsCore = extractedTxt.Substring(pFrom, pTo - pFrom);
                    propsCore = true;
                }
                if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains("app.xml"))
                {
                    extractedPropsApp = extractedTxt.Substring(pFrom, pTo - pFrom);
                    propsApp = true;
                }

                if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains("document.xml"))
                {
                    extractedDocument = extractedTxt.Substring(pFrom, pTo - pFrom);
                    propsApp = true;
                }

                if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains("footnotes.xml"))
                {
                    extractedFootNotes = extractedTxt.Substring(pFrom, pTo - pFrom);
                    footNotes = true;
                }
                if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains("endnotes.xml"))
                {
                    extractedEndNotes = extractedTxt.Substring(pFrom, pTo - pFrom);
                    endNotes = true;
                }
                result += ++count + "- " + extractedTxt.Substring(pFrom, pTo - pFrom) + "\n\n";
                if (pCheck > 0)
                    goto ploop;


                label2.Text = result.ToString();
            }
            catch (Exception ex)
            {
                label2.Text = ex.Message.ToString();
            }

            //-----------------------------------------DOCUMENT.XML---------------------------------------------------

            //try
            {
                int tCheck = 0;
                result = "";
                extractedTxtPath = @"C:\extracted\" + extractedDocument;
                string extractedTxt1;
                extractedTxt1 = System.IO.File.ReadAllText(extractedTxtPath);
                int count = 0;
                lCheck = 1;
                lTo = 0;
                textWithTags = extractedTxt1;
                List<String> edits = new List<string>();
                //String[] edits = new String[100];
                while (lCheck > 0)
                {
                    if (countt != 0)
                        textBox1.AppendText(Environment.NewLine);
                    countt++;
                    count = 0;
                    /*if ((extractedTxt1.IndexOf("<w:tbl", lTo)) >= (extractedTxt1.IndexOf("</w:tbl>", lTo)) && extractedTxt1.IndexOf("</w:tbl>", lTo) != -1)
                    {
                        break;
                    }*/

                    if ((extractedTxt1.IndexOf("<w:p w", lTo)) >= (extractedTxt1.IndexOf("<w:tbl", lTo)) && extractedTxt1.IndexOf("<w:tbl", lTo) != -1)
                    {

                        tCheck = 1;
                        lFrom = extractedTxt1.IndexOf("<w:tbl", lTo) + "<w:tbl".Length;
                        lFrom = extractedTxt1.IndexOf(">", lFrom) + 1;
                        lTo = extractedTxt1.IndexOf("</w:tbl>", lFrom);
                        lTo += "</w:tbl>".Length;
                        lCheck = extractedTxt1.IndexOf("<w:p w", lTo);
                        extractedTxt = extractedTxt1.Substring(lFrom, lTo - lFrom);
                    }
                    else
                    {
                        tCheck = 0;
                        lFrom = extractedTxt1.IndexOf("<w:p w", lTo) + "<w:p w".Length;
                        line = extractedTxt1.Substring(extractedTxt1.IndexOf("w:rsidR=\"", lTo) + "w:rsidR=\"".Length, extractedTxt1.IndexOf("\"", lFrom) - lFrom);
                        lFrom = extractedTxt1.IndexOf(">", lFrom) + 1;
                        lTo = extractedTxt1.IndexOf("</w:p>", lFrom);
                        if (lTo < 0)
                            break;
                        lCheck = extractedTxt1.IndexOf("<w:p w", lTo);
                        extractedTxt = extractedTxt1.Substring(lFrom, lTo - lFrom);
                        /*textBox1.AppendText("----------------------------------------------------");
                        textBox1.AppendText(Environment.NewLine);
                        textBox1.AppendText(extractedTxt);
                        textBox1.AppendText(Environment.NewLine);
                        textBox1.AppendText("----------------------------------------------------");
                        textBox1.AppendText(Environment.NewLine);
                        textBox1.AppendText(lCheck.ToString());
                        textBox1.AppendText(Environment.NewLine);*/
                        //textBox1.AppendText(line);
                        if (!ID.Contains(line))
                        {
                            ID.Add(line);
                        }
                        if (ID.IndexOf(line) == 0)
                            textBox1.AppendText("ID: " + 1 + " "); //for disable first line showing ID 0
                        else
                            textBox1.AppendText("ID: " + ID.IndexOf(line) + " ");
                        textBox1.AppendText("LINE " + countt.ToString() + ":  ");
                    }
                    edits.Clear();
                    pTo = 0;
                    lineBlocksWithTags.Add(extractedTxt);
                    lineB = "";
                ploop1:
                    count++;
                    if (tCheck == 1)
                    {
                        textBox1.AppendText(Environment.NewLine);
                        textBox1.AppendText("**TABLE: ");
                        pFrom = extractedTxt.IndexOf("<w:t>", pTo) + "<w:t>".Length;
                        pTo = extractedTxt.IndexOf("</w:t>", pFrom);
                        if (pTo < 0)
                        {
                            lineBlocks.Add(lineB);
                            continue;
                        }
                        pCheck = extractedTxt.IndexOf("<w:t", pTo);

                    }
                    else
                    {
                    anotherloop:
                        if (!extractedTxt.Contains("<w:t>") && !extractedTxt.Contains("<w:t xml:space=\"preserve\">"))
                        {
                            lineBlocks.Add(lineB);
                            continue;
                        }
                        if (extractedTxt.Contains("<w:textAlignment"))
                        {
                            lineBlocks.Add(lineB);
                            continue;
                        }
                        pFrom = extractedTxt.IndexOf("<w:t", pTo) + "<w:t".Length;
                        if (extractedTxt.Substring(pFrom - 4, 10).Contains("<w:tab"))
                        {
                            pFrom = extractedTxt.IndexOf(">", pFrom) + 1;
                            extractedTxt = extractedTxt.Substring(pFrom, extractedTxt.Length - pFrom);
                            pTo = 0;
                            goto anotherloop;
                        }
                        pFrom = extractedTxt.IndexOf(">", pFrom) + 1;
                        pTo = extractedTxt.IndexOf("</w:t>", pFrom);
                        if (pTo < 0)
                        {
                            lineBlocks.Add(lineB);
                            continue;
                        }
                        pCheck = extractedTxt.IndexOf("<w:t", pTo);
                    }

                    edits.Add(extractedTxt.Substring(pFrom, pTo - pFrom));
                    //edits[count-1] = extractedTxt.Substring(pFrom, pTo - pFrom);
                    blocks.Add(extractedTxt.Substring(pFrom, pTo - pFrom));
                    lineB += extractedTxt.Substring(pFrom, pTo - pFrom);
                    textBox1.AppendText(extractedTxt.Substring(pFrom, pTo - pFrom));



                    if (pCheck > 0)
                        goto ploop1;
                    lineBlocks.Add(lineB);
                    if (tCheck == 1 && count > 1)
                    {
                        textBox1.AppendText(Environment.NewLine);
                        textBox1.AppendText("**The above table has " + count.ToString() + " cells");
                    }
                    else if (count > 1)
                    {
                        textBox1.AppendText(Environment.NewLine);
                        textBox1.AppendText("**Something added to the above line " + (count - 1).ToString() + " Times");
                        if (extractedTxt.Contains(line))
                        {
                            pFrom = extractedTxt.IndexOf(line, 0) + line.Length;
                            pFrom = extractedTxt.IndexOf("<w:t", pFrom) + "<w:t".Length;
                            pFrom = extractedTxt.IndexOf(">", pFrom) + 1;
                            pTo = extractedTxt.IndexOf("</w:t>", pFrom);
                            textBox1.AppendText("   |  Original was: \"" + extractedTxt.Substring(pFrom, pTo - pFrom) + "\"");
                            textBox1.AppendText(" Added: ");
                            for (int i = 0; i < count; i++)
                            {
                                if (edits[i] == extractedTxt.Substring(pFrom, pTo - pFrom))
                                    continue;
                                textBox1.AppendText("\"" + edits[i].ToString() + "\" ");
                            }
                        }
                        else
                        {
                            textBox1.AppendText("   |  Original was: \"" + edits[0].ToString() + "\"");
                            textBox1.AppendText(" Added: ");
                            for (int i = 1; i < count; i++)
                            {
                                textBox1.AppendText("\"" + edits[i].ToString() + "\" ");
                            }
                        }
                    }


                }


            }
            /* catch (Exception ex)
             {
                 textBox1.Text = ex.Message.ToString();
             }*/


            //-----------------------------------------CORE.XML---------------------------------------------------

            if (propsCore)
            {
                tabPage3.ShowPageInTabControl(tabControl1);
                string title = "", created = "", modified = "", keywords = "", authors = "", modifiedDate = "", createdDate = "";
                string modifiedTimes = "";
                try
                {

                    result = "";
                    extractedPropsCore = @"C:\extracted\" + extractedPropsCore;

                    extractedTxt = System.IO.File.ReadAllText(extractedPropsCore);

                    pFrom = extractedTxt.IndexOf("<dc:title>");
                    if (pFrom > 0)
                    {
                        pFrom = extractedTxt.IndexOf("<dc:title>", 0) + "<dc:title>".Length;
                        pTo = extractedTxt.IndexOf("</dc:title>", pFrom);

                        title = extractedTxt.Substring(pFrom, pTo - pFrom);
                    }

                    pFrom = extractedTxt.IndexOf("<cp:keywords>");
                    if (pFrom > 0)
                    {
                        pFrom = extractedTxt.IndexOf("<cp:keywords>", 0) + "<cp:keywords>".Length;
                        pTo = extractedTxt.IndexOf("</cp:keywords>", pFrom);

                        keywords = extractedTxt.Substring(pFrom, pTo - pFrom);
                    }

                    pFrom = extractedTxt.IndexOf("<cp:revision>");
                    if (pFrom > 0)
                    {
                        pFrom = extractedTxt.IndexOf("<cp:revision>", 0) + "<cp:revision>".Length;
                        pTo = extractedTxt.IndexOf("</cp:revision>", pFrom);

                        int revision = Convert.ToInt32(extractedTxt.Substring(pFrom, pTo - pFrom));
                        revision = revision / 2;

                        modifiedTimes = "modofied " + revision + " times";
                    }

                    pFrom = extractedTxt.IndexOf("<cp:lastModifiedBy>");
                    if (pFrom > 0)
                    {
                        pFrom = extractedTxt.IndexOf("<cp:lastModifiedBy>", 0) + "<cp:lastModifiedBy>".Length;
                        pTo = extractedTxt.IndexOf("</cp:lastModifiedBy>", pFrom);

                        modified = extractedTxt.Substring(pFrom, pTo - pFrom);
                    }

                    pFrom = extractedTxt.IndexOf("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">");
                    if (pFrom > 0)
                    {
                        pFrom = extractedTxt.IndexOf("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">", 0) + "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">".Length;
                        pTo = extractedTxt.IndexOf("</dcterms:modified>", pFrom);

                        modifiedDate = extractedTxt.Substring(pFrom, pTo - pFrom - 1);
                        modifiedDate = modifiedDate.Replace('T', ' ');
                    }

                    pFrom = extractedTxt.IndexOf("<dcterms:created xsi:type=\"dcterms:W3CDTF\">");
                    if (pFrom > 0)
                    {
                        pFrom = extractedTxt.IndexOf("<dcterms:created xsi:type=\"dcterms:W3CDTF\">", 0) + "<dcterms:created xsi:type=\"dcterms:W3CDTF\">".Length;
                        pTo = extractedTxt.IndexOf("</dcterms:created>", pFrom);

                        createdDate = extractedTxt.Substring(pFrom, pTo - pFrom - 1);
                        createdDate = createdDate.Replace('T', ' ');
                    }

                    pFrom = extractedTxt.IndexOf("<dc:creator>");
                    if (pFrom > 0)
                    {
                        pFrom = extractedTxt.IndexOf("<dc:creator>", 0) + "<dc:creator>".Length;
                        pTo = extractedTxt.IndexOf("</dc:creator>", pFrom);
                        if (extractedTxt.Substring(pFrom, pTo - pFrom).Contains(";"))
                        {
                            string txt = extractedTxt.Substring(pFrom, pTo - pFrom);
                            string[] txtResult;
                            txtResult = txt.Split(';');

                            foreach (string s in txtResult)
                            {
                                authors += s + " , ";
                            }
                            created = txtResult[0];
                        }
                        else
                            created = extractedTxt.Substring(pFrom, pTo - pFrom);
                    }

                    label3.Text = "Title: " + title;
                    label4.Text = "Tags: " + keywords;
                    label5.Text = "Authors: " + authors;
                    label6.Text = "Creator: created by " + created + " on " + createdDate;
                    label7.Text = "Modifier: modified by " + modified + " on " + modifiedDate;
                    label8.Text = "Modify: " + modifiedTimes;
                }
                catch (Exception ex)
                {
                    label3.Text = ex.Message.ToString();
                }

            }

            //-----------------------------------------APP.XML---------------------------------------------------

            if (propsApp)
            {
                tabPage4.ShowPageInTabControl(tabControl1);

                string characters = "", charactersWithSpaces = "", words = "", paragraphs = "";
                try
                {

                    result = "";
                    extractedPropsApp = @"C:\extracted\" + extractedPropsApp;

                    extractedTxt = System.IO.File.ReadAllText(extractedPropsApp);

                    pFrom = extractedTxt.IndexOf("</Words>");
                    if (pFrom > 0)
                    {
                        pFrom = extractedTxt.IndexOf("<Words>", 0) + "<Words>".Length;
                        pTo = extractedTxt.IndexOf("</Words>", pFrom);

                        words = extractedTxt.Substring(pFrom, pTo - pFrom);
                    }

                    pFrom = extractedTxt.IndexOf("<Characters>");
                    if (pFrom > 0)
                    {
                        pFrom = extractedTxt.IndexOf("<Characters>", 0) + "<Characters>".Length;
                        pTo = extractedTxt.IndexOf("</Characters>", pFrom);

                        characters = extractedTxt.Substring(pFrom, pTo - pFrom);
                    }

                    pFrom = extractedTxt.IndexOf("<CharactersWithSpaces>");
                    if (pFrom > 0)
                    {
                        pFrom = extractedTxt.IndexOf("<CharactersWithSpaces>", 0) + "<CharactersWithSpaces>".Length;
                        pTo = extractedTxt.IndexOf("</CharactersWithSpaces>", pFrom);

                        charactersWithSpaces = extractedTxt.Substring(pFrom, pTo - pFrom);
                    }

                    pFrom = extractedTxt.IndexOf("<Paragraphs>");
                    if (pFrom > 0)
                    {
                        pFrom = extractedTxt.IndexOf("<Paragraphs>", 0) + "<Paragraphs>".Length;
                        pTo = extractedTxt.IndexOf("</Paragraphs>", pFrom);

                        paragraphs = extractedTxt.Substring(pFrom, pTo - pFrom);
                    }

                    label13.Text = "Paragraphs: " + paragraphs;
                    label14.Text = "Words: " + words;
                    label15.Text = "Characters: " + characters;
                    label16.Text = "Characters With Spaces: " + charactersWithSpaces;
                    label17.Text = "Lines: " + countt;
                    label18.Text = "Added newline (different session): " + (ID.Count - 1).ToString() + " times";
                }
                catch (Exception ex)
                {
                    label3.Text = ex.Message.ToString();
                }
            }

            //-----------------------------------------FOOTNOTES.XML---------------------------------------------------

            if (footNotes)
            {

                tabPage5.ShowPageInTabControl(tabControl1);
                try
                {
                    pTo = 0;
                    extractedFootNotes = @"C:\extracted\" + extractedFootNotes;
                    extractedTxt = System.IO.File.ReadAllText(extractedFootNotes);

                    pFrom = extractedTxt.IndexOf("<w:t", pTo) + "<w:t".Length;
                    if (extractedTxt.IndexOf("<w:t", pTo) == -1)
                        textBox3.Text = "Empty Foot Notes";
                    else
                    {
                        pFrom = extractedTxt.IndexOf(">", pFrom) + 1;
                        pTo = extractedTxt.IndexOf("</w:t>", pFrom);
                        textBox3.Text = extractedTxt.Substring(pFrom, pTo - pFrom);
                    }
                }
                catch (Exception ex)
                {
                    textBox3.Text = ex.Message.ToString();
                }
            }

            //-----------------------------------------ENDNOTES.XML---------------------------------------------------

            if (endNotes)
            {
                tabPage6.ShowPageInTabControl(tabControl1);
                try
                {
                    pTo = 0;
                    extractedEndNotes = @"C:\extracted\" + extractedEndNotes;
                    extractedTxt = System.IO.File.ReadAllText(extractedEndNotes);

                    pFrom = extractedTxt.IndexOf("<w:t", pTo) + "<w:t".Length;
                    if (extractedTxt.IndexOf("<w:t", pTo) == -1)
                        textBox2.Text = "Empty End Notes";
                    else
                    {
                        pFrom = extractedTxt.IndexOf(">", pFrom) + 1;
                        pTo = extractedTxt.IndexOf("</w:t>", pFrom);
                        textBox2.Text = extractedTxt.Substring(pFrom, pTo - pFrom);
                    }
                }
                catch (Exception ex)
                {
                    textBox2.Text = ex.Message.ToString();
                }
            }
            tabPage7.HidePage();
            tabPage7.ShowPageInTabControl(tabControl1);
        }



        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabPage3.IsVisible())
                tabPage3.HidePage();
            if (tabPage4.IsVisible())
                tabPage4.HidePage();
            if (tabPage5.IsVisible())
                tabPage5.HidePage();
            if (tabPage6.IsVisible())
                tabPage6.HidePage();
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            label2.Text = "";
            blocks.Clear();
            lineBlocks.Clear();
            lineBlocksWithTags.Clear();
            textWithTags = "";
        }

        private void closeAndClearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabPage3.IsVisible())
                tabPage3.HidePage();
            if (tabPage4.IsVisible())
                tabPage4.HidePage();
            if (tabPage5.IsVisible())
                tabPage5.HidePage();
            if (tabPage6.IsVisible())
                tabPage6.HidePage();
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            label2.Text = "";
            blocks.Clear();
            lineBlocks.Clear();
            lineBlocksWithTags.Clear();
            textWithTags = "";
            string extractPath = @"C:\extracted";
            System.IO.DirectoryInfo directory = new System.IO.DirectoryInfo(extractPath);
            foreach (System.IO.FileInfo file in directory.GetFiles()) file.Delete();
            foreach (System.IO.DirectoryInfo subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
        }

        private void saveAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter sw = new StreamWriter(saveFileDialog1.FileName))
                    sw.WriteLine(textBox1.Text);
            }
        }

        private void saveByBlockToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog sf = new FolderBrowserDialog();
            if (sf.ShowDialog() == DialogResult.OK)
            {
                int count = 0;
                string savePath = sf.SelectedPath;
                Directory.CreateDirectory(savePath + "\\Blocks");
                foreach (string x in blocks)
                {
                    string c = x.Replace(" ", "");
                    if (c.Length > 2)
                        File.WriteAllText(savePath + "\\Blocks\\block" + count++ + ".txt", x);
                }
            }
        }

        private void saveByLineincludeTagsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog sf = new FolderBrowserDialog();
            if (sf.ShowDialog() == DialogResult.OK)
            {
                int count = 0;
                string savePath = sf.SelectedPath;
                Directory.CreateDirectory(savePath + "\\Lines_with_tags");
                foreach (string x in lineBlocksWithTags)
                {
                    string c = x.Replace(" ","");
                    if (c.Length > 2)
                        File.WriteAllText(savePath + "\\Lines_with_tags\\block" + count++ + ".txt", x);
                }
            }
        }

        private void saveAllincludeTagsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter sw = new StreamWriter(saveFileDialog1.FileName))
                    sw.WriteLine(textWithTags);
            }
        }

        private void saveByLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog sf = new FolderBrowserDialog();
            if (sf.ShowDialog() == DialogResult.OK)
            {
                int count = 0;
                string savePath = sf.SelectedPath;
                Directory.CreateDirectory(savePath + "\\Lines");
                foreach (string x in lineBlocks)
                {
                    string c = x.Replace(" ", "");
                    if (c.Length > 2)
                        File.WriteAllText(savePath + "\\Lines\\block" + count++ + ".txt", x);
                }
            }
        }

        private void chooseFolerToSearchToolStripMenuItem_Click(object sender, EventArgs e)
        {
                int count = 0;
                int countt = 0;
                MessageBox.Show("Select RAM dump file then select blocks folder");
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.InitialDirectory = @"C:\";
                openFileDialog1.Title = "Browse RAM Dump";
                openFileDialog1.CheckFileExists = true;
                openFileDialog1.CheckPathExists = true;



                openFileDialog1.DefaultExt = "DMP";

                openFileDialog1.Filter = "DMP files (*.dmp)|*.dmp|All files (*.*)|*.*";

                openFileDialog1.FilterIndex = 1;

                openFileDialog1.RestoreDirectory = true;



                openFileDialog1.ReadOnlyChecked = true;

                openFileDialog1.ShowReadOnly = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)

                {
                    FolderBrowserDialog fd = new FolderBrowserDialog();
                    fd.Description = "Select blocks folder";
                    if (fd.ShowDialog() == DialogResult.OK)
                    {

                        //string dmp = File.ReadAllText(openFileDialog1.FileName);
                        byte[] bytes = File.ReadAllBytes(openFileDialog1.FileName);
                        string dmp = System.Text.Encoding.UTF8.GetString(bytes);
                        string path = fd.SelectedPath;
                        string[] blocks = Directory.GetFiles(path, "*.txt");
                        foreach (string block in blocks)
                        {
                            try
                            {
                                string x = File.ReadAllText(block);
                                textBox4.AppendText(block + "...");
                                if (dmp.Contains(x))
                                {
                                    count++;
                                    textBox4.AppendText("Found (" + Regex.Matches(dmp, x).Count + ") times");
                                }
                                else
                                 textBox4.AppendText("Not Found");
                                textBox4.AppendText(Environment.NewLine);
                                countt++;
                                Application.DoEvents();
                            }
                            catch (Exception )
                            {
                                textBox4.AppendText("Not Found");
                                textBox4.AppendText(Environment.NewLine);
                                countt++;
                                Application.DoEvents();
                            }
                        }
                    }
                }
                double avg = (double)count / (double)countt;
                avg *= 100;
                textBox4.AppendText("-----FINISHED Found " + count + " out of " + countt + " (%" + Math.Round(avg, 2) + ")");
                textBox4.AppendText(Environment.NewLine);
            
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            textBox4.ScrollBars = ScrollBars.Vertical;
        }

        private void chooseFileToSearchToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try {
                MessageBox.Show("Select RAM dump file then select blocks folder");
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.InitialDirectory = @"C:\";
                openFileDialog1.Title = "Browse RAM Dump";
                openFileDialog1.CheckFileExists = true;
                openFileDialog1.CheckPathExists = true;



                openFileDialog1.DefaultExt = "DMP";

                openFileDialog1.Filter = "DMP files (*.dmp)|*.dmp|All files (*.*)|*.*";

                openFileDialog1.FilterIndex = 1;

                openFileDialog1.RestoreDirectory = true;



                openFileDialog1.ReadOnlyChecked = true;

                openFileDialog1.ShowReadOnly = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)

                {
                    OpenFileDialog fd = new OpenFileDialog();
                    fd.Title = "Select file for search";
                    fd.DefaultExt = "txt";

                    fd.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    if (fd.ShowDialog() == DialogResult.OK)
                    {
                        //string dmp = File.ReadAllText(openFileDialog1.FileName);
                        byte[] bytes = File.ReadAllBytes(openFileDialog1.FileName);
                        string dmp = System.Text.Encoding.UTF8.GetString(bytes);
                        string x = File.ReadAllText(fd.FileName);
                        textBox4.AppendText(fd.FileName + "...");
                        if (dmp.Contains(x))
                            textBox4.AppendText("Found (" + Regex.Matches(dmp, x).Count + ") times");
                        else
                            textBox4.AppendText("Not Found");
                        textBox4.AppendText(Environment.NewLine);
                    }
                }
                textBox4.AppendText("-----FINISHED");
                textBox4.AppendText(Environment.NewLine);
            }
            catch (Exception ex)
            {
                textBox4.AppendText(Environment.NewLine);
                textBox4.AppendText(ex.Message.ToString());
            }
        }

        private void selectAllToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            textBox4.SelectAll();
        }

        private void selectNoneToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            textBox4.DeselectAll();
        }

        private void clearToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            textBox4.Clear();
        }

        private void copyToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            textBox4.Copy();
        }

        private void cutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            textBox4.Cut();
        }

        private void saveToFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter sw = new StreamWriter(saveFileDialog1.FileName))
                    sw.WriteLine(textBox4.Text);
            }
        }
    }
}
