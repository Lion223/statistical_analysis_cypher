using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;

namespace DIT
{
    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        #region variables
        // 1 - src text, 2-3 - PT (plain text), ET (encrypted text) in letter combinations, original - copy of src text
        string tempFile, tempFile2, tempFile3, original;

        // pair "symbol(s) - frequency" for the construction of n-grams, 4-5 - Vigenere monograms
        Dictionary<string, int> symCount = new Dictionary<string, int>();
        Dictionary<string, int> symCount2 = new Dictionary<string, int>();
        Dictionary<string, int> symCount3 = new Dictionary<string, int>();
        Dictionary<string, int> symCount4 = new Dictionary<string, int>();
        Dictionary<string, int> symCount5 = new Dictionary<string, int>();

        // alphabet based on src text
        List<char> alphabet = new List<char>();

        // alphabet for Caesar cipher
        List<char> cAlphabet = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', ' ', '.', ',', ';', '-', '\'' };
        List<char> copy;

        // alphabet for Hill cipher. Alphabet power - simple number for ability to get inverted matrix
        List<char> hAlphabet = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', ' ', '.', ',', ';', '-' };

        // alphabet for Feistel cipher
        List<char> fAlphabet = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', ' ', '.', ',', ';', '-', '\'' };

        // alphabet for RSA
        List<char> rAlphabet = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', ' ', '.', ',', ';', '-', '\'' };

        // alphabet for simple substitution cipher
        List<char> sAlphabet = new List<char>();
        Dictionary<string, int> topFifteen = new Dictionary<string, int>();

        // alphabet-key for simple substitution cipher
        List<char> sKey = new List<char>();

        // Calphabet copy for Vigenere cipher
        List<char> vAlphabet = new List<char>();

        // text index (word search)
        int startIndex = 0;

        // text index (word search, word's first symbol)
        int index = 0;

        // counter (word search)
        int countByWord = 0;

        // counter (word search by letter count)
        int countByCount = 0;

        // word list (search by letter count)
        List<string> wordsByCount = new List<string>();

        // check for space in symbol set
        bool firstSp, secondSp, thirdSp = false;

        // RSA results and coefficients
        BigInteger[] rsaEncoded = null;
        int p, q, e, d, n, euler;
        string numericalEncoded = string.Empty;

        // ET DES
        byte[] encrypted;
        #endregion

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            metroTabControl1.SelectedIndex = -1;
            metroTabControl1.TabPages.Remove(metroTabPage1);
            metroTabControl1.TabPages.Remove(metroTabPage2);
            metroTabControl1.TabPages.Remove(metroTabPage3);
            metroTabControl1.TabPages.Remove(metroTabPage4);
            metroTabControl1.TabPages.Remove(metroTabPage5);
            metroTabControl1.TabPages.Remove(metroTabPage6);
            metroTabControl1.TabPages.Remove(metroTabPage7);
            metroTabControl1.TabPages.Remove(metroTabPage8);
            metroTabControl1.TabPages.Remove(metroTabPage9);
            metroTabControl1.TabPages.Remove(metroTabPage10);
            metroTabControl1.TabPages.Remove(metroTabPage11);
            metroTabControl1.TabPages.Remove(metroTabPage12);
            metroTabControl1.TabPages.Remove(metroTabPage13);

            // main
            metroTabControl1.TabPages.Add(metroTabPage1);

            // monograms
            metroTabControl1.TabPages.Add(metroTabPage2);

            // bigrams
            metroTabControl1.TabPages.Add(metroTabPage3);

            // trigrams
            metroTabControl1.TabPages.Add(metroTabPage4);

            // Caesar
            metroTabControl1.TabPages.Add(metroTabPage5);

            // SSC
            metroTabControl1.TabPages.Add(metroTabPage6);

            // Vigener
            metroTabControl1.TabPages.Add(metroTabPage7);

            // Hill
            metroTabControl1.TabPages.Add(metroTabPage10);

            // Feistel
            metroTabControl1.TabPages.Add(metroTabPage11);

            // RSA
            metroTabControl1.TabPages.Add(metroTabPage12);

            // DES
            metroTabControl1.TabPages.Add(metroTabPage13);

            // comparative monograms
            metroTabControl1.TabPages.Add(metroTabPage8);

            // letter combinations
            metroTabControl1.TabPages.Add(metroTabPage9);
        }

        // load data
        private void fileBtn_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                fileLbl.Text = ofd.SafeFileName;
                fileTb.Text = File.ReadAllText(ofd.FileName);
                string temp = String.Join("", fileTb.Text.Where(c => c != '\n' && c != '\r' && c != '\t'));
                fileTb.Text = temp;
                original = fileTb.Text;
                fileTb_C.Text = fileTb.Text;
                fileTb_S.Text = fileTb.Text;
                LoadData();
            }
        }

        // reset, load alphabets and data for n-grams
        private void LoadData()
        {
            Clear();
            LoadAlphabet();
            metroTabControl1.SelectedIndex = 0;
        }

        // reset
        private void Clear()
        {
            alphabet.Clear();
            sAlphabet.Clear();
            sKey.Clear();
            CkeyLbl.Text = string.Empty;
            fromTb.Text = string.Empty;
            toTb.Text = string.Empty;
            SalphaLbl.Text = string.Empty;
            SkeyLbl.Text = string.Empty;
            symCount.Clear();
            symCount2.Clear();
            symCount3.Clear();
            symCount4.Clear();
            symCount5.Clear();
            symDgv.Rows.Clear();
            symDgv2.Rows.Clear();
            symDgv3.Rows.Clear();
            symDgv4.Rows.Clear();
            freqChart.Series[0].Points.Clear();
            freqChart2.Series[0].Points.Clear();
            freqChart3.Series[0].Points.Clear();
            freqChart4.Series[0].Points.Clear();
            freqChart4.Series[1].Points.Clear();
            sortCb.SelectedIndex = -1;
            sortCb2.SelectedIndex = -1;
            wordTb.Text = string.Empty;
            countTb.Text = string.Empty;
            openwordTb.Text = string.Empty;
            opencountTb.Text = string.Empty;
            cryptwordTb.Text = string.Empty;
            cryptcountTb.Text = string.Empty;
        }

        private void LoadAlphabet()
        {
            alphabet.Clear();
            foreach (char c in fileTb.Text)
            {
                if (!alphabet.Contains(c) && c != '\r' && c != '\n')
                    alphabet.Add(c);
            }
            alphabet.Sort();

            for (int i = 0; i < alphabet.Count; i++)
            {
                if ((int)alphabet[i] < 65 || ((int)alphabet[i] > 90 && (int)alphabet[i] < 175))
                {
                    char iter = alphabet[i];
                    alphabet.Remove(alphabet[i]);
                    alphabet.Add(iter);
                    i--;
                }
                else
                    break;
            }
            alphaLbl.Text = String.Join("", alphabet);
            alphaLbl.Text = alphaLbl.Text.Replace(' ', '_');
            SalphaLbl.Text = String.Join("", alphabet);
            SalphaLbl.Text = SalphaLbl.Text.Replace(' ', '_');
            sAlphabet = new List<char>(alphabet);
            sKey = new List<char>(alphabet);
            for (int i = 0; i < sKey.Count; i++)
                sKey[i] = '?';
            SkeyLbl.Text = String.Join("", sKey);
            vAlphabet = new List<char>(cAlphabet);
            SalphaDgv.Rows.Clear();
            for (int i = 0; i < SalphaLbl.Text.Length; i++)
                SalphaDgv.Rows.Add(SalphaLbl.Text[i], SkeyLbl.Text[i]);
            copy = new List<char>(cAlphabet);
            int sp = copy.IndexOf(' ');
            copy[sp] = '_';
            CalphaDgv.Rows.Clear();
            for (int i = 0; i < cAlphabet.Count; i++)
                CalphaDgv.Rows.Add(copy[i], '?');
        }

        // reset for src n-gram load
        private void loadngramBtn_Click(object sender, EventArgs e)
        {
            if (fileLbl.Text.Length == 0)
                return;
            if (!fileTb.Text.Equals(original))
                fileLbl.Text = "Changed";
            else
                fileLbl.Text = ofd.SafeFileName;
            symCount.Clear();
            symCount2.Clear();
            symCount3.Clear();
            symDgv.Rows.Clear();
            symDgv2.Rows.Clear();
            symDgv3.Rows.Clear();
            freqChart.Series[0].Points.Clear();
            freqChart2.Series[0].Points.Clear();
            freqChart3.Series[0].Points.Clear();
            LoadAlphabet();
            sortCb.SelectedIndex = -1;
            Ngram(1);
            Ngram(2);
            Ngram(3);
            MessageBox.Show("N-gram analysis has completed");
        }

        // searched word repeat count for n-gram
        private int Search(string word, RichTextBox tb)
        {
            countByWord = 0;
            countByWord += (tb.Text.Length - tb.Text.Replace(word, String.Empty).Length) / word.Length;
            return countByWord;
        }

        private void HighlightWords(RichTextBox tb, string word, bool wordbyamount)
        {
            int s_start = tb.SelectionStart;
            startIndex = 0;
            countByWord = 0;
            index = 0;

            if (!wordbyamount)
            {
                while ((index = tb.Text.IndexOf(word, startIndex)) != -1)
                {
                    tb.Select(index, word.Length);
                    tb.SelectionBackColor = Color.Green;
                    tb.SelectionColor = Color.Black;
                    startIndex = index + word.Length;
                    countByWord++;
                }
                tb.SelectionStart = s_start;
                tb.SelectionLength = 0;
                tb.SelectionBackColor = Color.Black;
                if (tb.Name == "fileTb")
                    wordLbl.Text = countByWord.ToString();
                else if (tb.Name == "openTb")
                    openwordLbl.Text = countByWord.ToString();
                else if (tb.Name == "cryptTb")
                    cryptwordLbl.Text = countByWord.ToString();
            }
            else
            {
                while ((index = tb.Text.IndexOf(word, startIndex)) != -1)
                {
                    // word starts at i and ends at the end of Tb
                    try
                    {
                        // word starts at the start of Tb
                        if (index == 0 && tb.Text[index + word.Length] == ' ')
                        {
                            tb.Select(index, word.Length);
                            tb.SelectionBackColor = Color.Green;
                            tb.SelectionColor = Color.Black;
                            startIndex = index + word.Length;
                            countByCount++;
                        }
                    }
                    catch (IndexOutOfRangeException)
                    {
                        tb.Select(index, word.Length);
                        tb.SelectionBackColor = Color.Green;
                        tb.SelectionColor = Color.Black;
                        startIndex = index + word.Length;
                        countByCount++;
                    }

                    // word starts elsewhere and ends at the end of Tb
                    try
                    {
                        // word starts elsewhere
                        if (index != 0)
                        {
                            if (tb.Text[index - 1] == ' ' && tb.Text[index + word.Length] == ' ')
                            {
                                tb.Select(index, word.Length);
                                tb.SelectionBackColor = Color.Green;
                                tb.SelectionColor = Color.Black;
                                startIndex = index + word.Length;
                                countByCount++;
                            }
                        }
                    }
                    catch (IndexOutOfRangeException)
                    {
                        tb.Select(index, word.Length);
                        tb.SelectionBackColor = Color.Green;
                        tb.SelectionColor = Color.Black;
                        startIndex = index + word.Length;
                        countByCount++;
                    }
                    startIndex = index + word.Length;
                    continue;
                }
                tb.SelectionStart = s_start;
                tb.SelectionLength = 0;
                tb.SelectionBackColor = Color.Black;
                if (tb.Name == "fileTb")
                    countLbl.Text = countByCount.ToString();
                else if (tb.Name == "openTb")
                    opencountLbl.Text = countByCount.ToString();
                else if (tb.Name == "cryptTb")
                    cryptcountLbl.Text = countByCount.ToString();
            }
        }

        // load data for n-grams
        private void Ngram(int n)
        {
            string w;
            string temp;
            int freq;

            switch (n)
            {
                case 1:
                    foreach (char c in alphabet)
                    {
                        if (c == ' ')
                        {
                            char space = '_';
                            symCount.Add(space.ToString(), fileTb.Text.Where(x => x == c).Count());
                        }
                        else
                            symCount.Add(c.ToString(), fileTb.Text.Where(x => x == c).Count());
                    }
                    freqChart.Series[0].Points.DataBindXY(symCount.Keys, symCount.Values);
                    if (freqChart.Series[0].Points.Count > 15)
                        freqChart.ChartAreas[0].AxisX.ScaleView.Size = 15;
                    foreach (KeyValuePair<string, int> item in symCount)
                        symDgv.Rows.Add(item.Key, item.Value);
                    break;
                case 2:
                    for (int i = 0; i < fileTb.Text.Length - 1; i++)
                    {
                        w = fileTb.Text[i].ToString() + fileTb.Text[i + 1].ToString();
                        if (w.Contains("\n") || w.Contains("\r"))
                            continue;
                        freq = Search(w, fileTb);
                        if (w.Contains(" "))
                        {
                            if (w[0] == ' ' && w[1] == ' ')
                            {
                                temp = '_'.ToString() + '_'.ToString();
                                w = temp;
                            }
                            if (w[0] == ' ')
                            {
                                temp = '_'.ToString() + fileTb.Text[i + 1].ToString();
                                w = temp;
                            }
                            else
                            {
                                temp = fileTb.Text[i].ToString() + '_'.ToString();
                                w = temp;
                            }
                        }
                        if (!symCount2.ContainsKey(w))
                            symCount2.Add(w, freq);
                    }
                    topFifteen = (
                        from entry in symCount2
                        orderby entry.Value descending
                        select entry)
                        .Take(15)
                        .ToDictionary(pair => pair.Key, pair => pair.Value);
                    freqChart2.Series[0].Points.DataBindXY(topFifteen.Keys, topFifteen.Values);
                    if (freqChart2.Series[0].Points.Count > 15)
                        freqChart2.ChartAreas[0].AxisX.ScaleView.Size = 15;
                    foreach (KeyValuePair<string, int> item in symCount2)
                        symDgv2.Rows.Add(item.Key, item.Value);
                    break;
                case 3:
                    for (int i = 0; i < fileTb.Text.Length - 2; i++)
                    {
                        temp = string.Empty;
                        w = fileTb.Text[i].ToString() + fileTb.Text[i + 1].ToString() + fileTb.Text[i + 2].ToString();
                        if (w.Contains("\n") || w.Contains("\r"))
                            continue;
                        freq = Search(w, fileTb);
                        if (w.Contains(" "))
                        {
                            if (w[0] == ' ')
                                firstSp = true;
                            if (w[1] == ' ')
                                secondSp = true;
                            if (w[2] == ' ')
                                thirdSp = true;
                            if (firstSp)
                                temp = '_'.ToString();
                            else
                                temp = fileTb.Text[i].ToString();
                            if (secondSp)
                                temp += '_'.ToString();
                            else
                                temp += fileTb.Text[i + 1].ToString();
                            if (thirdSp)
                                temp += '_'.ToString();
                            else
                                temp += fileTb.Text[i + 2].ToString();
                            w = temp;
                            firstSp = false;
                            secondSp = false;
                            thirdSp = false;
                        }
                        if (!symCount3.ContainsKey(w))
                            symCount3.Add(w, freq);
                    }
                    topFifteen = (
                        from entry in symCount3
                        orderby entry.Value descending
                        select entry)
                        .Take(15)
                        .ToDictionary(pair => pair.Key, pair => pair.Value);
                    freqChart3.Series[0].Points.DataBindXY(topFifteen.Keys, topFifteen.Values);
                    if (freqChart3.Series[0].Points.Count > 15)
                        freqChart3.ChartAreas[0].AxisX.ScaleView.Size = 15;
                    foreach (KeyValuePair<string, int> item in symCount3)
                        symDgv3.Rows.Add(item.Key, item.Value);
                    break;
            }
        }

        // load texts for monograms comparison
        private void CompareMonograms(string cipher)
        {
            RichTextBox rb_open = null, rb_crypt = null;
            if (cipher == "vigenere")
            {
                rb_open = fileTb_V;
                rb_crypt = resTb_V;
            }
            if (cipher == "hill")
            {
                rb_open = fileTb_H;
                rb_crypt = resTb_H;
            }
            if (cipher == "feistel")
            {
                rb_open = fileTb_F;
                rb_crypt = resTb_F;
            }
            if (cipher == "rsa")
            {
                rb_open = fileTb_R;
                rb_crypt = resTb_R;
            }
            if (cipher == "des")
            {
                rb_open = fileTb_D;
                rb_crypt = resTb_D;
            }
            CheckMonograms(rb_open, rb_crypt);
        }

        // load monograms for comparison
        private void CheckMonograms(RichTextBox rb_open, RichTextBox rb_crypt)
        {
            symCount4.Clear();
            symCount5.Clear();
            freqChart4.Series[0].Points.Clear();
            freqChart4.Series[1].Points.Clear();
            symDgv4.Rows.Clear();
            string w;
            string temp;
            int freq, freq2;
            for (int i = 0; i < rb_open.Text.Length; i++)
            {
                w = rb_open.Text[i].ToString();
                if (w.Contains("\n") || w.Contains("\r"))
                    continue;
                if (w.Contains(" "))
                {
                    if (!symCount4.ContainsKey("_") && !symCount5.ContainsKey("_"))
                    {
                        freq = Search(w, rb_open);
                        freq2 = Search(w, rb_crypt);
                        temp = "_";
                        w = temp;
                        symCount4.Add(w, freq);
                        symCount5.Add(w, freq2);
                        continue;
                    }
                    else
                        continue;
                }
                if (!symCount4.ContainsKey(w) && !symCount5.ContainsKey(w))
                {
                    freq = Search(w, rb_open);
                    freq2 = Search(w, rb_crypt);

                    symCount4.Add(w, freq);
                    symCount5.Add(w, freq2);
                }
            }
            for (int i = 0; i < rb_crypt.Text.Length; i++)
            {
                w = rb_crypt.Text[i].ToString();
                if (w.Contains("\n") || w.Contains("\r"))
                    continue;
                if (w.Contains(" "))
                {
                    if (!symCount4.ContainsKey("_") && !symCount5.ContainsKey("_"))
                    {
                        freq = Search(w, rb_open);
                        freq2 = Search(w, rb_crypt);
                        temp = "_";
                        w = temp;
                        symCount4.Add(w, freq);
                        symCount5.Add(w, freq2);
                        continue;
                    }
                    else
                        continue;
                }
                if (!symCount4.ContainsKey(w) && !symCount5.ContainsKey(w))
                {
                    freq = Search(w, rb_open);
                    freq2 = Search(w, rb_crypt);

                    symCount4.Add(w, freq);
                    symCount5.Add(w, freq2);
                }
            }
            freqChart4.Series[0].Points.DataBindXY(symCount4.Keys, symCount4.Values);
            freqChart4.Series[1].Points.DataBindXY(symCount5.Keys, symCount5.Values);
            if (freqChart4.Series[0].Points.Count > 15)
                freqChart4.ChartAreas[0].AxisX.ScaleView.Size = 5;
            var target = symCount4.ToList();
            var target2 = symCount5.ToList();
            for (int i = 0; i < symCount4.Count; i++)
                symDgv4.Rows.Add(target[i].Key, target[i].Value, target2[i].Value);

            var ordered_keys = symCount4.OrderByDescending(x => x.Value).ToList();
            var ordered_keys2 = symCount5.OrderByDescending(x => x.Value).ToList();
            List<string> keys = new List<string>();
            for (int i = 0; i < symCount4.Count; i++)
            {
                if (ordered_keys[i].Key != ordered_keys2[i].Key)
                    keys.Add(ordered_keys[i].Key + Environment.NewLine + ordered_keys2[i].Key);
                else
                    keys.Add(ordered_keys[i].Key);
            }

            List<int> ordered_values = symCount4.Values.ToList();
            ordered_values.Sort();
            ordered_values.Reverse();
            freqChart4.Series[2].Points.DataBindXY(keys, ordered_values);
            ordered_values = symCount5.Values.ToList();
            ordered_values.Sort();
            ordered_values.Reverse();
            freqChart4.Series[3].Points.DataBindY(ordered_values);
            freqChart4.Series[2].Enabled = false;
            freqChart4.Series[3].Enabled = false;
        }

        // Caesar decryption
        void Cdecrypt()
        {
            List<char> shiftalp = new List<char>();
            string temp = String.Join("", fileTb_C.Text.Where(c => c != '\n' && c != '\r' && c != '\t'));
            fileTb_C.Text = temp;
            resTb_C.Text = "";
            if (symCount.Count == 0)
                Ngram(1);
            char maxfreqchar = symCount.FirstOrDefault(x => x.Value.Equals(symCount.Values.Max())).Key[0];
            char space = ' ';
            int max = 0, spbar = 0;
            for (int i = 0; i < cAlphabet.Count; i++)
            {
                if (cAlphabet[i] == maxfreqchar)
                    max = i;
                if (cAlphabet[i] == space)
                    spbar = i;
            }
            int key = spbar - max;
            CkeyLbl.Text = key.ToString();
            shiftalp = ShiftRight(copy, key);
            CalphaDgv.Rows.Clear();
            for (int i = 0; i < copy.Count; i++)
                CalphaDgv.Rows.Add(copy[i], shiftalp[i]);

            int step = 0;
            for (int i = 0; i < fileTb_C.Text.Length; i++)
            {
                for (int j = 0; j < cAlphabet.Count; j++)
                {
                    if (fileTb_C.Text[i] == cAlphabet[j])
                    {
                        step = j;
                        break;
                    }
                }

                if ((key + step) > (cAlphabet.Count - 1))
                    resTb_C.Text += cAlphabet[key + step - cAlphabet.Count];
                else
                    resTb_C.Text += cAlphabet[key + step];
            }
        }

        public List<T> ShiftRight<T>(List<T> lst, int shift)
        {
            List<T> result = new List<T>();
            for (int i = 0; i < lst.Count; i++)
            {
                if (i + shift > lst.Count - 1)
                    result.Add(lst[i + shift - lst.Count]);
                else
                    result.Add(lst[i + shift]);
            }
            return result;
        }

        private void CdecyphBtn_Click(object sender, EventArgs e)
        {
            Cdecrypt();
            if (fileTb_V.Text == "PT")
                fileTb_V.Text = resTb_C.Text;
            if (fileTb_H.Text == "Pt")
                fileTb_H.Text = resTb_C.Text;
            if (fileTb_F.Text == "PT")
                fileTb_F.Text = resTb_C.Text;
            if (fileTb_R.Text == "PT")
                fileTb_R.Text = resTb_C.Text;
            if (fileTb_D.Text == "PT")
                fileTb_D.Text = resTb_C.Text;
            MessageBox.Show("Caesar decryption has completed");
        }

        // substitution
        private void replaceBtn_Click(object sender, EventArgs e)
        {
            if (fromTb.Text.Length == toTb.Text.Length)
            {
                for (int i = 0; i < fromTb.Text.Length; i++)
                {
                    for (int j = 0; j < sKey.Count; j++)
                    {
                        if (fromTb.Text[i] == sAlphabet[j])
                        {
                            if (!sKey.Contains(toTb.Text[i]) || toTb.Text[i] == '?')
                                sKey[j] = toTb.Text[i];
                            else
                            {
                                MessageBox.Show("This letter is already in the key");
                                return;
                            }
                        }
                    }
                }
                SkeyLbl.Text = String.Join("", sKey);
                if (SkeyLbl.Text.Contains(" "))
                    SkeyLbl.Text = SkeyLbl.Text.Replace(' ', '_');
                resTb_S.Text = "";
                for (int i = 0; i < fileTb_S.Text.Length; i++)
                {
                    for (int j = 0; j < sAlphabet.Count; j++)
                    {

                        if (fileTb_S.Text[i] == sAlphabet[j])
                        {
                            if (sKey[j] == '?')
                                resTb_S.AppendText(fileTb_S.Text[i].ToString());
                            else
                                resTb_S.AppendText(sKey[j].ToString());
                        }
                    }
                }
                fromTb.Text = string.Empty;
                toTb.Text = string.Empty;
                SalphaDgv.Rows.Clear();
                for (int i = 0; i < SalphaLbl.Text.Length; i++)
                    SalphaDgv.Rows.Add(SalphaLbl.Text[i], SkeyLbl.Text[i]);
                MessageBox.Show("Substituion has completed");
            }
            else
                MessageBox.Show("К-сть символів \"Від\" і \"До\" має бути однаковою.");
        }

        // save the current alphabet-key in a separate file
        private void saveSkeyBtn_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("The previous key will be overwritten. Are you sure?",
                "Save alphabet-key", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string save = string.Join("", sKey.ToArray());
                File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "Skey.txt", save);
                MessageBox.Show("The key is saved to a text file, located in a bin folder");
            }
        }

        // load alphabet-key from the file
        private void loadSkeyBtn_Click(object sender, EventArgs e)
        {
            if (fileLbl.Text.Length == 0)
            {
                MessageBox.Show("No text loaded");
                return;

            }
            if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "Skey.txt"))
            {
                string load = File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + "Skey.txt");
                char[] loadchar = load.ToCharArray();
                sKey = new List<char>(loadchar);
                SkeyLbl.Text = string.Join("", sKey);
                if (SkeyLbl.Text.Contains(" "))
                    SkeyLbl.Text = SkeyLbl.Text.Replace(' ', '_');
                SalphaDgv.Rows.Clear();
                for (int i = 0; i < SalphaLbl.Text.Length; i++)
                    SalphaDgv.Rows.Add(SalphaLbl.Text[i], SkeyLbl.Text[i]);
                MessageBox.Show("The key has been successfully loaded");
            }
            else
            {
                if (ofd2.ShowDialog() == DialogResult.OK)
                {
                    string load = File.ReadAllText(ofd.FileName);
                    char[] loadchar = load.ToCharArray();
                    sKey = new List<char>(loadchar);
                    SkeyLbl.Text = string.Join("", sKey);
                    if (SkeyLbl.Text.Contains(" "))
                        SkeyLbl.Text = SkeyLbl.Text.Replace(' ', '_');
                    SalphaDgv.Rows.Clear();
                    for (int i = 0; i < SalphaLbl.Text.Length; i++)
                        SalphaDgv.Rows.Add(SalphaLbl.Text[i], SkeyLbl.Text[i]);
                    MessageBox.Show("The key has been successfully loaded");
                }
            }
        }

        // reset the current alphabet-key
        private void SclearBtn_Click(object sender, EventArgs e)
        {
            sKey = new List<char>(alphabet);
            for (int i = 0; i < sKey.Count; i++)
                sKey[i] = '?';
            SkeyLbl.Text = String.Join("", sKey);
            resTb_S.Text = "Cypher results";
            fromTb.Text = string.Empty;
            toTb.Text = string.Empty;
            SalphaDgv.Rows.Clear();
            for (int i = 0; i < SalphaLbl.Text.Length; i++)
                SalphaDgv.Rows.Add(SalphaLbl.Text[i], SkeyLbl.Text[i]);
        }

        // Vigenere decryption
        private void VdecyphBtn_Click(object sender, EventArgs e)
        {
            string temp = String.Join("", fileTb_V.Text.Where(c => c != '\n' && c != '\r' && c != '\t'));
            fileTb_V.Text = temp;
            resTb_V.Text = "";
            string key = VkeyTb.Text;
            int j = 0;
            for (int i = 0; i < fileTb_V.Text.Length; i++)
            {
                resTb_V.Text += GetKey(key[j], fileTb_V.Text[i], true);
                j++;
                if (j >= key.Length) j = 0;
            }
            MessageBox.Show("Vigenere decryption has completed");
        }

        // Vigenere cryption
        private void VсyphBtn_Click(object sender, EventArgs e)
        {
            string temp = String.Join("", fileTb_V.Text.Where(c => c != '\n' && c != '\r' && c != '\t'));
            fileTb_V.Text = temp;
            resTb_V.Text = "";
            string key = VkeyTb.Text;
            int j = 0;
            for (int i = 0; i < fileTb_V.Text.Length; i++)
            {
                resTb_V.Text += GetKey(key[j], fileTb_V.Text[i], false);
                j++;
                if (j >= key.Length) j = 0;
            }
            MessageBox.Show("Vigenere cryption has completed");
        }

        // each symbol key definition for Vigenere cypher
        private char GetKey(char key, char x, bool cypher)
        {
            int indexI = 0;
            int indexJ = 0;
            for (int i = 0; i < vAlphabet.Count; i++)
            {
                if (vAlphabet[i] == key) indexI = i;
                if (vAlphabet[i] == x) indexJ = i;
            }
            if (cypher == true)
            {
                int z = indexJ - indexI;
                if (z < 0) z += vAlphabet.Count;
                return vAlphabet[z];
            }
            else
            {
                int z = indexJ + indexI;
                if (z >= vAlphabet.Count) z -= vAlphabet.Count;
                return vAlphabet[z];
            }
        }

        // Hill decryption
        private void HdecyphBtn_Click(object sender, EventArgs e)
        {
            string temp = String.Join("", fileTb_H.Text.Where(c => c != '\n' && c != '\r' && c != '\t'));
            fileTb_H.Text = temp;
            string input = fileTb_H.Text;
            resTb_H.Text = "";
            string key = HkeyTb.Text;
            int[,] keymatrix = KeyMatrix(key);
            int det = DetMatrix(keymatrix);
            if (det == 0 || CommonDivisor(det))
            {
                MessageBox.Show("There is no inverted matrix for this alphabet and key combination");
                return;
            }
            int[] xy = { 1, 1 };
            int x = ExtendedEuclidean(det, hAlphabet.Count, xy);
            int invdet = 0;
            if (det < 0 && x > 0)
                invdet = x;
            else if (det > 0 && x < 0)
                invdet = hAlphabet.Count + x;
            else if (det > 0 && x > 0)
                invdet = x;
            else if (det < 0 && x < 0)
                invdet = -x;
            int[,] adjmatrix = AdjMatrix(keymatrix);
            int[,] invertmatrix = InvertibleMatrix(adjmatrix, invdet);
            while (input.Length % 3 != 0)
            {
                input += " ";
            }
            for (int i = 0; i < input.Length;)
            {
                int[] block = new int[3];
                for (int j = 0; j < 3; j++)
                {
                    block[j] = hAlphabet.IndexOf(input[i]);
                    i++;
                }
                block = MultVectorMatrix(block, invertmatrix);
                for (int j = 0; j < 3; j++)
                    block[j] %= hAlphabet.Count;
                for (int j = 0; j < 3; j++)
                    resTb_H.Text += hAlphabet[block[j]];
            }
            MessageBox.Show("Hill decryption has completed");
        }

        // Hill cryption
        private void HсyphBtn_Click(object sender, EventArgs e)
        {
            string temp = String.Join("", fileTb_H.Text.Where(c => c != '\n' && c != '\r' && c != '\t'));
            fileTb_H.Text = temp;
            string input = fileTb_H.Text;
            while (input.Length % 3 != 0)
            {
                input += " ";
            }
            string key = HkeyTb.Text;
            int[,] keymatrix = KeyMatrix(key);
            int det = DetMatrix(keymatrix);
            if (det == 0 || CommonDivisor(det))
            {
                MessageBox.Show("There is no inverted matrix for this alphabet and key combination");
                return;
            }
            resTb_H.Text = "";
            for (int i = 0; i < input.Length;)
            {
                int[] block = new int[3];
                for (int j = 0; j < 3; j++)
                {
                    block[j] = hAlphabet.IndexOf(input[i]);
                    i++;
                }
                block = MultVectorMatrix(block, keymatrix);
                for (int j = 0; j < 3; j++)
                    block[j] %= hAlphabet.Count;
                for (int j = 0; j < 3; j++)
                    resTb_H.Text += hAlphabet[block[j]];
            }
            MessageBox.Show("Hill cryption has completed");
        }

        // matrix-key initialization
        private int[,] KeyMatrix(string key)
        {
            int[,] result = new int[3, 3];
            int iter = 0;
            List<int> keyindex = new List<int>();
            for (int i = 0; i < key.Length; i++)
                keyindex.Add(hAlphabet.FindIndex(a => a == key[i]));

            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    result[i, j] = keyindex[iter];
                    iter++;
                }
            }
            return result;
        }

        // matrix determinant search
        private int DetMatrix(int[,] matrix)
        {
            if (matrix.Length == 9)
            {
                return matrix[0, 0] * matrix[1, 1] * matrix[2, 2] - matrix[0, 0] * matrix[1, 2] * matrix[2, 1] -
                matrix[0, 1] * matrix[1, 0] * matrix[2, 2] + matrix[0, 1] * matrix[1, 2] * matrix[2, 0] +
                matrix[0, 2] * matrix[1, 0] * matrix[2, 1] - matrix[0, 2] * matrix[1, 1] * matrix[2, 0];
            }
            else if (matrix.Length == 4)
            {
                return matrix[0, 0] * matrix[1, 1] - matrix[1, 0] * matrix[0, 1];
            }
            return 0;
        }

        // extented Euclidean algorithm to find coefficient y
        private int ExtendedEuclidean(int a, int b, int[] x)
        {
            if (a == 0)
            {
                x[0] = 0;
                x[1] = 1;
                return x[0];
            }
            int[] t = { 1, 1 };
            int gcd = ExtendedEuclidean(b % a, a, t);
            x[0] = t[1] - (b / a) * t[0];
            x[1] = t[0];
            return x[0];
        }

        // check for common divider with alphabet power
        private bool CommonDivisor(int det)
        {
            for (int i = 2; i <= hAlphabet.Count / 2; i++)
            {
                if (det % i == 0 && hAlphabet.Count % i == 0)
                    return true;
            }
            return false;
        }

        // finding inverted matrix
        private int[,] InvertibleMatrix(int[,] matrix, int invdet)
        {
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    matrix[i, j] %= hAlphabet.Count;
                }
            }
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    matrix[i, j] *= invdet;
                }
            }
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    matrix[i, j] %= hAlphabet.Count;
                }
            }
            int[,] transmatrix = TransposeMatrix(matrix);
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    if (transmatrix[i, j] < 0)
                        transmatrix[i, j] = hAlphabet.Count + transmatrix[i, j];
                }
            }
            return transmatrix;
        }

        // transposition of the matrix
        private int[,] TransposeMatrix(int[,] matrix)
        {
            int[,] result = new int[3, 3];
            Array.Copy(matrix, result, 9);
            int temp;
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < i; j++)
                {
                    temp = result[i, j];
                    result[i, j] = result[j, i];
                    result[j, i] = temp;
                }
            }
            return result;
        }

        // assembly matrix elements of algebraic additions
        private int[,] AdjMatrix(int[,] matrix)
        {
            int[,] result = new int[3, 3];
            int[,] detminor = new int[2, 2];
            detminor[0, 0] = matrix[1, 1];
            detminor[0, 1] = matrix[1, 2];
            detminor[1, 0] = matrix[2, 1];
            detminor[1, 1] = matrix[2, 2];
            result[0, 0] = DetMatrix(detminor);
            detminor[0, 0] = matrix[1, 0];
            detminor[0, 1] = matrix[1, 2];
            detminor[1, 0] = matrix[2, 0];
            detminor[1, 1] = matrix[2, 2];
            result[0, 1] = DetMatrix(detminor);
            detminor[0, 0] = matrix[1, 0];
            detminor[0, 1] = matrix[1, 1];
            detminor[1, 0] = matrix[2, 0];
            detminor[1, 1] = matrix[2, 1];
            result[0, 2] = DetMatrix(detminor);
            detminor[0, 0] = matrix[0, 1];
            detminor[0, 1] = matrix[0, 2];
            detminor[1, 0] = matrix[2, 1];
            detminor[1, 1] = matrix[2, 2];
            result[1, 0] = DetMatrix(detminor);
            detminor[0, 0] = matrix[0, 0];
            detminor[0, 1] = matrix[0, 2];
            detminor[1, 0] = matrix[2, 0];
            detminor[1, 1] = matrix[2, 2];
            result[1, 1] = DetMatrix(detminor);
            detminor[0, 0] = matrix[0, 0];
            detminor[0, 1] = matrix[0, 1];
            detminor[1, 0] = matrix[2, 0];
            detminor[1, 1] = matrix[2, 1];
            result[1, 2] = DetMatrix(detminor);
            detminor[0, 0] = matrix[0, 1];
            detminor[0, 1] = matrix[0, 2];
            detminor[1, 0] = matrix[1, 1];
            detminor[1, 1] = matrix[1, 2];
            result[2, 0] = DetMatrix(detminor);
            detminor[0, 0] = matrix[0, 0];
            detminor[0, 1] = matrix[0, 2];
            detminor[1, 0] = matrix[1, 0];
            detminor[1, 1] = matrix[1, 2];
            result[2, 1] = DetMatrix(detminor);
            detminor[0, 0] = matrix[0, 0];
            detminor[0, 1] = matrix[0, 1];
            detminor[1, 0] = matrix[1, 0];
            detminor[1, 1] = matrix[1, 1];
            result[2, 2] = DetMatrix(detminor);
            result[0, 1] *= -1;
            result[1, 0] *= -1;
            result[1, 2] *= -1;
            result[2, 1] *= -1;
            return result;
        }

        // vector-matrix multiplication
        static int[] MultVectorMatrix(int[] a, int[,] b)
        {
            int[] result = new int[a.GetLength(0)];
            for (int i = 0; i < a.GetLength(0); i++)
            {
                for (int j = 0; j < b.GetLength(1); j++)
                {
                    result[i] += a[j] * b[j, i];
                }
            }
            return result;
        }

        // Feistel decryption
        private void FdecyphBtn_Click(object sender, EventArgs e)
        {
            FeistelCipher(false);
        }

        // Feistel cryption
        private void FcyphBtn_Click(object sender, EventArgs e)
        {
            FeistelCipher(true);
        }

        // Feistel implementation
        private void FeistelCipher(bool crypt)
        {
            string temp_string = String.Join("", fileTb_F.Text.Where(c => c != '\n' && c != '\r' && c != '\t'));
            fileTb_F.Text = temp_string;
            string input = fileTb_F.Text;
            int K0 = Convert.ToInt16(K0Tb.Text);
            int n = Convert.ToInt16(nTb.Text);
            resTb_F.Text = "";
            while (input.Length % 2 != 0)
                input += " ";
            int[] letterArray = new int[input.Length];
            for (int i = 0; i < input.Length; i++)
                letterArray[i] = fAlphabet.IndexOf(input[i]);
            for (int j = 0; j < input.Length; j += 2)
            {
                int K = crypt ? K0 : (K0downRBtn.Checked ? K0 - n + 1 : K0 + n - 1);
                int L = letterArray[j];
                int R = letterArray[j + 1];
                for (int i = 0; i < n; i++)
                {
                    if (i == n - 1)
                    {
                        int temp = R ^ FeistelFunc(L, K);
                        R = temp;
                    }
                    else
                    {
                        int temp = R ^ FeistelFunc(L, K);
                        R = L;
                        L = temp;
                    }
                    K += crypt ? (K0downRBtn.Checked ? -1 : 1) : (K0downRBtn.Checked ? 1 : -1);
                }
                letterArray[j] = L;
                letterArray[j + 1] = R;
            }
            foreach (var letter in letterArray)
                resTb_F.Text += fAlphabet[letter];
            if (crypt)
                MessageBox.Show("Feistel cryption has completed");
            else MessageBox.Show("Feistel decryption has completed");
        }

        // Feistel function
        private int FeistelFunc(int L, int K)
        {
            return ((L + K) % fAlphabet.Count);
        }

        // RSA decryption
        private void RdecyphBtn_Click(object sender, EventArgs e)
        {
            RSA(false);
        }

        // RSA cryption
        private void RcyphBtn_Click(object sender, EventArgs e)
        {
            RSA(true);
        }

        // RSA implementation
        private void RSA(bool crypt)
        {
            string input = fileTb_R.Text;
            if (crypt)
            {
                string temp = String.Join("", fileTb_R.Text.Where(c => c != '\n' && c != '\r' && c != '\t'));
                fileTb_R.Text = temp;
                p = Convert.ToInt16(pTb.Text);
                q = Convert.ToInt16(qTb.Text);
                e = 0;
                d = 0;
                n = p * q;
                euler = (p - 1) * (q - 1);
                Random random = new Random();
                int[] xy = { 1, 1 };
                resTb_R.Text = "";
                if (p == q && (!SimpleNumber(p) || !SimpleNumber(q)))
                {
                    MessageBox.Show("p and q values should not be equal. Value p or q is not simple");
                    return;
                }
                if (eTb.Text == string.Empty && dTb.Text == string.Empty)
                {
                    do
                    {
                        e = RSA_e(euler);
                        int[] dxy = ExtendedEuclideanRSA(e, euler);
                        d = dxy[1];
                    } while (d < 1 || d == e);
                    publickeyLbl.Text = "( " + n + ", " + e + " )";
                    privatekeyLbl.Text = "( " + n + ", " + d + " )";
                }
                else
                {
                    e = Convert.ToInt32(eTb.Text);
                    d = Convert.ToInt32(dTb.Text);
                    publickeyLbl.Text = "( " + n + ", " + e + " )";
                    privatekeyLbl.Text = "( " + n + ", " + d + " )";
                }
                BigInteger[] text = ConvertToIndex(fileTb_R.Text);
                rsaEncoded = RSA_E(e, n, text);
                string result = "";
                foreach (var block in rsaEncoded)
                {
                    numericalEncoded += ((int)block).ToString() + " ";
                    result += rAlphabet[(int)block % rAlphabet.Count].ToString();
                }

                resTb_R.Text = result;
                MessageBox.Show("RSA cryption has completed");
                if (MessageBox.Show("Copy numerical representation of ET into clipboard?", "RSA cryption",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    Clipboard.SetText(numericalEncoded);
            }
            else
            {
                if (rsaEncoded != null)
                {
                    BigInteger[] rsa_decoded = RSA_D(d, n, rsaEncoded);
                    resTb_R.Text = ConvertToSymbol(rsa_decoded);
                    MessageBox.Show("RSA decryption has completed");
                }
            }
        }

        // PT RSA convertation into numerical blocks
        public BigInteger[] ConvertToIndex(string input)
        {
            BigInteger[] result = new BigInteger[input.Length];

            for (int i = 0; i < input.Length; i++)
                result[i] = rAlphabet.IndexOf(input[i]);

            return result;
        }

        // ET RSA convertation into symbols
        public string ConvertToSymbol(BigInteger[] input)
        {
            string result = "";
            try
            {
                foreach (var letter in input)
                    result += rAlphabet[(int)letter % rAlphabet.Count].ToString();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            return result;
        }

        // RSA encryption
        private BigInteger[] RSA_E(BigInteger e, BigInteger n, BigInteger[] input)
        {
            input = ConvertToOneHalf(input);
            BigInteger[] result = input;
            for (int i = 0; i < input.Length; i++)
            {
                input[i] = BigInteger.ModPow(input[i], e, n);
            }
            return result;
        }

        // RSA decryption
        public BigInteger[] RSA_D(BigInteger e, BigInteger n, BigInteger[] input)
        {
            for (int i = 0; i < input.Length; i++)
            {
                input[i] = BigInteger.ModPow(input[i], e, n);
            }
            return ConvertToOne(input);
        }

        // convertion RSA numerical blocks from 1 to 1.5 symbol
        public BigInteger[] ConvertToOneHalf(BigInteger[] input)
        {
            List<BigInteger> text = new List<BigInteger>(input);
            for (int i = 0; i < text.Count % 3; i++)
                text.Add(rAlphabet.IndexOf(' '));
            List<BigInteger> result = new List<BigInteger>();
            for (int i = 0; i < text.Count; i += 3)
            {
                result.Add(100 * (int)(text[i] / 10) + 10 * (int)(text[i] % 10) + (int)(text[i + 1] / 10));
                result.Add(100 * (int)(text[i + 1] % 10) + 10 * (int)(text[i + 2] / 10) + (int)(text[i + 2] % 10));
            }
            return result.ToArray();
        }

        // convertion RSA numerical blocks from 1.5 to 1 symbol
        public BigInteger[] ConvertToOne(BigInteger[] tex)
        {
            List<BigInteger> text = new List<BigInteger>(tex);
            List<BigInteger> result = new List<BigInteger>();
            for (int i = 0; i < text.Count; i += 2)
            {
                result.Add(10 * (int)(text[i] / 100) + (int)(text[i] / 10) % 10);
                result.Add(10 * (int)(text[i] % 10) + (int)(text[i + 1] / 100));
                result.Add(10 * ((int)(text[i + 1] / 10) % 10) + (int)(text[i + 1] % 10));
            }
            return result.ToArray();
        }

        // if number is simple
        private bool SimpleNumber(int a)
        {
            if (a < 2 || a % 2 == 0)
                return false;

            if (a == 2)
                return true;

            for (long i = 2; i < a; i++)
                if (a % i == 0)
                    return false;

            return true;
        }

        // RSA search of e value
        private int RSA_e(int euler)
        {
            Random random = new Random();
            int e = random.Next(1, euler);
            while (true)
            {
                if (SimpleNumber(e) && GCD(e, euler) == 1)
                    break;
                e = random.Next(1, euler);
            }
            return e;
        }

        // finding GCD
        private int GCD(int a, int b)
        {
            while (a != 0 && b != 0)
            {
                if (a > b)
                    a %= b;
                else
                    b %= a;
            }
            return a == 0 ? b : a;
        }

        // extented Euclidean algorithm to find d inverted to e by the Euler number
        private int[] ExtendedEuclideanRSA(int a, int b)
        {
            int[] dxy = new int[3];
            if (b == 0)
            {
                dxy[0] = a;
                dxy[1] = 1;
                dxy[2] = 0;
                return dxy;
            }
            int[] t = { 1, 1 };
            dxy = ExtendedEuclideanRSA(b, a % b);
            t[0] = dxy[1];
            t[1] = dxy[2];
            dxy[1] = dxy[2];
            dxy[2] = t[0] - a / b * t[1];
            return dxy;
        }

        // DES implementation
        private byte[] DES(bool crypt, byte[] text)
        {
            string old_key = DkeyTb.Text;
            byte[] byte_key = Encoding.ASCII.GetBytes(old_key);
            int avg_key = 0;
            foreach (byte item in byte_key)
                avg_key += item;
            byte last_byte = (byte)(avg_key >> 4);
            string key = string.Empty;
            try
            {
                key = old_key.Substring(0, 7) + (char)last_byte;
            }
            catch (ArgumentOutOfRangeException)
            {
                MessageBox.Show("The key has not enough symbols");
                return null;
            }
            string initVector = old_key.Substring(0, 8).ToLower();
            using (var des = new DESCryptoServiceProvider())
            {
                try
                {
                    des.Key = Encoding.ASCII.GetBytes(key);
                    des.Mode = CipherMode.CBC;
                    des.Padding = PaddingMode.PKCS7;
                    des.IV = Encoding.ASCII.GetBytes(initVector);
                    using (var memStream = new MemoryStream())
                    {
                        CryptoStream cStream = null;
                        if (crypt)
                            cStream = new CryptoStream(memStream, des.CreateEncryptor(), CryptoStreamMode.Write);
                        else
                            cStream = new CryptoStream(memStream, des.CreateDecryptor(), CryptoStreamMode.Write);
                        cStream.Write(text, 0, text.Length);
                        cStream.FlushFinalBlock();
                        return memStream.ToArray();
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    return null;
                }
            }
        }

        // DES decryption
        private void DdecyphBtn_Click(object sender, EventArgs e)
        {
            resTb_D.Text = Encoding.UTF8.GetString(DES(false, encrypted));
            MessageBox.Show("RSA decryption has completed");
        }

        // DES encryption
        private void DcyphBtn_Click(object sender, EventArgs e)
        {
            string temp = String.Join("", fileTb_D.Text.Where(c => c != '\n' && c != '\r' && c != '\t'));
            fileTb_D.Text = temp;
            encrypted = DES(true, Encoding.UTF8.GetBytes(fileTb_D.Text));
            if (encrypted != null)
            {
                resTb_D.Text = BitConverter.ToString(encrypted).Replace("-", "");
                MessageBox.Show("RSA encryption has completed");
            }
        }

        // search for word and its count
        private void findbywordBtn_Click(object sender, EventArgs e)
        {
            string temp = fileTb.Text;
            fileTb.Clear();
            fileTb.Text = temp;
            if (wordTb.Text.Count() > 0)
                HighlightWords(fileTb, wordTb.Text, false);
            else
                wordLbl.Text = string.Empty;
        }

        private void findbyamountBtn_Click(object sender, EventArgs e)
        {
            wordsByCount.Clear();
            if (countTb.Text != string.Empty && countTb.Text.All(char.IsDigit))
            {
                wordsByCount = fileTb.Text.Split(' ').ToList();
                for (int i = 0; i < wordsByCount.Count; i++)
                {
                    if (!wordsByCount[i].All(char.IsLetter))
                    {
                        wordsByCount.RemoveAt(i);
                        i--;
                        continue;
                    }
                    if (wordsByCount[i].Length != Int32.Parse(countTb.Text))
                    {
                        wordsByCount.RemoveAt(i);
                        i--;
                        continue;
                    }
                }
                if (wordsByCount.Count == 0)
                {
                    countLbl.Text = "0";
                    return;
                }
                wordsByCount = wordsByCount.Distinct().ToList();
                for (int i = 0; i < wordsByCount.Count; i++)
                    HighlightWords(fileTb, wordsByCount[i], true);
                countByCount = 0;
            }
            else
                countLbl.Text = string.Empty;
        }

        private void findbyopenwordBtn_Click(object sender, EventArgs e)
        {
            string temp = openTb.Text;
            openTb.Clear();
            openTb.Text = temp;
            if (openwordTb.Text.Count() > 0)
                HighlightWords(openTb, openwordTb.Text, false);
            else
                openwordLbl.Text = string.Empty;
        }

        private void findbyopenamountBtn_Click(object sender, EventArgs e)
        {
            wordsByCount.Clear();
            if (opencountTb.Text != string.Empty && opencountTb.Text.All(char.IsDigit))
            {
                wordsByCount = openTb.Text.Split(' ').ToList();
                for (int i = 0; i < wordsByCount.Count; i++)
                {
                    if (!wordsByCount[i].All(char.IsLetter))
                    {
                        wordsByCount.RemoveAt(i);
                        i--;
                        continue;
                    }
                    if (wordsByCount[i].Length != Int32.Parse(opencountTb.Text))
                    {
                        wordsByCount.RemoveAt(i);
                        i--;
                        continue;
                    }
                }
                if (wordsByCount.Count == 0)
                {
                    opencountLbl.Text = "0";
                    return;
                }
                wordsByCount = wordsByCount.Distinct().ToList();
                for (int i = 0; i < wordsByCount.Count; i++)
                    HighlightWords(openTb, wordsByCount[i], true);
                countByCount = 0;
            }
            else
                opencountLbl.Text = string.Empty;
        }

        private void findbycryptwordBtn_Click(object sender, EventArgs e)
        {
            string temp = cryptTb.Text;
            cryptTb.Clear();
            cryptTb.Text = temp;
            if (cryptwordTb.Text.Count() > 0)
                HighlightWords(cryptTb, cryptwordTb.Text, false);
            else
                cryptwordLbl.Text = string.Empty;
        }

        private void findbycryptamountBtn_Click(object sender, EventArgs e)
        {
            wordsByCount.Clear();
            if (cryptcountTb.Text != string.Empty && cryptcountTb.Text.All(char.IsDigit))
            {
                wordsByCount = cryptTb.Text.Split(' ').ToList();
                for (int i = 0; i < wordsByCount.Count; i++)
                {
                    if (!wordsByCount[i].All(char.IsLetter))
                    {
                        wordsByCount.RemoveAt(i);
                        i--;
                        continue;
                    }
                    if (wordsByCount[i].Length != Int32.Parse(cryptcountTb.Text))
                    {
                        wordsByCount.RemoveAt(i);
                        i--;
                        continue;
                    }
                }
                if (wordsByCount.Count == 0)
                {
                    cryptcountLbl.Text = "0";
                    return;
                }
                wordsByCount = wordsByCount.Distinct().ToList();
                for (int i = 0; i < wordsByCount.Count; i++)
                    HighlightWords(cryptTb, wordsByCount[i], true);
                countByCount = 0;
            }
            else
                cryptcountLbl.Text = string.Empty;
        }

        private void countTb_TextChanged(object sender, EventArgs e)
        {
            fileTb.TextChanged -= fileTb_TextChanged;
            fileTb.Clear();
            fileTb.Text = tempFile;
            fileTb.TextChanged += fileTb_TextChanged;
            countLbl.Text = string.Empty;
        }

        private void wordTb_TextChanged(object sender, EventArgs e)
        {
            fileTb.TextChanged -= fileTb_TextChanged;
            fileTb.Clear();
            fileTb.Text = tempFile;
            fileTb.TextChanged += fileTb_TextChanged;
            wordLbl.Text = string.Empty;
        }

        private void fileTb_TextChanged(object sender, EventArgs e)
        {
            fileTb_C.Text = fileTb.Text;
            fileTb_S.Text = fileTb.Text;
            tempFile = fileTb.Text;
        }

        private void openwordTb_TextChanged(object sender, EventArgs e)
        {
            openTb.TextChanged -= openTb_TextChanged;
            openTb.Clear();
            openTb.Text = tempFile2;
            openTb.TextChanged += openTb_TextChanged;
            openwordLbl.Text = string.Empty;
        }

        private void opencountTb_TextChanged(object sender, EventArgs e)
        {
            openTb.TextChanged -= openTb_TextChanged;
            openTb.Clear();
            openTb.Text = tempFile2;
            openTb.TextChanged += openTb_TextChanged;
            opencountLbl.Text = string.Empty;
        }

        private void cryptwordTb_TextChanged(object sender, EventArgs e)
        {
            cryptTb.TextChanged -= cryptTb_TextChanged;
            cryptTb.Clear();
            cryptTb.Text = tempFile3;
            cryptTb.TextChanged += cryptTb_TextChanged;
            cryptwordLbl.Text = string.Empty;
        }

        private void cryptcountTb_TextChanged(object sender, EventArgs e)
        {
            cryptTb.TextChanged -= cryptTb_TextChanged;
            cryptTb.Clear();
            cryptTb.Text = tempFile3;
            cryptTb.TextChanged += cryptTb_TextChanged;
            cryptcountLbl.Text = string.Empty;
        }

        private void openTb_TextChanged(object sender, EventArgs e)
        {
            tempFile2 = openTb.Text;
        }

        private void savecompareBtn_Click(object sender, EventArgs e)
        {
            ExportToPNG(freqChart4);
        }

        private void savecomparetableBtn_Click(object sender, EventArgs e)
        {
            ExportToExcel(symDgv4);
        }

        private void cryptTb_TextChanged(object sender, EventArgs e)
        {
            tempFile3 = cryptTb.Text;
        }

        // load data for monograms comparison
        private void loadvigeneregramBtn_Click(object sender, EventArgs e)
        {
            sourcetextmonoCb.SelectedIndex = -1;
            sourcetextmonoCb.SelectedIndex = 0;
            if (MessageBox.Show("Monogram analysis has completed. Go to results?", "Vigenere monograms",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                metroTabControl1.SelectedIndex = 11;
        }

        private void loadhillgramBtn_Click(object sender, EventArgs e)
        {
            sourcetextmonoCb.SelectedIndex = -1;
            sourcetextmonoCb.SelectedIndex = 1;
            if (MessageBox.Show("Monogram analysis has completed. Go to results?", "Hill monograms",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                metroTabControl1.SelectedIndex = 11;
        }

        private void loadfeistelgramBtn_Click(object sender, EventArgs e)
        {
            sourcetextmonoCb.SelectedIndex = -1;
            sourcetextmonoCb.SelectedIndex = 2;
            if (MessageBox.Show("Monogram analysis has completed. Go to results?", "Feistel monograms",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                metroTabControl1.SelectedIndex = 11;
        }

        private void loadrsagramBtn_Click(object sender, EventArgs e)
        {
            sourcetextmonoCb.SelectedIndex = -1;
            sourcetextmonoCb.SelectedIndex = 3;
            if (MessageBox.Show("Monogram analysis has completed. Go to results?", "RSA monograms",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                metroTabControl1.SelectedIndex = 11;
        }

        private void loaddesgramBtn_Click(object sender, EventArgs e)
        {
            sourcetextmonoCb.SelectedIndex = -1;
            sourcetextmonoCb.SelectedIndex = 4;
            if (MessageBox.Show("Monogram analysis has completed. Go to results?", "DES monograms",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                metroTabControl1.SelectedIndex = 11;
        }

        private void exitBtn_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void SavemonogramBtn_Click(object sender, EventArgs e)
        {
            ExportToPNG(freqChart);
        }

        private void SavemonotableBtn_Click(object sender, EventArgs e)
        {
            ExportToExcel(symDgv);
        }

        private void savebigramBtn_Click(object sender, EventArgs e)
        {
            ExportToPNG(freqChart2);
        }

        private void savebitableBtn_Click(object sender, EventArgs e)
        {
            ExportToExcel(symDgv2);
        }

        private void savetrigramBtn_Click(object sender, EventArgs e)
        {
            ExportToPNG(freqChart3);
        }

        private void savetritableBtn_Click(object sender, EventArgs e)
        {
            ExportToExcel(symDgv3);
        }

        private void ExportToPNG(Chart chart)
        {
            string file = string.Empty;
            if (chart.Name == "freqChart")
                file = "monograms.png";
            if (chart.Name == "freqChart2")
                file = "bigrams.png";
            if (chart.Name == "freqChart3")
                file = "trigrams.png";
            if (chart.Name == "freqChart4")
                file = "compare_monograms.png";
            chart.Size = new Size(1510, 801);
            if (chart.Name != "freqChart4")
            {
                chart.Series[0].Font = new Font("Microsoft Sans Serif", 16f);
            }
            else
            {
                chart.Series[2].Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Bold);
                chart.Series[3].Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Bold);
            }
            chart.ChartAreas[0].AxisX.ScaleView.Size = chart.Series[0].Points.Count;
            chart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Microsoft Sans Serif", 16f);
            chart.Legends[0].Font = new Font("Microsoft Sans Serif", 16f);
            chart.SaveImage(AppDomain.CurrentDomain.BaseDirectory + file, ChartImageFormat.Png);
            if (chart.Name == "freqChart")
            {
                chart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Microsoft Sans Serif", 14.25f);
                chart.Series[0].Font = new Font("Microsoft Sans Serif", 12f);
            }

            if (chart.Name == "freqChart2")
            {
                chart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Microsoft Sans Serif", 9.75f);
                chart.Series[0].Font = new Font("Microsoft Sans Serif", 12f);
            }

            if (chart.Name == "freqChart3")
            {
                chart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Microsoft Sans Serif", 8.25f);
                chart.Series[0].Font = new Font("Microsoft Sans Serif", 12f);
            }
            if (chart.Name == "freqChart4")
            {
                if (freqChart4.Series[0].Points.Count > 15)
                    freqChart4.ChartAreas[0].AxisX.ScaleView.Size = 5;
                chart.Legends[0].Font = new Font("Microsoft Sans Serif", 8.25f);
                chart.Series[2].Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular);
                chart.Series[3].Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular);
                chart.Size = new Size(591, 351);
            }
            if (chart.Name != "freqChart4")
            {
                chart.Size = new Size(669, 351);
                chart.ChartAreas[0].AxisX.ScaleView.Size = 15;
            }
            if (MessageBox.Show("The picture has been successfully saved to the executable folder. Want to open it?", "Picture created",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                Process.Start(AppDomain.CurrentDomain.BaseDirectory + file);
        }

        private void ExportToExcel(DataGridView dgv)
        {
            dgv.SelectAll();
            DataObject dataObj = dgv.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);

            Excel.Application xlexcel;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }

        private void sourcetextmonoCb_SelectedIndexChanged(object sender, EventArgs e)
        {
            sortCb2.SelectedIndex = -1;
            if (sourcetextmonoCb.SelectedIndex == 0)
                CompareMonograms("vigenere");
            if (sourcetextmonoCb.SelectedIndex == 1)
                CompareMonograms("hill");
            if (sourcetextmonoCb.SelectedIndex == 2)
                CompareMonograms("feistel");
            if (sourcetextmonoCb.SelectedIndex == 3)
                CompareMonograms("rsa");
            if (sourcetextmonoCb.SelectedIndex == 4)
                CompareMonograms("des");
        }

        private void sortCb2_SelectedIndexChanged(object sender, EventArgs e)
        {
            freqChart4.Series[0].Enabled = true;
            freqChart4.Series[1].Enabled = true;
            freqChart4.Series[2].Enabled = false;
            freqChart4.Series[3].Enabled = false;
            if (sortCb2.SelectedIndex == 0)
            {
                symDgv4.Sort(symDgv4.Columns[0], ListSortDirection.Ascending);
                freqChart4.DataManipulator.Sort(PointSortOrder.Ascending, "AxisLabel", "PT,ET");
            }
            else if (sortCb2.SelectedIndex == 1)
            {
                symDgv4.Sort(symDgv4.Columns[1], ListSortDirection.Descending);
                freqChart4.DataManipulator.Sort(PointSortOrder.Descending, "Y", "PT,ET");
            }
            else if (sortCb2.SelectedIndex == 2)
            {
                symDgv4.Sort(symDgv4.Columns[2], ListSortDirection.Descending);
                freqChart4.DataManipulator.Sort(PointSortOrder.Descending, "Y", "ET,PT");
            }
            else if (sortCb2.SelectedIndex == 3)
            {
                freqChart4.Series[0].Enabled = false;
                freqChart4.Series[1].Enabled = false;
                freqChart4.Series[2].Enabled = true;
                freqChart4.Series[3].Enabled = true;
            }
        }

        private void sortCb_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (sortCb.SelectedIndex == 0)
            {
                symDgv.Sort(symDgv.Columns[0], ListSortDirection.Ascending);
                freqChart.Series[0].Sort(PointSortOrder.Ascending, "AxisLabel");
            }
            else
            {
                symDgv.Sort(symDgv.Columns[1], ListSortDirection.Descending);
                freqChart.Series[0].Sort(PointSortOrder.Descending, "Y");
            }

        }
    }
}

