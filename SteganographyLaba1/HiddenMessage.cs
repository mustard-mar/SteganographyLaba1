using System;
using System.Collections;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace SteganographyLaba1
{
    class HiddenMessage
    {
       // private ArrayList message;
        //private ArrayList text;
        //private long sizeOfText;
        

        
        public HiddenMessage() {
           
        }
        public ArrayList hideMessage(ArrayList text, string m)
        {
            ArrayList message = new ArrayList();
            long sizeOfText = text.Count;
            byte sizeOfMessage = (byte)m.Length;
            message.Add(sizeOfMessage);
            for (int j = 0; j < sizeOfMessage; j++)
            {
                message.Add((byte)m[j]);
            }
            ArrayList result = new ArrayList();
            int n = 0;
            int r = 0;
            int maxL = 0;
            ArrayList L = new ArrayList();
            for (int j = 0; j < sizeOfText; j++)
            {
                if ((byte)text[j] != '\n') n = n + 1;
                else if ((byte)text[j] == '\n')
                {
                    if (maxL < n - 1) maxL = n - 1;
                    L.Add(n - 1);
                    n = 0;
                    r = r + 1;
                }
            }
            L.Add(n);
            r = 0;
            int mu = 0;
            int i = 0; ;
            int countOfBits = 0;
            while (mu < 8 * message.Count)
                {
                    for (; i < sizeOfText; i++)
                    {
                        if ((byte)text[i] != 13) result.Add(text[i]);
                        else if (i + 2 < sizeOfText)
                        {
                            i += 2;
                            int k = 1;
                            while (k <= maxL - (int)L[r])
                            {
                                if (GetBit(mu % 8, (byte)message[countOfBits]) == 0) result.Add((byte)32);
                                if (GetBit(mu % 8, (byte)message[countOfBits]) == 1) result.Add((byte)160);
                                mu = mu + 1;
                                countOfBits = mu / 8;
                                if (mu >= 8 * message.Count)
                                    break;
                                k = k + 1;
                            }

                            result.Add((byte)13);
                            result.Add((byte)10);
                            r = r + 1;
                            
                            break;
                        }
                    }
                }
            for (; i < sizeOfText; i++)
                {
                    result.Add(text[i]);
                }
            
            return result;

        }

        internal ArrayList Read(string pathText)
        {
            ArrayList text = new ArrayList();
            try
            {
                using (StreamReader sr = new StreamReader(new FileStream(pathText,
                                                          FileMode.Open), System.Text.Encoding.UTF8))
                {
                    byte b;
                    var fi = new FileInfo(pathText);
                    long sizeOfText = fi.Length;
                    int j = 0;
                    while (j < sizeOfText)
                    {
                        b = (byte)sr.Read();
                        text.Add(b);
                        j++;

                    }
                    text.Add((byte)13);
                    text.Add((byte)10);
                    sizeOfText = fi.Length + 2;
                   
                }
                return text;
            }
            catch (Exception e)
            {
                // Let the user know what went wrong.
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
                return null;
            }
        }
        internal ArrayList ReadHard(string pathText)
        {
            ArrayList text = new ArrayList();
            Word._Application word_app = new Word.Application();
            Word._Document word_doc = word_app.Documents.Open(pathText);
            string str = "";
            for (int i = 0; i < word_doc.Paragraphs.Count; i++)
            {
                str += word_doc.Paragraphs[i + 1].Range.Text +"\n";
            }
            word_doc.Close();
            word_app.Quit();
            int j = 0;
            byte b = 0;
            while (j < str.Length)
            {
                b = (byte)str[j];
                text.Add(b);
                j++;
            }
            return text;
        }
        internal void Write(string pathStegText, ArrayList text)
        {
            using (StreamWriter writer = new StreamWriter(new FileStream(pathStegText,
                                                          FileMode.Create), System.Text.Encoding.UTF8))
            {
                for (int i = 0; i < text.Count; i++)
                    writer.Write((char)((byte)text[i]));
            }
        }
        internal void WriteHard(string pathStegText, ArrayList text)
        {
            Word._Application word_app = new Word.Application();
            Word._Document word_doc = word_app.Documents.Add();
            object index = 0;
            string str = "";
            Word.Range rng = word_doc.Range(ref index, ref index);
            for (int i = 0; i < text.Count; i++)
            {
                str = str +(char)((byte)text[i]);
            }
            rng.Text = str;
            word_doc.SaveAs2(pathStegText);
            word_doc.Close();
            word_app.Quit();
        }

        public void findMessage(ArrayList stegText)
        {

            ArrayList ArrayByte = new ArrayList();
            byte sizeOfmess = 0;
            long sizeOfsteg = stegText.Count;
            int mu = 0;
            for (int i = 0; i < sizeOfsteg; i++)
            {
                if((byte)stegText[i]== 13)
                {
                    int j = i;
                    while ((byte)stegText[j - 1] == 32 || (byte)stegText[j - 1] == 160) j--;
                    if(j!=i)for (; j <= i-1; j++)
                    {
                        ArrayByte.Add(stegText[j]);
                        mu++;
                    }

                }
            }
            ArrayList result = new ArrayList();
            for (int i = 0; i < (ArrayByte.Count / 8); i++)
            {
                result.Add((byte)0);
            }
            GetMess(ArrayByte, result);
            sizeOfmess = (byte)result[0];
            Console.WriteLine("Размер сообщения: "+ sizeOfmess);
            Console.WriteLine("Сообщение: \n");
            for (int i = 1; i < result.Count; i++)
                Console.Write("" + (char)((byte)result[i]));
        }
        private byte GetBit(int mu, byte m)
        {
            return (byte)(((byte)m & (0b00000001 << mu)) >> mu);
        }

        private void GetMess(ArrayList arrByte,ArrayList mess)
        {
            for (int i = 0; i < arrByte.Count; i++)
            {
                if ((byte)arrByte[i] == 160)
                    SetBit(mess,i);
            }
        }

        private void SetBit(ArrayList mess, int i)
        {
            mess[i / 8] = (byte)(((byte)mess[i / 8]) | (1 << (i % 8)));
        }
    }
}
