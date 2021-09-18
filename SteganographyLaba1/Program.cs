
using System;
using System.Collections;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace SteganographyLaba1
{
    class Program
    {
        static void TestSteganography(string pathText,string pathStegText, string mess)
        {
            HiddenMessage exp = new HiddenMessage();
            ArrayList text = exp.Read(pathText);
            ArrayList res = exp.hideMessage(text, mess);
            exp.Write(pathStegText,res);
            ArrayList stegText = exp.Read(pathStegText);
            exp.findMessage(stegText);
            
        }
        static void TestSteganographyForStrange(string pathText, string pathStegText, string mess)
        {
            HiddenMessage exp = new HiddenMessage();
            ArrayList text = exp.ReadHard(pathText);
            ArrayList res = exp.hideMessage(text, mess);
            exp.WriteHard(pathStegText, res);
            ArrayList stegText = exp.ReadHard(pathStegText);
            exp.findMessage(stegText);


        }
        static void Main(string[] args)
        {
            TestSteganography(@"C:\Users\mustard\source\repos\SteganographyLaba1\SteganographyLaba1\Sourses\HTML.html",
                              @"C:\Users\mustard\source\repos\SteganographyLaba1\SteganographyLaba1\Sourses\HTMLsteg.html",
                              "TheCakeisaLie");
            TestSteganographyForStrange(@"C:\Users\mustard\source\repos\SteganographyLaba1\SteganographyLaba1\Sourses\DOC.doc",
                              @"C:\Users\mustard\source\repos\SteganographyLaba1\SteganographyLaba1\Sourses\DOCsteg.doc",
                              "TheCakeisaLie");
            TestSteganographyForStrange(@"C:\Users\mustard\source\repos\SteganographyLaba1\SteganographyLaba1\Sourses\RTF.rtf",
                              @"C:\Users\mustard\source\repos\SteganographyLaba1\SteganographyLaba1\Sourses\RTFsteg.rtf",
                              "TheCakeisaLie");
            TestSteganography(@"C:\Users\mustard\source\repos\SteganographyLaba1\SteganographyLaba1\Sourses\CPP.cpp",
                              @"C:\Users\mustard\source\repos\SteganographyLaba1\SteganographyLaba1\Sourses\CPPsteg.cpp",
                              "TheCakeisaLie");
            TestSteganography(@"C:\Users\mustard\source\repos\SteganographyLaba1\SteganographyLaba1\Sourses\TXT.txt",
                              @"C:\Users\mustard\source\repos\SteganographyLaba1\SteganographyLaba1\Sourses\TXTsteg.txt",
                              "TheCakeisaLie");
        }
    }
}
