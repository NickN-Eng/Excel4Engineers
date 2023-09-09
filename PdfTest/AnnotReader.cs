using iTextSharp.text.pdf;
using Org.BouncyCastle.Asn1;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfTest
{
    //public class AnnotReader
    //{
    //    public static void ReadAnnotations(string inputPath)
    //    {
    //        PdfReader pdfReader = new PdfReader(inputPath);

    //        //get the PdfDictionary of the 1st page
    //        PdfDictionary pageDict = pdfReader.GetPageN(1);

    //        //get annotation array
    //        PdfArray annotArray = pageDict.GetAsArray(PdfName.ANNOTS);

    //        //iterate through annotation array
    //        int size = annotArray.Size;
    //        for (int i = 0; i < size; i++)
    //        {

    //            //get value of /Contents
    //            PdfDictionary dict = annotArray.GetAsDict(i);
    //            PdfString contents = dict.GetAsString(PdfName.CONTENTS);
    //            PdfArray rectArray = dict.GetAsArray(PdfName.RECT);



    //            //check if /Contents key exists
    //            if (contents != null)
    //            {

    //                //set new value
    //                dict.Put(PdfName.CONTENTS, new PdfString("value has been changed"));
    //                dict.Put(PdfName.DA, new PdfString("value has been changed"));
    //                dict.Remove(PdfName.AP);
    //                dict.Remove(PdfName.N);
    //                //dict.Put(PdfName.AP, new PdfString("value has been changed"));
    //            }
    //        }
    //    }
    //}
}
