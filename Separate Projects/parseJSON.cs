using System;
using System.Linq;
using System.IO;

//Needs a file named benTest1.txt on the desktop to work.
namespace SandBox
{
	class Program
	{
        static void Main()
        {
            string text = System.IO.File.ReadAllText(@"C:\Users\ben\Desktop\benTest1.txt");
            String[] s1 = new String[9];
            int indexOfCDR = 0;
            int indexOfColon = 0;
            int diffBetweenCDRandColon = 0;
            int i = 0;
            string toEnter;


            text = text.Remove(0, 2); //Removes the first two characters
            text = text.Replace("\\", ""); //Removes the backslashes

            //Will enter loop if there is a matching cdr and will not exit until they have all been removed.
            while (text.IndexOf("cdr") != -1)
            {
                indexOfCDR = text.IndexOf("cdr") - 1; //-1 so it heads to the front of cdr...
                indexOfColon = text.IndexOf(":", indexOfCDR) + 1; //+1 so it grabs the final :, two parameters so it starts where the correct indexOfCDR is.
                Console.WriteLine(indexOfCDR);
                Console.WriteLine(indexOfColon);
                diffBetweenCDRandColon = indexOfColon - indexOfCDR;
                text = text.Remove(indexOfCDR, diffBetweenCDRandColon);
            }
            text = text.Insert(0, "[");
            text = text.Insert(text.Length, "]");

            File.WriteAllText(@"C:\Users\ben\Desktop\benTest2.txt", text);


            //Need to split up the file into several pastable sections
            //Console.WriteLine("The total length of the file is: {0}", text.Length);

            Console.WriteLine("Press any key to exit.");
            System.Console.ReadKey();
        }
       
    }
}
