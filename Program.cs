using System;
using System.IO;

namespace ConversorRPTtoCSV
{
    public class Program
    {
        /// <summary>
        /// The main entry point to the application.
        /// </summary>
        /// <param name="args">Command line arguments.</param>
        internal static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    string inputFile;
                    string outputFile;

                    inputFile = args[i];
                    outputFile = Path.GetFileNameWithoutExtension(args[i]) + ".csv";

                    Environment.CurrentDirectory =
                      Path.GetDirectoryName(inputFile).Length == 0 ?
                      Environment.CurrentDirectory :
                      Path.GetFullPath(Path.GetDirectoryName(inputFile));

                    using (StreamReader inputReader = File.OpenText(inputFile))
                    {
                        string firstLine = inputReader.ReadLine();
                        string secondLine = inputReader.ReadLine();

                        string[] underscores = secondLine.Split(new char[] { ' ' },
                                 StringSplitOptions.RemoveEmptyEntries);

                        string[] fields = new string[underscores.Length];
                        int[] fieldLengths = new int[underscores.Length];

                        for (int j = 0; j < fieldLengths.Length; j++)
                        {
                            fieldLengths[j] = underscores[j].Length;
                        }

                        int fileNumber = 0;

                        StreamWriter outputWriter = null;

                        try
                        {
                            outputWriter = File.CreateText(outputFile.Insert(
                              outputFile.LastIndexOf("."), "_" +
                              fileNumber.ToString()));
                            fileNumber++;

                            int lineNumber = 0;

                            WriteLineToCsv(outputWriter, fieldLengths, firstLine);
                            lineNumber++;

                            string line;

                            while ((line = inputReader.ReadLine()) != null)
                            {
                                if (lineNumber >= 65536)
                                {
                                    outputWriter.Close();
                                    outputWriter = File.CreateText(outputFile.Insert(
                                      outputFile.LastIndexOf("."),
                                      "_" + fileNumber.ToString()));
                                    fileNumber++;

                                    lineNumber = 0;

                                    WriteLineToCsv(outputWriter, fieldLengths, firstLine);
                                    lineNumber++;
                                }

                                if (!WriteLineToCsv(outputWriter, fieldLengths, line))
                                {
                                    break;
                                }

                                lineNumber++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                            Console.WriteLine("NOTE: Input file must not " +
                                              "have any newline characters " +
                                              "as field contents.");
                            Console.WriteLine();
                            Console.WriteLine("Press any key to continue...");
                            Console.ReadKey(true);
                        }
                        finally
                        {
                            if (outputWriter != null)
                            {
                                outputWriter.Close();
                            }
                        }

                        // If we only had one file created,
                        // we don't need the file number in the name.
                        if (fileNumber == 1)
                        {
                            try
                            {
                                if (File.Exists(outputFile))
                                {
                                    File.Delete(outputFile);
                                }

                                File.Move(outputFile.Insert(
                                  outputFile.LastIndexOf("."), "_0"),
                                  outputFile);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex);
                                Console.WriteLine("Press any key to continue...");
                                Console.ReadKey(true);
                            }
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Converts the ouput of a SQL Server " +
                                  "Management Studio .rpt file to a CSV file.");
                Console.WriteLine("You can generate a .rpt file " +
                  "by selecting \"Results to File\" in the toolbar.");
                Console.WriteLine();
                Console.WriteLine("Usage: RptToCsv.exe <inputFile1> [<inputFile2> ...]");
                return;
            }
        }

        private static bool WriteLineToCsv(StreamWriter outputWriter,
                                          int[] fieldLengths, string line)
        {
            if (line.Length == 0)
            {
                return false;
            }

            int index = 0;

            for (int i = 0; i < fieldLengths.Length; i++)
            {
                string value;

                if (i < fieldLengths.Length - 1)
                {
                    value = line.Substring(index, fieldLengths[i]);
                }
                else
                {
                    value = line.Substring(index);
                }

                value = value.Replace("\"", "\"\"");
                value = value.Trim();

                if (value == "NULL")
                {
                    value = string.Empty;
                }

                outputWriter.Write("\"{0}\"", value);
                index += fieldLengths[i] + 1;

                if (i < fieldLengths.Length - 1)
                {
                    outputWriter.Write(",");
                }
                else
                {
                    outputWriter.WriteLine();
                }
            }

            return true;
        }
    }
}
