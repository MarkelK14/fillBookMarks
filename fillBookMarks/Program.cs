using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;

namespace fillBookMarks
{
    class Program
    {
        static Microsoft.Office.Interop.Word.Application myWord;
        static Microsoft.Office.Interop.Word.Document doc;
        static string wordResultado;
        static string pathWord;
        static string[] bookmarks;
        static string[] values;
        static char delimiter1 = '|';
        static char delimiter2 = ';';
        static string rutadestino;
        static int result;
        static void Main(string[] args)
        {
            result = 0;
            string ficheroLog = fillBookMarks.Properties.Settings.Default.PATH_LOG;
            Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            Console.WriteLine("Ejecutando...");

            try
            {
                Console.WriteLine("1. Obteniendo argumentos...");


                //args = new string[4];
                ////Plantilla
                //args.SetValue();
                ////Bookmarks
                //args.SetValue();
                ////Valores
                //args.SetValue();
                ////Resultado
                //args.SetValue();


                if (args.Length == 0)
                {
                    throw new ArgumentException("No hay argumentos");
                }
                else if (args.Length != 4)
                {
                    if (args.Length > 4)
                    {
                        throw new ArgumentException("Menos de 4 argumentos");
                    }
                    else
                    {
                        throw new ArgumentException("Más de 4 argumentos");
                    }
                }
                else
                {
                    Console.WriteLine("2. Validando argumentos...");
                    if (args[0] == null || args[0].ToString().CompareTo(string.Empty) == 0)
                    {
                        throw new ArgumentException("Falta el primero argumento, la ruta de la plantilla Word");
                    }
                    else
                    {
                        pathWord = args[0];
                    }

                    if (args[1] == null || args[1].ToString().CompareTo(string.Empty) == 0)
                    {
                        throw new ArgumentException("Falta el segundo argumento, el nombre de los Bookmarks separados por ; o por |");
                    }
                    else
                    {
                        bookmarks = args[1].Split(delimiter1);
                        if (bookmarks.Length < 2)
                        {
                            bookmarks = args[1].Split(delimiter2);
                        }
                        if (bookmarks.Length == 0)
                        {
                            throw new ArgumentException("El nombre de los Bookmarks debe estar separado por ; o por |");
                        }
                    }

                    if (args[2] == null || args[2].ToString().CompareTo(string.Empty) == 0)
                    {
                        throw new ArgumentException("Falta el tercer argumento, el valor de los Bookmarks separados por ; o por |");
                    }
                    else
                    {
                        values = args[2].Split(delimiter1);
                        if (values.Length < 2)
                        {
                            values = args[2].Split(delimiter2);
                        }
                        if (values.Length == 0)
                        {
                            throw new ArgumentException("El valor de los Bookmarks debe estar separado por ; o por |");
                        }
                    }


                    if (args[3] == null || args[3].ToString().CompareTo(string.Empty) == 0)
                    {
                        throw new ArgumentException("Falta el cuarto argumento, la ruta del documento resultado");
                    }
                    else
                    {
                        wordResultado = args[3];
                        rutadestino = System.IO.Path.GetDirectoryName(wordResultado);
                    }

                    if (!System.IO.File.Exists(pathWord))
                    {
                        throw new System.IO.FileNotFoundException("No existe la plantilla " + pathWord);
                    }

                    if (!System.IO.File.Exists(rutadestino))
                    {
                        throw new System.IO.FileNotFoundException("No existe la ruta donde debería guardar el documento:  " + rutadestino);
                    }

                    if (bookmarks.Length != values.Length)
                    {
                        throw new ArgumentException("Nombres de marcadores <> valores de marcadores");
                    }

                    Console.WriteLine("3. Rellenando fichero Word...");

                    myWord = new Microsoft.Office.Interop.Word.Application();
                    doc = myWord.Documents.Add(pathWord);

                    for (int index = 0; index < bookmarks.Length; index++)
                    {
                        if (!buscarCambiarBookmarks(bookmarks[index], values[index]))
                        {
                            continue;
                        }
                    }

                    Console.WriteLine("4. Guardando fichero Word...");
                    doc.SaveAs(wordResultado, 16);


                }
            }
            catch (Exception ex)
            {
                result = -1;
                Console.WriteLine("No se puede generar el documento : " + ex.Message);
                File.WriteAllText(ficheroLog, ex.Message + " " + ex.StackTrace);
            }
            finally
            {
                if (doc != null)
                {
                    ((_Document)doc).Close(WdSaveOptions.wdDoNotSaveChanges);
                }

                if(myWord != null)
                {
                    ((_Application)myWord).Quit();
                }

                Console.WriteLine("5. Fin del programa");
                Environment.ExitCode = result;

            }
        }



        static Boolean buscarCambiarBookmarks(string nombreBookmark, string valor)
        {
            for (int index = 1; index <= doc.Bookmarks.Count; index++)
            {
                if (doc.Bookmarks[index].Name == nombreBookmark)
                {
                    doc.Bookmarks[1].Range.Text = valor;
                    return true;
                }
            }
            return false;
        }


    }
}
