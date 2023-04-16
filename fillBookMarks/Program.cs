using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace fillBookMarks
{
    class Program
    {
        static Microsoft.Office.Interop.Word.Application templateWord;
        static Microsoft.Office.Interop.Word.Document documentoWord;
        static string resultWord;
        static string pathWord;
        static string[] marcadores;
        static string[] valores;
        static char delimitador1 = '|';
        static char delimitador2 = ';';
        static string pathDestino;
        static int result;
        static void Main(string[] args)
        {
            result = 0;

            //El log se guarda en la misma ruta donde se encuentre el exe
            string ficheroLog = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            Console.WriteLine("Ejecutando programa...");

            try
            {
                Console.WriteLine("Obteniendo argumentos...");

                //args = new string[4];
                ////Plantilla
                //args.SetValue("C:\\vs_projects\\fillBookMarks\\docx\\Template.docx",0);
                ////Bookmarks
                //args.SetValue("NOMBRE|DINERO|CONSECUENCIA", 1);
                ////Valores
                //args.SetValue("Alfredo|536,75|llamaremos a nuestros abogados", 2);
                ////Resultado
                //args.SetValue("C:\\vs_projects\\fillBookMarks\\docx\\Resultado.docx", 3);


                if (args.Length == 0)
                {
                    throw new ArgumentException("No hay argumentos");
                }
                else if (args.Length != 4)
                {
                    if (args.Length > 4)
                    {
                        throw new ArgumentException("Menos de 4 argumentos: " + args.Length + "argumentos encontrados");
                    }
                    else
                    {
                        throw new ArgumentException("Más de 4 argumentos: " + args.Length + "argumentos encontrados");
                    }
                }
                else
                {
                    Console.WriteLine("Validando argumentos...");
                    if (args[0] == null || args[0].ToString().CompareTo(string.Empty) == 0)
                    {
                        throw new ArgumentException("Falta el primer argumento: Ruta de la plantilla Word");
                    }
                    else
                    {
                        pathWord = args[0];
                    }

                    if (args[1] == null || args[1].ToString().CompareTo(string.Empty) == 0)
                    {
                        throw new ArgumentException("Falta el segundo argumento: Nombre de los marcadores separados por ; o por |");
                    }
                    else
                    {
                        marcadores = args[1].Split(delimitador1);
                        if (marcadores.Length < 2)
                        {
                            marcadores = args[1].Split(delimitador2);
                        }
                        if (marcadores.Length == 0)
                        {
                            throw new ArgumentException("El delimitador de los marcadores debe ser ; o |");
                        }
                    }

                    if (args[2] == null || args[2].ToString().CompareTo(string.Empty) == 0)
                    {
                        throw new ArgumentException("Falta el tercer argumento: Valor de los marcadores separados por ; o por |");
                    }
                    else
                    {
                        valores = args[2].Split(delimitador1);
                        if (valores.Length < 2)
                        {
                            valores = args[2].Split(delimitador2);
                        }
                        if (valores.Length == 0)
                        {
                            throw new ArgumentException("El delimitador de los valores debe ser ; o |");
                        }
                    }


                    if (args[3] == null || args[3].ToString().CompareTo(string.Empty) == 0)
                    {
                        throw new ArgumentException("Falta el cuarto argumento: Ruta del documento resultado");
                    }
                    else
                    {
                        resultWord = args[3];
                        pathDestino = System.IO.Path.GetDirectoryName(resultWord);
                    }

                    if (!System.IO.File.Exists(pathWord))
                    {
                        throw new System.IO.FileNotFoundException("No existe la template " + pathWord);
                    }

                    if (!System.IO.Directory.Exists(pathDestino))
                    {
                        throw new System.IO.FileNotFoundException("No existe la ruta donde debería guardar el documento:  " + pathDestino);
                    }

                    if (marcadores.Length != valores.Length)
                    {
                        if (marcadores.Length > valores.Length)
                        {
                            throw new ArgumentException("Hay más marcadores que valores: " + marcadores.Length + " marcadores y " + valores.Length + " valores");
                        }
                        else
                        {
                            throw new ArgumentException("Hay menos marcadores que valores: " + marcadores.Length + " marcadores y " + valores.Length + " valores");
                        }
                    }

                    Console.WriteLine("Rellenando fichero Word...");

                    templateWord = new Microsoft.Office.Interop.Word.Application();
                    documentoWord = templateWord.Documents.Add(pathWord);

                    for (int index = 0; index < marcadores.Length; index++)
                    {
                        if (!buscarCambiarBookmarks(marcadores[index], valores[index]))
                        {
                            continue;
                        }
                    }

                    Console.WriteLine("Guardando fichero Word...");
                    documentoWord.SaveAs(resultWord, 16);

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
                if (documentoWord != null)
                {
                    ((_Document)documentoWord).Close(WdSaveOptions.wdDoNotSaveChanges);
                }

                if(templateWord != null)
                {
                    ((_Application)templateWord).Quit();
                }

                Console.WriteLine("Programa finalizado");
                Environment.ExitCode = result;

            }
        }



        static Boolean buscarCambiarBookmarks(string nombreBookmark, string valor)
        {
            for (int index = 1; index <= documentoWord.Bookmarks.Count; index++)
            {
                if (documentoWord.Bookmarks[index].Name == nombreBookmark)
                {
                    documentoWord.Bookmarks[index].Range.Text = valor;
                    return true;
                }
            }
            return false;
        }


    }
}
