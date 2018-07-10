using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;


namespace ConsoleApp1
{

    class Program
    {
        static void Main(string[] args)
        {

            //string resultado = @"C:\Users\oscarsanchez2\Documents\Automation Anywhere Files\Automation Anywhere\My Tasks\ATT_PRODUCTIVO\IVRMarketing";
            ///

            //string path = AppDomain.CurrentDomain.BaseDirectory;
            //Console.WriteLine("PATH:    " + path);

            //path = path.Before(@"LinqDemoFr\");

            //path += @"\rutasConfigRobot.txt";

            //Console.WriteLine(path);


            //string[] rutasDeEntrada = File.ReadAllLines(path);

            //foreach (var item in rutasDeEntrada)
            //{
            //    //item.Replace("\"", "\\");
            //    //Console.WriteLine(item);
            //}

            ////string pathToExcelFile = @"C:\Users\oscarsanchez2\Documents\AT_T\linq+sql\numerosPruebas.xlsx";

            //string pathToExcelFile = rutasDeEntrada[0];

            //Console.WriteLine("Path To Excel File: " + pathToExcelFile);

            ///

            string path1 = AppDomain.CurrentDomain.BaseDirectory;

            string path3 = Environment.CurrentDirectory.ToString();

            string path2 = Directory.GetCurrentDirectory();

            Console.WriteLine("PATH 1 :    " + path1);

            Console.WriteLine("PATH 2 : " + path2);

            Console.WriteLine("PATH 3 :" + path3);
            ///
            /// La variable path 1 usa el APP DOMAIN. CURRENT DOMAIN
            /// las otras 2 variables se van hasta el system al ser ejecutadas desde el exe, en cambio 
            /// el current domain base directory se queda con toda la ruta
            ///



            string pathALaIniciativa = SubstringExtensions.Before(path1, "Config");

            
            Console.WriteLine("Path a la iniciativa:" + pathALaIniciativa);

            string pathAArchivoRutas = pathALaIniciativa + @"Config\rutasConfigRobot.txt";

            Console.WriteLine("PATH A ARCHIVO RUTAS: " + pathAArchivoRutas);


            string fichero = pathAArchivoRutas;
           
            

            string[] lineas = File.ReadAllLines(fichero);

            foreach (string linea in lineas)
                {
                    Console.WriteLine(linea);
                }






            //string pathToExcelFile = @"C:\Users\oscarsanchez2\Documents\Automation Anywhere Files\Automation Anywhere\My Tasks\ATT_PRODUCTIVO\IVRMarketing\Input\numerosPruebas.xlsx";
            string pathToExcelFile = lineas[0];
            
            

            var excel = new ExcelQueryFactory(pathToExcelFile);
            var query1 = from a in excel.Worksheet<numerosIVRSms>() select a;

            //string fechaComparacion = DateTime.Today.AddDays(-25).ToShortDateString();
            string fechaComparacion = DateTime.Today.AddDays(-1).ToShortDateString();

            foreach (var registro in query1)
            {
                registro.FECHA_ESTATUS = DateTime.Parse(registro.FECHA_ESTATUS).ToShortDateString();
                //Console.WriteLine("MSISDN: " + registro.MSISDN + "\tFECHA: " + registro.FECHA_ESTATUS + "Tipo fecha: " + registro.FECHA_ESTATUS.GetType());
                //Console.WriteLine(registro.FECHA_ESTATUS + " " + registro.FECHA_ESTATUS.GetType() + "   " + fechaComparacion.Equals(registro.FECHA_ESTATUS));

            }
            //Console.ReadLine();
            //Filename.Replace("\"", "\\");
            //Console.WriteLine(query1.Count());

            Console.WriteLine("Fecha de comparacion: " + fechaComparacion + " " + fechaComparacion.GetType());


            var filteredQuery = from a in query1
                                where (a.FECHA_ESTATUS.Contains(fechaComparacion))
                                select a;


            //var filteredSQL = query1.Where(a => a.FECHA_ESTATUS.ToString() == fechaComparacion.ToString());
            //var filteredSQL = query1.Where(a => a.FECHA_ESTATUS.Equals(fechaComparacion)==true).ToList();



            //foreach (var filtered in filteredQuery)
            //{
            //    Console.WriteLine("MDN: " + filtered.MSISDN + "\tFecha: " + filtered.FECHA_ESTATUS);
            //}



            ///Este IF me valida que haya números con la especificacion dada (-1, -5, -9), en caso de no haber numeros que tengan esa especificacion 
            ///no hará nada
            //if (filteredQuery.Count() == 0)
            //{
            //    Console.WriteLine("vacia");
            //}
            //else
            //{
            //    foreach (var filtered in filteredQuery)
            //    {
            //        Console.WriteLine("MDN: " + filtered.MSISDN + "\tFecha: " + filtered.FECHA_ESTATUS);
            //    }

            //}

            //if (DateTime.Today.DayOfWeek == DayOfWeek.Monday)
            //{
            //    Console.WriteLine("Hoy es lunes ¬¬ ");

            //    try
            //    {
            //        StreamWriter sw = new StreamWriter(@"C:\Users\oscarsanchez2\Documents\AT_T\linq+sql\numerosLunes.txt", false);
            //        foreach (var item in filteredQuery)
            //        {
            //            sw.WriteLine(item.MSISDN);
            //        }
            //        sw.Close();
            //    }
            //    catch (Exception e)
            //    {

            //    }
            //}




            string diaActual = DateTime.Today.DayOfWeek.ToString();
            //string diaSiguiente = DateTime.Today.AddDays(1).DayOfWeek.ToString();

            string appendedToFile = "";
            
            //string rutaPrincipalAlArchivo = @"C:\Users\oscarsanchez2\Documents\AT_T\linq+sql\";
            string rutaPrincipalAlArchivo = lineas[1];
         

            switch (diaActual)
            {
                case "Monday":
                    string dTC = DateTime.Today.AddDays(-1).ToShortDateString();

                    var filterMonday = filterQuery(query1, dTC);

                    appendedToFile = diaActual + "FirstMsg_";

                    sendQueryToFile(filterMonday, rutaPrincipalAlArchivo, appendedToFile);

                    //El lunes se manda el de -1 y -2, que serían los del sábado y los del domingo

                    dTC = DateTime.Today.AddDays(-2).ToShortDateString();

                    filterMonday = filterQuery(query1, dTC);

                    appendedToFile = diaActual + "2FirstMsg_";

                    sendQueryToFile(filterMonday, rutaPrincipalAlArchivo, appendedToFile);



                    dTC = DateTime.Today.AddDays(-5).ToShortDateString();

                    filterMonday = filterQuery(query1, dTC);

                    appendedToFile = diaActual + "SecondMsg_";

                    sendQueryToFile(filterMonday, rutaPrincipalAlArchivo, appendedToFile);

                    dTC = DateTime.Today.AddDays(-9).ToShortDateString();

                    filterMonday = filterQuery(query1, dTC);

                    appendedToFile = diaActual + "ThirdMsg_";

                    sendQueryToFile(filterMonday, rutaPrincipalAlArchivo, appendedToFile);


                    break;
                 
                case "Tuesday":
                case "Wednesday":
                case "Thursday":
                    string dateToCompare = DateTime.Today.AddDays(-1).ToShortDateString();

                    var filtroQuery = filterQuery(query1, dateToCompare);

                    appendedToFile = diaActual + "FirstMsg_";

                    sendQueryToFile(filtroQuery, rutaPrincipalAlArchivo, appendedToFile);

                    dateToCompare = DateTime.Today.AddDays(-5).ToShortDateString();

                    filtroQuery = filterQuery(query1, dateToCompare);

                    appendedToFile = diaActual + "SecondMsg_";

                    sendQueryToFile(filtroQuery, rutaPrincipalAlArchivo, appendedToFile);

                    dateToCompare = DateTime.Today.AddDays(-9).ToShortDateString();

                    filtroQuery = filterQuery(query1, dateToCompare);

                    appendedToFile = diaActual + "ThirdMsg_";

                    sendQueryToFile(filtroQuery, rutaPrincipalAlArchivo, appendedToFile);


                    break;
                case "Friday":
                case "Saturday":
                case "Sunday":


                    ///VIERNES
                        string dateToCompareFriday = DateTime.Today.AddDays(-1).ToShortDateString();


                        var filtroQueryFtS = filterQuery(query1, dateToCompareFriday);

                        appendedToFile = "FridayFirstMsg_";

                        sendQueryToFile(filtroQueryFtS, rutaPrincipalAlArchivo, appendedToFile);

                        dateToCompareFriday = DateTime.Today.AddDays(-5).ToShortDateString();

                        filtroQueryFtS = filterQuery(query1, dateToCompareFriday);

                        appendedToFile = "FridaySecondMsg_";

                        sendQueryToFile(filtroQueryFtS, rutaPrincipalAlArchivo, appendedToFile);

                        dateToCompareFriday = DateTime.Today.AddDays(-9).ToShortDateString();

                        filtroQueryFtS = filterQuery(query1, dateToCompareFriday);
                        appendedToFile = "FridayThirdMsg_";

                        sendQueryToFile(filtroQueryFtS, rutaPrincipalAlArchivo, appendedToFile);


                    /// SABADO

                        dateToCompareFriday = DateTime.Today.AddDays(1).AddDays(-1).ToShortDateString();

                        filtroQueryFtS = filterQuery(query1, dateToCompareFriday);

                        appendedToFile = "SaturdayFirstMsg_";

                        sendQueryToFile(filtroQueryFtS, rutaPrincipalAlArchivo, appendedToFile);

                        dateToCompareFriday = DateTime.Today.AddDays(1).AddDays(-5).ToShortDateString();

                        filtroQueryFtS = filterQuery(query1, dateToCompareFriday);

                        appendedToFile = "SaturdaySecondMsg_";

                        sendQueryToFile(filtroQueryFtS, rutaPrincipalAlArchivo, appendedToFile);

                        dateToCompareFriday = DateTime.Today.AddDays(1).AddDays(-9).ToShortDateString();

                        filtroQueryFtS = filterQuery(query1, dateToCompareFriday);

                        appendedToFile = "SaturdayThirdMsg_";

                        sendQueryToFile(filtroQueryFtS, rutaPrincipalAlArchivo, appendedToFile);
                    
                    ///DOMINGO

                        dateToCompareFriday = DateTime.Today.AddDays(2).AddDays(-1).ToShortDateString();

                        filtroQueryFtS = filterQuery(query1, dateToCompareFriday);
            
                        appendedToFile = "SundayFirstMsg_";

                        sendQueryToFile(filtroQueryFtS, rutaPrincipalAlArchivo, appendedToFile);

                        dateToCompareFriday = DateTime.Today.AddDays(2).AddDays(-5).ToShortDateString();

                        filtroQueryFtS = filterQuery(query1, dateToCompareFriday);

                        appendedToFile = "SundaySecondMsg_";

                        sendQueryToFile(filtroQueryFtS, rutaPrincipalAlArchivo, appendedToFile);

                        dateToCompareFriday = DateTime.Today.AddDays(2).AddDays(-9).ToShortDateString();

                        filtroQueryFtS = filterQuery(query1, dateToCompareFriday);
                    
                        appendedToFile = "SundayThirdMsg_";

                        sendQueryToFile(filtroQueryFtS, rutaPrincipalAlArchivo, appendedToFile);

                    break;

                    

                default:
                    break;
            }



        }

        private static IQueryable<numerosIVRSms> filterQuery(IQueryable<numerosIVRSms> query1, string dateToCompare)
        {
            return from a in query1
                   where a.FECHA_ESTATUS.Contains(dateToCompare)
                   select a;
        }

        private static void sendQueryToFile(IQueryable<numerosIVRSms> filteredQuery,string rutaPrincipal, string appendedToFile)
        {
            if (filteredQuery.Count() == 0)
            {
                Console.WriteLine("vacia");
            }
            else
            {
                
                try
                {
                    StreamWriter sw = new StreamWriter(rutaPrincipal+appendedToFile+filteredQuery.Count()+".txt", false);
                    foreach (var item in filteredQuery)
                    {
                        sw.WriteLine(item.MSISDN);
                        //Console.WriteLine(item.MSISDN);
                    }
                    sw.Close();
                }
                catch (Exception e)
                {

                }

            }
        }
    }

    public class ConnexionExcel
    {
        public string _pathExcelFile;
        public ExcelQueryFactory _urlConnexion;
        public ConnexionExcel(string path)
        {
            this._pathExcelFile = path;
            this._urlConnexion = new ExcelQueryFactory(_pathExcelFile);

        }

        public string PathExccelFile => _pathExcelFile;

        public ExcelQueryFactory UrlConnexion => _urlConnexion;

    }

    public class numerosIVRSms
    {
        public string MSISDN { get; set; }
        public string FECHA_ESTATUS { get; set; }
    }

    public class Product
    {
        public int ProductId { get; set; }
        public string ProductName { get; set; }
        public string CategoryName { get; set; }
    }

    static class SubstringExtensions
    {
        /// <summary>
        /// Get string value between [first] a and [last] b.
        /// </summary>
        public static string Between(this string value, string a, string b)
        {
            int posA = value.IndexOf(a);
            int posB = value.LastIndexOf(b);
            if (posA == -1)
            {
                return "";
            }
            if (posB == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= posB)
            {
                return "";
            }
            return value.Substring(adjustedPosA, posB - adjustedPosA);
        }

        /// <summary>
        /// Get string value after [first] a.
        /// </summary>
        public static string Before(this string value, string a)
        {
            int posA = value.IndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            return value.Substring(0, posA);
        }

        /// <summary>
        /// Get string value after [last] a.
        /// </summary>
        public static string After(this string value, string a)
        {
            int posA = value.LastIndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= value.Length)
            {
                return "";
            }
            return value.Substring(adjustedPosA);
        }
    }

    
}
