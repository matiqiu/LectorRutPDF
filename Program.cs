using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Layout;
using iText.Layout.Element;
using System.Text.RegularExpressions;
using SpreadsheetLight;

/**
* Author: matiqiu
*/

Console.WriteLine("\nNOTAS:");
Console.WriteLine("**Los archivos a utilizarse deben estar cerrados.");
Console.WriteLine("**Los 'path' son las rutas de los archivos.");
Console.WriteLine("**Para el 'archivo a crear' se ingresa la ruta de la ubicación \n  donde se ubicará el archivo nuevo más su nombre y formato(PDF).");

Console.WriteLine("\n\n\nIngresa path del archivo PDF a leer:");
var pathExistente = Console.ReadLine();

Console.WriteLine("\nIngresa path del archivo Excel a leer:");
var pathExcel = Console.ReadLine();

Console.WriteLine("\nIngresa path del archivo a crear(PDF):");
var pathCreado = Console.ReadLine();

//Console.WriteLine("\nIngresa path del archivo temporal");
//var pathTemporal = Console.ReadLine();

try
{
    var excel = pathExcel;
    //var excel = "C:\\Users\\matiqiu\\Desktop\\rodrigo\\tottus_btm_202202.xlsx";
    SLDocument sl = new SLDocument(excel);

    int iRow = 2;
    List<string> listaRut = new List<string>();

    while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
    {
        string rut = sl.GetCellValueAsString(iRow, 1);

        listaRut.Add(rut);

        iRow++;
    }

    if (listaRut.Count <= 0)
        throw new Exception("El archivo Excel no contiene ningún rut o ninguno tiene el formato correspondiente. Ej: 11.111.111-1");

    //if (listaRut.Contains("20.132.093-3"))
    //{
    //    Console.WriteLine("bien");
    //}

    // pdf original
    var path = pathExistente;
    //var path = "C:\\Users\\matiqiu\\Desktop\\rodrigo\\consolidado_tottus_202202.pdf";
    PdfReader pdf = new PdfReader(path);
    var pdfDocument = new PdfDocument(pdf);

    // pdf creado
    MemoryStream ms = new MemoryStream();
    PdfWriter pw = new PdfWriter(pathCreado);
    //PdfWriter pw = new PdfWriter("C:\\Users\\matiqiu\\Desktop\\rodrigo\\testPDF.pdf");
    PdfDocument pdfDocument2 = new PdfDocument(pw);
    Document documento = new Document(pdfDocument2);

    //// archivo temporal
    //string linea = String.Empty;
    ////char[] array = new char[] { '8' };
    //string lineaAux = String.Empty;
    //string lineaFinal = String.Empty;
    //StreamReader sr = new StreamReader(pathTemporal);
    ////StreamReader sr = new StreamReader("C:\\Users\\matiqiu\\Desktop\\rodrigo\\testPDF.txt");

    // logica
    string text = String.Empty;
    List<string> listaRutPdf = new List<string>();
    var cantidadPaginas = pdfDocument.GetNumberOfPages();
    for (int i = 1; i <= pdfDocument.GetNumberOfPages(); ++i)
    {
        var page = pdfDocument.GetPage(i);
        text = PdfTextExtractor.GetTextFromPage(page);
        //documento.Add(new Paragraph(text));

        var splitText = text.Split("Rut: ");

        var rut = splitText[1].Substring(0, 12);

        //var rutSeteado = Regex.Replace(rut, @"\s", ""); "[@,\\.\";'\\\\]"
        var rutSeteado = Regex.Replace(rut, @"\,", "");

        //Console.WriteLine(rut);

        if (listaRut.Contains(rutSeteado))
        {
            pdfDocument2.AddPage(page.CopyTo(pdfDocument2));
        }
        listaRutPdf.Add(rutSeteado);
    }
    pdfDocument.Close();
    pdfDocument2.Close();

    Console.WriteLine("\n\n************************************************************");
    Console.WriteLine("\n\nRut's NO encontrados en el PDF:\n");
    foreach (var rut in listaRut)
    {
        if(!listaRutPdf.Contains(rut))
        {
            Console.WriteLine("> " + rut);
        }
    }
    Console.WriteLine("\n\n************************************************************");

    Console.WriteLine("\n\n                                #######                               ");
    Console.WriteLine("                                #######                               ");
    Console.WriteLine("                            tttt##DDDDDtt                             ");
    Console.WriteLine("                            ######;ttti##                             ");
    Console.WriteLine("                            ######titii##                             ");
    Console.WriteLine("                          ##LGGGGG###tttt##                           ");
    Console.WriteLine("                          ##GGGGGGW##iiii##                           ");
    Console.WriteLine("                        ##GLGGGGGLLLG########                         ");
    Console.WriteLine("                        ##GGGGGGGLGLL########                         ");
    Console.WriteLine("                        ##GGGGGGGGGLL##iiitLL##                       ");
    Console.WriteLine("                        ##GGGGGGGGGGL##ttttGL##                       ");
    Console.WriteLine("                      ##tiGGGGGGGGGGGLG####GL##                       ");
    Console.WriteLine("                      ##ttLGGGGGGGGGGGL####LG##                       ");
    Console.WriteLine("                 i######ttLGGG..       LGGG  ##                       ");
    Console.WriteLine("                 t######ttGGGL:.       GGGG  ##                       ");
    Console.WriteLine("             ####Eiiii##ttGG..    #####..##  ##                       ");
    Console.WriteLine("             ####Ettti##ttGG..    #####..##  ##                       ");
    Console.WriteLine("             WWWWDittiW#ttLG..    #####..##  ##      .        .       ");
    Console.WriteLine("           ##itttttttttt##LG..    #####..##  ##      L########        ");
    Console.WriteLine("           ##iittttttttt##LG..    #####..##  ##      f########        ");
    Console.WriteLine("         ##LGttttttttttt##GG....       ..  ..########EGGGGGG######    ");
    Console.WriteLine("         ##GLttttiitttit##GG.:..       ..  ..########EGGLLLL######    ");
    Console.WriteLine("         ##GGGGGLfti##ittt##GG..#########..####tttt##EGLtttttt##tt##  ");
    Console.WriteLine("         ##GGGGGLfti##ittt##GL..#########..####tttt##EGLtttttt##tt##  ");
    Console.WriteLine("         ##GGGGGGE##GL##tttt##...........##ittttttt##EGGLLLLLL##LL##  ");
    Console.WriteLine("         ##GLGGGGE##GL##tttt##:.........:##tttttttt##EGGGGGLLG##GG##  ");
    Console.WriteLine("           ##LGGGE##GLGG##titi###########ttittttttt##EGGGGGG######    ");
    Console.WriteLine("           ##LGGGE##GLLG##tttt###########tttttttttt##EGGGGGG######    ");
    Console.WriteLine("             ##GGLLLGGGG##ttttiiiiiiiiiiiW#########  L########        ");
    Console.WriteLine("             ##LGGLGGLGL##ttttttttttttittW#########  f########        ");
    Console.WriteLine("               W#EGGGLW#tittttttttttttiW#                             ");
    Console.WriteLine("               ##KGGGL##ttttttttttttttt##                             ");
    Console.WriteLine("               ffLEEEEKKffttttttttttiGGff                             ");
    Console.WriteLine("                 t####GGGGtttttttttti##                               ");
    Console.WriteLine("                 t####GGGGtttttttttti##                               ");
    Console.WriteLine("                    ##GGGGGGGGGGGGGGL##                               ");
    Console.WriteLine("                    ##LGGGGGGGGGGGLGL##                               ");
    Console.WriteLine("                 t##ttLGGGGGGGGGGGtti##                               ");
    Console.WriteLine("                 t##ttLGGGGGGGGGGGtti##                               ");
    Console.WriteLine("               ##DitttttGLLGGGGGtttttii##                             ");
    Console.WriteLine("               W#EttttttLGLLLGLGtttttit##                             ");
    Console.WriteLine("             ##GGftttttttt####tttttttGGGG##                           ");
    Console.WriteLine("             ##GGftttttttt####tttttttGGGG##                           ");
    Console.WriteLine("         ####LLGGLGGtttt##    ##LGGGLGGGG##                           ");
    Console.WriteLine("         ####LGGGGGGiitt##    ##LGGGGGGGG##                           ");
    Console.WriteLine("     GGGLEEEELGGGGGGffEEtt    ##LGGGGGLKK##LL                         ");
    Console.WriteLine("     ####LGGGGGGGGGGGL##      ##LGGGGGL######                         ");
    Console.WriteLine("   . WWW#LLGGGGGGGGLGGW#    ..#WLGGGGGL#WWWWW..                       ");
    Console.WriteLine("   ##GGGGLLLLLLLLLLG##      ##GLLGLGGGLGGGGGL##                       ");
    Console.WriteLine("   ##LLLLLLLLLLLLLLG##      ##LLLLLLLLLLLLLLL##                       ");
    Console.WriteLine("   ###################      ###################                       ");
    Console.WriteLine("   ###################      ###################                       \n\n");

    Console.WriteLine("Presiona Enter para cerrar.");
    Console.ReadLine();
}
catch (Exception ex)
{
    Console.WriteLine("\n\n **** ERROR **** \n\n");
    Console.WriteLine("Mensaje de error:  " + ex.Message + "\n\n");
    Console.WriteLine("Presiona Enter para cerrar.");
    Console.ReadLine();

    throw new Exception(ex.Message);
}