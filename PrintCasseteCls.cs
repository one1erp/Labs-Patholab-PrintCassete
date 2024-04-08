using ADODB;
using LSEXT;
using LSSERVICEPROVIDERLib;
using Patholab_Common;
using Patholab_DAL_V1;
using System;
using System.Configuration;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ZXing;
using ZXing.Datamatrix;

namespace PrintCassete
{

    [ComVisible ( true )]
    [ProgId ( "PrintCassete.PrintCasseteCls" )]
    public class PrintCasseteCls : IWorkflowExtension
    {
        INautilusServiceProvider sp;
        private DataLayer dal;
        private AppSettingsSection _appSettings;




        Bitmap  mixImage;
        Bitmap textImage;

        bool save=false;
        public bool DEBUG;
        public void Execute ( ref LSExtensionParameters Parameters )
        {
            try
            {
                long tableID=0;
                string ext = ".1.1";
                string    aliqPathoName = "P_    1-19";
                string    aliqNautilusName = "B000001/19.1.1";
                string pcol="1";
                #region param
                if ( !DEBUG )
                {
                    string tableName = Parameters["TABLE_NAME"];



                    int i = 1;
                    sp = Parameters [ "SERVICE_PROVIDER" ];

                    Recordset rs = Parameters["RECORDS"];

                    rs.MoveLast ( );


                    try
                    {
                        long.TryParse ( rs.Fields [ "ALIQUOT_ID" ].Value.ToString ( ), out tableID );
                    }
                    catch ( Exception ex )
                    {
                        Logger.WriteLogFile ( ex );
                        MessageBox.Show ( "This program works on ALIQUOT only." );
                        return;
                    }
                }
                #endregion

                #region Config

                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                ExeConfigurationFileMap map = new ExeConfigurationFileMap();
                map.ExeConfigFilename = assemblyPath + ".config";
                Configuration cfg = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
                _appSettings = cfg.AppSettings;

                #endregion

                #region Data
                if ( !DEBUG )
                {
                    var ntlCon = Utils.GetNtlsCon(sp);
                    Utils.CreateConstring ( ntlCon );

                    dal = new DataLayer ( );
                    dal.Connect ( ntlCon );

                    var aliq = (from item in dal.GetAll<ALIQUOT>()
                                where item.ALIQUOT_ID == tableID
                                select
                                new
                                {
                                    NautilusName = item.NAME,
                                    PatholabName = item.SAMPLE.SDG.SDG_USER.U_PATHOLAB_NUMBER,
                                    printerCol = item.ALIQUOT_USER.U_PRINTER_COL
                                }).SingleOrDefault();

                    if ( aliq == null )
                    {
                        MessageBox.Show ( "Can't find the aliquot for the id" );
                        return;
                    }


                    int index = 10;
                    aliqNautilusName = aliq.NautilusName;
                    aliqPathoName = CreateShortName ( aliq.NautilusName, aliq.PatholabName );
                    ext = aliq.NautilusName.Substring ( index, aliq.NautilusName.Length - index );
                    pcol = aliq.printerCol;
                }




                #endregion

                #region Print



                Bitmap result;
                result = GenerateBarcode ( aliqNautilusName );
         

                var    barcodeImage = new Bitmap ( result );

                #region תמונה טקסט

                int textSizeWidth = int.Parse(_appSettings.Settings["TextWidth"].Value);
                int textSizeHeight = int.Parse(_appSettings.Settings["TextHeight"].Value);

                //Create bitmap for text
                textImage = new Bitmap ( textSizeWidth, textSizeHeight );

                //First row                             
                PointF TopTextLocation =GetPoint( "Top");
                Font firstFont =GetFont( "Top");

                //Second Row
                PointF bottomTextLocation = GetPoint ( "Bottom" );
                Font secondFont =GetFont( "Bottom");

                //PointF point = new PointF(textPointX, textPointY);

                //Draw
                WriteText ( textImage, aliqPathoName, firstFont, TopTextLocation );
                WriteText ( textImage, ext, secondFont, bottomTextLocation );




                #endregion



                #region תמונה משולבת

                int imageWidth = int.Parse(_appSettings.Settings["ImageWidth"].Value);   //רוחב תמונה
                int ImageHeight = int.Parse(_appSettings.Settings["ImageHeight"].Value); //אורך תמונה//

                mixImage = new Bitmap ( imageWidth, ImageHeight );

                Graphics gmix=Graphics.FromImage(mixImage);


                gmix.DrawImage ( barcodeImage, new PointF ( 0, 0 ) );

                //x,y,width ,height of string
                int TextX     = int.Parse(_appSettings.Settings[ "TextX" ].Value);
                int TextY     = int.Parse(_appSettings.Settings[ "TextY" ].Value);

                gmix.DrawImage ( textImage, TextX, TextY, textImage.Width, textImage.Height );


                gmix.SmoothingMode = SmoothingMode.AntiAlias;
                gmix.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gmix.PixelOffsetMode = PixelOffsetMode.HighQuality;

                mixImage.RotateFlip ( RotateFlipType.Rotate180FlipNone );
        
                gmix.Flush ( );

                //Save the resulting image
                // mixImage.Save ( "C:\\a\\mixImage.png" );
                #endregion

                int p = 1;
                if ( pcol != null )
                    int.TryParse ( pcol, out p );

                PrintDocument pd = new PrintDocument();


                //שרוול
                if ( Environment.MachineName.ToUpper ( ) != "ONE1PC1518" )
                    pd.DefaultPageSettings.PaperSource =
                          pd.PrinterSettings.PaperSources [ p ];


                pd.PrintPage += PrintPage;


                pd.Print ( );

                #endregion


            }
            catch ( Exception ex )
            {
                MessageBox.Show ( "נכשלה הדפסת קסטה." );
                Logger.WriteLogFile ( ex );
            }
            finally
            {
                if ( dal != null ) dal.Close ( );
            }
        }

        private void WriteText ( Bitmap bitmap, string content, Font font, PointF point )
        {
            using ( Graphics graphics = Graphics.FromImage ( bitmap ) )
            {

                graphics.DrawString ( content, font, Brushes.Black, point );

                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphics.Flush ( );

            }
        }

        private Bitmap GenerateBarcode ( string aliqNautilusName )
        {
            int     barcodeX = int.Parse ( _appSettings.Settings [ "barcodeX" ].Value );
            int      barcodeY = int.Parse ( _appSettings.Settings [ "barcodeY" ].Value );
            int     barcodeWidth = int.Parse ( _appSettings.Settings [ "barcodeWidth" ].Value );
            int     barcodeHeight = int.Parse ( _appSettings.Settings [ "barcodeHeight" ].Value );

            DatamatrixEncodingOptions encodingOptions=new DatamatrixEncodingOptions();
            encodingOptions.SymbolShape = ZXing.Datamatrix.Encoder.SymbolShapeHint.FORCE_SQUARE;
            encodingOptions.Width = barcodeWidth;
            encodingOptions.Height = barcodeHeight;
          
            encodingOptions.Margin = 1;
            IBarcodeWriter writer = new BarcodeWriter();
            writer.Format = BarcodeFormat.DATA_MATRIX;
            writer.Options = encodingOptions;
            writer.Options.Height = barcodeHeight;
            writer.Options.Width = barcodeWidth;
            var   result = writer.Write ( aliqNautilusName );

            return result;
        }

        Point GetPoint ( string cfgKey )
        {
            int textPointX = int.Parse(_appSettings.Settings[cfgKey + "TextX"].Value);
            int textPointY = int.Parse(_appSettings.Settings[cfgKey + "TextY"].Value);
            Point pointF = new Point(textPointX, textPointY);
            return pointF;
        }
        Font GetFont ( string cfgKey )
        {
            string textFont = _appSettings.Settings[cfgKey + "TextFont"].Value;
            float textFontSize = float.Parse(_appSettings.Settings[cfgKey + "TextFontSize"].Value);
            string textbold = _appSettings.Settings[cfgKey + "TextBold"].Value;
            FontStyle fontStyle = textbold == "True" ? FontStyle.Bold : FontStyle.Regular;
            Font font = new Font ( textFont, textFontSize,fontStyle );
            return font;
        }

        private string CreateShortName ( string aliqoutName, string sdgPathoName )
        {


            StringBuilder sb = new StringBuilder();
            sb.Append ( sdgPathoName [ 0 ] );
            sb.Append ( sdgPathoName [ 1 ] );

            bool StopRemove = true;
            for ( int i = 2; i < sdgPathoName.Length; i++ )
            {
                if ( StopRemove )
                {


                    if ( sdgPathoName [ i ] == '0' )
                    {
                        sb.Append ( ' ' );
                    }
                    else
                    {
                        StopRemove = false;
                        sb.Append ( sdgPathoName [ i ] );
                    }
                }
                else
                {
                    sb.Append ( sdgPathoName [ i ] );
                }
            }


            var aliqPathoName = sb.ToString();// +ext;
            aliqPathoName = aliqPathoName.Replace ( '/', '-' );


            return aliqPathoName;
        }

        private void PrintPage ( object o, PrintPageEventArgs e )
        {
            int imgPointX = int.Parse(_appSettings.Settings["ImgPointX"].Value);
            int imgPointY = int.Parse(_appSettings.Settings["ImgPointY"].Value);

            Point loc = new Point(imgPointX, imgPointY);
       //     Sharpen ( mixImage );
            e.Graphics.DrawImage ( mixImage, loc );
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            e.Graphics.Flush ( );

        }
   


    }
}
